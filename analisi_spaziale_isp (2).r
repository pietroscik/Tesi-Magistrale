# =============================================================
# ANALISI SPAZIALE ISP 
# =============================================================
# -------------------------------------------------------------
#  CARTELLE CREATE (se non esistono):
#  ├── 01_rds/      (oggetti RDS, GPKG)
#  ├── 02_maps/     (png, pdf)
#  ├── 03_tables/   (csv, txt, md)
#  ├── 04_html/     (leaflet – htmlwidgets)
#  └── 05_logs/     (output sink / diagnostica)
# -------------------------------------------------------------
#  Tutte le chiamate a ggsave(), saveWidget(), write.csv(), st_write()
#  sono state sostituite con funzioni wrapper che indirizzano nel folder
#  corretto.
# LIBRERIE #####
libs <- c("car","classInt","dplyr","ggplot2","kableExtra","leaflet","lmtest",
          "maptools","mapview","moments","nortest","patchwork","RColorBrewer",
          "readxl","rcompanion","sf","spData","spdep","spatialreg","stargazer",
          "tripack","FNN","future","future.apply","GWmodel","spgwr",
          "htmlwidgets","viridis","tmap")

invisible(lapply(libs, require, character.only = TRUE))

# 2. FUNZIONI DI SALVATAGGIO E LOG
root_out <- getwd()                      # directory di lavoro corrente
subdirs  <- c("01_rds","02_maps","03_tables","04_html","05_logs")
for(dir in file.path(root_out, subdirs)) if(!dir.exists(dir)) dir.create(dir)

save_plot <- function(plot, name, w=8, h=6, dpi=300, ext="png")
  ggsave(file.path("02_maps", paste0(name, ".", ext)), plot=plot, width=w, height=h, dpi=dpi)

save_pdf  <- function(plot, name, w=8, h=6)
  ggsave(file.path("02_maps", paste0(name, ".pdf")), plot=plot, width=w, height=h)

save_html <- function(widget, name)
  htmlwidgets::saveWidget(widget, file.path("04_html", paste0(name, ".html")), selfcontained=TRUE)

save_rds  <- function(obj, name) saveRDS(obj, file.path("01_rds", paste0(name, ".rds")))

save_table <- function(df, name, row.names=FALSE)
  write.csv(df, file.path("03_tables", paste0(name, ".csv")), row.names=row.names)

open_log  <- function(name) sink(file.path("05_logs", paste0(name, ".txt")), split=TRUE)
close_log <- function() sink(NULL)

# FUNZIONI DI PLOT ‘CHIARO’ #    (LISA, Gi*, coefficienti GWR)####

pal_div <- function(n)scales::div_gradient_pal(low="#2166ac", mid="white", high="#b2182b")(seq(0,1,len=n))

plot_lisa <- function(sf_obj,
                      ripgeo        = ripgeo,
                      titolo        = "Cluster LISA",
                      size_ns       = .6,        # punti non-signif.
                      size_sig      = .9,        # punti significativi
                      alpha_ns      = .35,       # trasparenza grigi
                      alpha_sig     = .90){
  
  # assicuro l’ordine dei livelli
  sf_obj$LISA_cluster <- factor(sf_obj$LISA_cluster,
                                levels = c("Non signif",
                                           "High-High","Low-Low",
                                           "High-Low","Low-High"))
  
  ggplot() +
    # basemap
    geom_sf(data = ripgeo,
            fill   = "grey95",
            colour = "grey60",
            linewidth = .12) +
    
    # 1) layer grigio (non significativi) – disegnato per primo
    geom_sf(data = sf_obj |> dplyr::filter(LISA_cluster == "Non signif"),
            colour      = "grey70",
            size        = size_ns,
            alpha       = alpha_ns,
            show.legend = FALSE) +
    
    # 2) layer dei cluster significativi – sopra quello grigio
    geom_sf(data = sf_obj |> dplyr::filter(LISA_cluster != "Non signif"),
            aes(colour  = LISA_cluster),
            size        = size_sig,
            alpha       = alpha_sig,
            show.legend = "point") +
    
    scale_colour_manual(
      values = c("High-High" = "#ca0020",
                 "Low-Low"   = "#0571b0",
                 "High-Low"  = "#fdae61",
                 "Low-High"  = "#92c5de")) +
    
    guides(colour = guide_legend(override.aes = list(size = 4))) +
    
    labs(title = titolo, colour = "Cluster") +
    
    theme_minimal(base_size = 11) +
    theme(panel.grid.major = element_blank(),
          panel.grid.minor = element_blank())
}


plot_gi <- function(sf_obj,
                    ripgeo   = ripgeo,
                    titolo   = "Hotspot (Gi*)",
                    size_sig = .9,
                    size_ns  = .4,
                    alpha_ns = .35){
  
  # ricavo / ricreo la classificazione in bande
  if(!"Gi_bin" %in% names(sf_obj)){
    sf_obj$Gi_bin <- cut(sf_obj$Gi,
                         breaks  = c(-Inf,-1.96,-1.65,1.65,1.96,Inf),
                         labels  = c("Coldspot 99%","Coldspot 95%","Non signif",
                                     "Hotspot 95%","Hotspot 99%"))
  }
  
  sf_ns  <- sf_obj |> dplyr::filter(Gi_bin == "Non signif")
  sf_sig <- sf_obj |> dplyr::filter(Gi_bin != "Non signif")
  
  ggplot() +
    geom_sf(data = ripgeo, fill = "grey95", colour = "grey60", linewidth = .15) +
    
    geom_sf(data = sf_ns,
            fill  = "grey80", colour = "black",
            shape = 21, size = size_ns, alpha = alpha_ns,
            show.legend = FALSE) +
    
    geom_sf(data = sf_sig,
            aes(fill = Gi_bin),
            shape = 21, colour = "black",
            size = size_sig, alpha = .9,
            show.legend = "polygon") +
    
    scale_fill_manual(values = c(
      "Coldspot 99%" = "#08306B",
      "Coldspot 95%" = "#4292C6",
      "Hotspot 95%"  = "#FC4E2A",
      "Hotspot 99%"  = "#99000D"
    )) +
    guides(fill = guide_legend(override.aes = list(size = 4))) +
    labs(title = titolo, fill = "Gi*") +
    theme_minimal(base_size = 11) +
    theme(panel.grid.major = element_blank(),
          panel.grid.minor = element_blank())
}

# Funzione rivista per plottare i coefficienti GWR ####
plot_gwr_coef <- function(sf_obj,coef_name,ripgeo,title= paste("Coeff. GWR:", coef_name),
                          pal_low    = "#0571b0",
                          pal_mid    = "white",
                          pal_high   = "#ca0020",
                          limits     = NULL,
                          pt_size    = .35,
                          legend     = TRUE) {
  
  # --- controlli ------------------------------------------------------
  if (!coef_name %in% names(sf_obj))
    stop("Colonna '", coef_name, "' non presente in sf_obj")
  if (!is.numeric(sf_obj[[coef_name]]))
    stop("La colonna '", coef_name, "' deve essere numerica")
  
  v <- sf_obj[[coef_name]]
  if (is.null(limits)) {
    max_abs <- max(abs(v), na.rm = TRUE)
    limits  <- c(-max_abs, max_abs)
  }
  
  # --- plot -----------------------------------------------------------
  ggplot() +
    geom_sf(data = ripgeo,
            fill   = "grey95",
            colour = "grey60",
            linewidth = .15) +
    
    geom_sf(data = sf_obj,
            aes(colour = .data[[coef_name]]),
            size  = pt_size,
            alpha = .85,
            show.legend = if (legend) "point" else FALSE) +
    
    scale_colour_gradient2(low       = pal_low,
                           mid       = pal_mid,
                           high      = pal_high,
                           midpoint  = 0,
                           limits    = limits,
                           name      = "Coef") +
    
    theme_minimal(base_size = 11) +
    labs(title = title) +
    theme(panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          legend.position  = if (legend) "right" else "none")
}

# 2.b FUNZIONI ANALITICHE 
# -- 2.b.1  k ottimale (versione «combined‑metric» identica a quella ####
scegli_k_knn <- function(imprese_cat,k_min = 5,k_max = 120,step  = 5,dist_fun = median,verbose  = FALSE) {
  if(verbose) message("[scegli_k_knn] range k = ", k_min, "-", k_max)
  
  coords <- sf::st_coordinates(imprese_cat)
  var_x  <- imprese_cat$ISP_bn
  k_seq  <- seq(k_min, k_max, by = step)
  
    res <- lapply(k_seq, function(k){
    nb <- spdep::knn2nb(spdep::knearneigh(coords, k = k))
    lw <- spdep::nb2listw(nb, style = "W", zero.policy = TRUE)
    mt <- spdep::moran.test(var_x, lw, zero.policy = TRUE)
    
    # distanza all’ennesimo vicino (colonna k della matrice distanze)
    kd <- FNN::get.knn(coords, k = k)$nn.dist[, k]
    
    data.frame(k = k,
               moran_i = mt$estimate[["Moran I statistic"]],
               p_value = mt$p.value,
               dist_k  = dist_fun(kd))
  }) |> dplyr::bind_rows()
  
  res <- dplyr::arrange(res, k)
  res <- dplyr::mutate(res,
                       d_moran = c(NA, abs(diff(moran_i))),
                       d_dist  = c(NA, abs(diff(dist_k))))
  # min-max scaling (ignora NA/Inf)
  sc <- function(x){
    out        <- rep(NA_real_, length(x))
    idx        <- is.finite(x)
    if(sum(idx) > 1){
      rng      <- range(x[idx])
      out[idx] <- (x[idx] - rng[1]) / diff(rng)
    }
    out
  }
  res$sc_moran <- sc(res$d_moran)
  res$sc_dist  <- sc(res$d_dist)
  res$combo    <- res$sc_moran + res$sc_dist
  
  # k ottimale = minimo combo  (escludo prima riga perché d_* = NA)
  k_best <- res$k[which.min(res$combo)]
  if(verbose) message("[scegli_k_knn]  k* = ", k_best)
  
  list(best_k = k_best, risultati = res)
}

# -- 2.b.2  Analisi spaziale punti --------------------------------------
analizza_spaziale_punti <- function(imprese_cat, ripgeo, dimensione = NULL) {
  message("[analizza_spaziale_punti] dim = ", dimensione)
  n0 <- nrow(imprese_cat)
  imprese_cat <- dplyr::filter(imprese_cat, !is.na(ISP_bn))
  if(nrow(imprese_cat) < n0)
    message("   – rimossi ", n0 - nrow(imprese_cat), " NA in ISP_bn")
  
  # jitter per sovrapposizioni
  imprese_cat$geometry <- sf::st_jitter(imprese_cat$geometry, amount = 0.0001)
  
  k_res <- scegli_k_knn(imprese_cat, 10, 30)
  if (is.null(k_res)) stop("k non valido per ", dimensione)
  k_sel <- k_res$best_k
  print(k_sel)
  nb <- spdep::knn2nb(spdep::knearneigh(sf::st_coordinates(imprese_cat), k = k_sel))
  lw <- spdep::nb2listw(nb, style = "W", zero.policy = TRUE)
  
  # globali
  mor <- spdep::moran.test(imprese_cat$ISP_bn, lw, zero.policy = TRUE)
  gea <- spdep::geary.test(imprese_cat$ISP_bn, lw, zero.policy = TRUE)
  
  # ---- LISA -----------------------------------------------------------
  lisa <- spdep::localmoran(imprese_cat$ISP_bn, lw)
  lag_ISP <- spdep::lag.listw(lw, imprese_cat$ISP_bn)
  
  imp <- imprese_cat %>%                       # copia dell’oggetto sf
    dplyr::mutate(
      lisa_p  = lisa[, 5],
      lisa_I  = lisa[, 1],
      ISP_std     = as.numeric(scale(ISP_bn)),
      lag_ISP_std = as.numeric(scale(lag_ISP)),
      LISA_cluster = dplyr::case_when(
        lisa_p <= .05 & ISP_std > 0 & lag_ISP_std > 0 ~ "High-High",
        lisa_p <= .05 & ISP_std < 0 & lag_ISP_std < 0 ~ "Low-Low",
        lisa_p <= .05 & ISP_std > 0 & lag_ISP_std < 0 ~ "High-Low",
        lisa_p <= .05 & ISP_std < 0 & lag_ISP_std > 0 ~ "Low-High",
        TRUE                                          ~ "Non signif"
      ) %>% factor(levels = c("High-High","Low-Low",
                              "High-Low","Low-High","Non signif"))
    )
  
  map_lisa <- ggplot() +
    geom_sf(data = ripgeo, fill = "grey90", color = "black", linewidth = .05) +
    geom_sf(data = imp, aes(color = LISA_cluster), size = .25) +
    scale_color_manual(values = c("High-High" = "#ca0020",
                                  "Low-Low"  = "#0571b0",
                                  "High-Low" = "#fdae61",
                                  "Low-High" = "#92c5de",
                                  "Non signif" = "grey70")) +
    labs(title = paste0("LISA – ", dimensione), color = "Cluster") +
    theme_minimal(base_size = 10) +
    theme(panel.grid = element_blank())
  # Gi*--------------------------------------------------------------
  Gi <- spdep::localG(imp$ISP_bn, lw)
  imp$Gi_bin <- cut(as.numeric(Gi),
                    breaks = c(-Inf, -1.96, -1.65, 1.65, 1.96, Inf),
                    labels = c("Coldspot 99%","Coldspot 95%","Non signif",
                               "Hotspot 95%","Hotspot 99%"))
  
  map_gi <- ggplot() +
    geom_sf(data = ripgeo, fill = "grey90", colour = "black", linewidth = .05) +
    geom_sf(data = imp,
            aes(fill = Gi_bin),
            shape = 21,            # cerchio con riempimento
            colour = NA,           # ← niente bordo
            size   = 0.6) +        # punto un po’ più grande
    scale_fill_manual(values = c(
      "Coldspot 99%" = "#08306B",
      "Coldspot 95%" = "#4292C6",
      "Non signif"   = "grey75",
      "Hotspot 95%"  = "#FC4E2A",
      "Hotspot 99%"  = "#99000D"
    )) +
    labs(title = paste0("Gi* – ", dimensione),
         fill  = "Gi* bin") +
    theme_minimal(base_size = 10) +
    theme(panel.grid = element_blank())
  
  
  list(imprese_cat   = imp,
       W_knn         = lw,
       k_opt         = k_sel,
       lisa_map      = map_lisa,
       gi_map        = map_gi,
       moran_global  = mor,
       geary_global  = gea)
}

# -- 2.b.3  ANALISI REGRESSIVA SPAZIALE -------------------------------
analisi_regressiva_spaziale <- function(sf_obj, listw, vars_model) {
  # formula
  fml <- as.formula(paste("ISP_bn ~", paste(vars_model, collapse = " + ")))
  
  # SAR (Error)
  sar <- spatialreg::errorsarlm(fml, data = sf_obj, listw = listw,
                                zero.policy = TRUE)
  
  # SDM (mixed Durbin)
  sdm <- spatialreg::lagsarlm(fml, data = sf_obj, listw = listw,
                              type = "mixed", zero.policy = TRUE)
  
  # Impacts (può richiedere tempo)
  imp <- tryCatch(spatialreg::impacts(sdm, listw = listw, R = 500),
                  error = function(e) NULL)
  
  # GMM (spatial error – two‑step)
  gmm <- spatialreg::GMerrorsar(fml, data = sf_obj, listw = listw,
                                zero.policy = TRUE)
  
  list(sar = sar, sdm = sdm, impacts = imp, gmm = gmm)
}
# 2.b.4  GWR (locale)  –  usa spgwr::gwr  +  subsample facoltativo
run_gwr <- function(sf_obj,vars_model,listw,samp_size = 5000,seed= 2025) {
  set.seed(seed)
  if (nrow(sf_obj) > samp_size)
    sf_obj <- sf_obj[sample(seq_len(nrow(sf_obj)), samp_size), ]
  
  # formula
  fml <- as.formula(paste("ISP_bn ~", paste(vars_model, collapse = " + ")))
  coords  <- sf::st_coordinates(sf_obj)
  
  # bandwidth ottimale (spgwr)  -----------------------
  bw <- spgwr::gwr.sel(fml, data = sf_obj, coords = coords,
                       adapt = FALSE, gweight = spgwr::gwr.Gauss,
                       longlat = FALSE)
  
  message("[run_gwr]  bandwidth = ", round(bw, 2), "  (", nrow(sf_obj), " obs)")
  
  # modello GWR  --------------------------------------
  gwr_mod <- spgwr::gwr(fml, data = sf_obj, coords = coords, bandwidth = bw,
                        hatmatrix = TRUE, gweight = spgwr::gwr.Gauss)
  
  coeff_sf <- cbind(sf_obj, as.data.frame(gwr_mod$SDF))
  list(model = gwr_mod, coeff_sf = coeff_sf, bw = bw)
}

# -- 2.b.4  Plot rapido di un coefficiente locale GWR ####
plot_gwr_coef <- function(coeff_sf,var_name,ripgeo,titolo = NULL,legend = FALSE,size_sf = 1) {
  p <- ggplot() +    geom_sf(data = ripgeo,fill   = "grey90",colour = "black",
            linewidth = .05) +geom_sf(data = coeff_sf,aes_string(colour = var_name),
            size = size_sf) +scale_colour_viridis_c(option = "viridis") +
    labs(title  = titolo %||% paste("GWR –", var_name),colour = paste("Coef", var_name)) +
    theme_minimal(base_size = 10) +theme(panel.grid = element_blank())
  
  if (!legend)
    p <- p + theme(legend.position = "none")
  
  p
}

# 3. BLOCCO ORIGINALE CON MODIFICHE DI SALVATAGGIO
# 3.1 Definizione variabili e scaling ---------------------------------
variabili_redditivita <- c("ROE_percentuale", "ROI_percentuale", "ROS_percentuale", 
                           "ROA_percentuale", "EBITDA_migl_EUR", "utile_netto_migl_EUR")
variabili_solidita <- c("patrimonio_netto_migl_EUR", "posizione_finanziaria_netta_migl_EUR", 
                        "debt_equity_ratio_percentuale", "debt_EBITDA_ratio_percentuale")
variabili_produttivita <- c("dipendenti", "ricavi_pro_capite_EUR", "valore_aggiunto_pro_capite_EUR", 
                            "rendimento_dipendenti")
variabili_liquidita <- c("indice_liquidita", "indice_corrente", "flusso_cassa_gestione_migl_EUR")
variabili_circolante <- c("giacenza_media_scorte_gg", "durata_media_crediti_lordo_IVA_gg", 
                          "durata_media_debiti_lordo_IVA_gg", "incidenza_circolante_operativo_percentuale")

categorie_variabili <- list(
  Redditività        = variabili_redditivita,
  Solidità           = variabili_solidita,
  Produttività       = variabili_produttivita,
  Liquidità          = variabili_liquidita,
  Capitale_Circolante= variabili_circolante
)
categorie_variabili$Performance <- c("ISP_bn","integrazione_verticale")
vars_model <- unname(unlist(categorie_variabili))
vars_model <- setdiff(vars_model, "ISP_bn")                 #  --  via la risposta

# 3.2 Import dati e creazione sf ---------------------------------------

dataset_completo <- readRDS("matrice.rds")

extra_cols <- c("EBITDA_su_vendite_percentuale","rotazione_cap_investito",
                "totale_attivita_migl_EUR","PFN_EBITDA_migl_EUR",
                "indice_indipendenza_finanziaria_percentuale",
                "grado_copertura_interessi_percentuale","indice_indebitamento_breve_percentuale",
                "rapporto_indebitamento","oneri_finanziari_fatt_percentuale")

dataset_z_score <- as.data.frame(scale(dataset_completo[c(vars_model, extra_cols)]))
dataset_z_score$ISP_bn <- scale(dataset_completo$ISP_bn)   # mantiene il valore originale

cols_keep <- c("codice_fiscale","sezione","dimensione_impresa","longitudine","latitudine")
for(cl in cols_keep) dataset_z_score[[cl]] <- dataset_completo[[cl]]

imprese <- dataset_z_score

# filtri dimensione
imprese_G <- filter(imprese, dimensione_impresa == "Grande")
imprese_M <- filter(imprese, dimensione_impresa == "Media")
imprese_P <- filter(imprese, dimensione_impresa == "Piccola")
imprese_m <- filter(imprese, dimensione_impresa == "Micro")

# sf conversione
imprese_sf     <- st_as_sf(imprese,   coords = c("longitudine","latitudine"), crs = 4326)
imprese_G_sf   <- st_as_sf(imprese_G, coords = c("longitudine","latitudine"), crs = 4326)
imprese_M_sf   <- st_as_sf(imprese_M, coords = c("longitudine","latitudine"), crs = 4326)
imprese_P_sf   <- st_as_sf(imprese_P, coords = c("longitudine","latitudine"), crs = 4326)
imprese_m_sf   <- st_as_sf(imprese_m, coords = c("longitudine","latitudine"), crs = 4326)

# trasformazione CRS metrico
for(nm in ls(pattern="^imprese_[A-Za-z]*_sf$|^imprese_sf$") )
  assign(paste0(nm,"_tr"), st_transform(get(nm), 32632))

# join con confini
comuni <- st_read("Com01012025_WGS84.shp", quiet=TRUE)
ripgeo <- st_read("RipGeo01012025_WGS84.shp", quiet=TRUE)

for(nm in ls(pattern="_sf_tr$") )
  assign(nm, st_join(get(nm), comuni, join = st_within))

non_joined <- get("imprese_sf_tr")[is.na(get("imprese_sf_tr")$Shape_Area), ]
if(nrow(non_joined)>0){
  for(nm in ls(pattern="_sf_tr$") )
    assign(nm, filter(get(nm), !codice_fiscale %in% non_joined$codice_fiscale))
}

# salvataggio rds
for(nm in ls(pattern="_sf_tr$") ) save_rds(get(nm), nm)

# 3.3 Mappe iniziali ----------------------------------------------------

# funzione di mappa rapida – punti chiari e legenda opzionale ####
plot_map <- function(data,ripgeo_layer = NULL,pt_col = "#2196F3",pt_shape = 21,
                     pt_size  = 0.7, titolo   = "",show_legend = FALSE) {
    if (is.null(ripgeo_layer)) ripgeo_layer <- get("ripgeo", envir = .GlobalEnv)
  ggplot() +geom_sf(data = ripgeo_layer,fill = "grey95", colour = "grey60", linewidth = .15) +
    geom_sf(data = data,fill = pt_col, colour = "black",shape = pt_shape, size = pt_size, alpha = .85,
            show.legend = show_legend) +theme_minimal(base_size = 11) +
    labs(title = titolo) +theme(panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),legend.position  = if (show_legend) "right" else "none")
}

p0 <- ggplot() +  geom_sf(data = ripgeo,fill  = "grey90",colour= "grey50",
          linewidth = .15) +geom_sf(data = imprese_sf_tr, aes(colour = dimensione_impresa),
          size  = .10,alpha = .8) +scale_colour_viridis_d(
    option = "D",    name   = "Dimensione\nimpresa") +
  guides(colour = guide_legend(override.aes = list(size = 3))) +
  coord_sf(crs = st_crs(imprese_sf_tr)) +
  theme_minimal(base_size = 11) +
  labs(title = "Imprese italiane per classe dimensionale") +
  theme(panel.grid.major = element_blank(), panel.grid.minor = element_blank())


p1 <- plot_map(imprese_G_sf_tr, pt_col = "blue", titolo = "Grandi Imprese")
p2 <- plot_map(imprese_M_sf_tr, pt_col = "red",    titolo = "Medie Imprese")
p3 <- plot_map(imprese_P_sf_tr, pt_col = "green",  titolo = "Piccole Imprese")
p4 <- plot_map(imprese_m_sf_tr, pt_col = "purple", titolo = "Micro Imprese")

grid <- (p1+p2)/(p3+p4) + patchwork::plot_layout(guides="collect")

save_pdf(p0,"mappa_imprese")
save_pdf(p1,"mappa_grandi_imprese")
save_pdf(p2,"mappa_medie_imprese")
save_pdf(p3,"mappa_piccole_imprese")
save_pdf(p4,"mappa_micro_imprese")
save_pdf(grid,"mappa_imprese_grid")

# 3.4 Analisi K ottimale & strutture di vicinanza ####

open_log("output_analisi_spaziali_k")

coords <- st_coordinates(imprese_sf_tr)
coords_df <- data.frame(coords, index=seq_len(nrow(coords)))

# jitter duplicati
coords_dup <- coords_df %>% count(X,Y) %>% filter(n>1)
coords_df <- coords_df %>% left_join(coords_dup, by=c("X","Y")) %>%
  mutate(X_new = ifelse(!is.na(n), jitter(X,1e-4), X),
         Y_new = ifelse(!is.na(n), jitter(Y,1e-4), Y))

imprese_sf_jittered <- imprese_sf_tr
st_geometry(imprese_sf_jittered) <- st_sfc(
  mapply(
    function(x, y) st_point(c(x, y)),
    coords_df$X_new,
    coords_df$Y_new,
    SIMPLIFY = FALSE
  ),
  crs = st_crs(imprese_sf_tr)
)

coords_jit <- st_coordinates(imprese_sf_jittered)

max_k <- 250
knn_res <- FNN::get.knn(coords_jit, k=max_k)

dist_median <- apply(knn_res$nn.dist,2,median)
dist_mean   <- apply(knn_res$nn.dist,2,mean)
dist_sd     <- apply(knn_res$nn.dist,2,sd)

plot_df <- data.frame(k=1:max_k, mean=dist_mean, median=dist_median, sd=dist_sd)

p_kdist <- ggplot(plot_df)+
  geom_line(aes(k,mean), color="blue", size=.6)+
  geom_line(aes(k,median), color="darkgreen", linetype=2)+
  geom_line(aes(k,sd), color="red", linetype=3)+
  theme_minimal()+labs(title="Distanza ai k‑vicini", y="Distanza", x="k")

save_pdf(p_kdist,"plot_k_distanza")

# elbow
elbow_k <- which.max(abs(diff(dist_median, differences=2)))+1
cat("\nPunto di flesso stimato a k =", elbow_k,"\n")

# validazione moran ----------------------------------------------------
validazione_moransI_k_parallel <- function(sf_data, var_name="ISP_bn", k_seq=seq(5,200,5), workers=4){
  future::plan(multisession, workers=workers)
  coords <- st_coordinates(sf_data)
  varvec <- sf_data[[var_name]]
  future.apply::future_lapply(k_seq, function(k){
    nb  <- spdep::knn2nb(spdep::knearneigh(coords,k=k))
    lw  <- spdep::nb2listw(nb,style="W",zero.policy=TRUE)
    mt  <- spdep::moran.test(varvec, lw, zero.policy=TRUE)
    data.frame(k=k, moran_i=mt$estimate[["Moran I statistic"]], p=mt$p.value)
  },
  future.seed = TRUE      # ← aggiungi questo
  ) |>
    do.call(what = rbind)
}
risultati_k <- validazione_moransI_k_parallel(imprese_sf_jittered)

save_table(risultati_k,"k_moran")

p_kmor <- ggplot(risultati_k, aes(k, moran_i, color=p<.05))+geom_line()+geom_point()+
          scale_color_manual(values=c("TRUE"="green","FALSE"="red"))+theme_minimal()+
          labs(title="Moran's I vs k", color="Significativo")

save_pdf(p_kmor,"plot_k_moran")

# metrica combinata dist+delta-moran
risultati_k$dist_median <- dist_median[risultati_k$k]
risultati_k <- risultati_k |> mutate(delta_moran = c(NA,diff(moran_i)),
                                     delta_dist  = c(NA,diff(dist_median[risultati_k$k])))

scale01 <- function(x) (x-min(x,na.rm=TRUE))/(max(x,na.rm=TRUE)-min(x,na.rm=TRUE))
risultati_k <- risultati_k |> mutate(score = scale01(abs(delta_moran))+scale01(abs(delta_dist)))

k_opt <- risultati_k$k[which.min(risultati_k$score)]
cat("\nk ottimale selezionato:", k_opt,"\n")

close_log()

# 3.5 Creazione lista vicinanza principale -----------------------------

knn_list <- spdep::knearneigh(coords_jit, k=k_opt)
knn_nb   <- spdep::knn2nb(knn_list)
listw_knn<- spdep::nb2listw(knn_nb, style="W", zero.policy=TRUE)
W_knn    <- spdep::nb2mat(knn_nb, style="W", zero.policy=TRUE)

save_rds(listw_knn,"listw_knn")

#  plot struttura KNN – versione super-compatta  (nessun cex= in points) ####
pdf(file.path("02_maps", "plot_knn_vicinanza.pdf"),
    width = 9, height = 7, pointsize = 6)

par(mar = c(0, 0, 1, 0),
    cex = 0.08)              # <-- tutti i simboli/etichette ≈ 8 % del default

plot(st_geometry(ripgeo),
     col    = "grey95",
     border = "grey70",
     lwd    = .3,
     main   = sprintf("Struttura vicinanza KNN  (k = %d)", k_opt))

plot(knn_nb, coords = coords_jit,
     add = TRUE,
     col = adjustcolor("#d73027", .55),
     lwd = .10)

points(coords_jit, pch = 20) # nessun “cex=” – eredita il 0.08 impostato sopra

legend("bottomleft",
       legend = sprintf("k = %d  –  n = %s", k_opt, format(nrow(coords_jit), big.mark=",")),
       bty = "n", text.col = "grey30", cex = .9)

dev.off()

# 3.6 Moran, Geary, LISA, Gi* -----------------------------------------

open_log("output_autocorrelazione_globale")
var_int <- imprese_sf_jittered$ISP_bn

mt  <- spdep::moran.test(var_int, listw_knn, zero.policy=TRUE)
print(mt)

ge  <- spdep::geary.test(var_int, listw_knn, zero.policy=TRUE)
print(ge)

close_log()

# scatter –versione PDF più pulita ---------------------------------
pdf(file.path("02_maps","moran_scatterplot_knn.pdf"),
    width = 8, height = 6)           # ≈ 2400×1800 px
par(mar = c(4, 4, 2, 1))             # margini standard
spdep::moran.plot(
  var_int,
  listw_knn,
  zero.policy = TRUE,
  pch   = 20,
  col   = "#3182bd",   # blu tenue
  cex = .2,           # puntini piccoli
  labels = FALSE        # nessuna etichetta ID
)
# linee asse
abline(h = 0, v = 0, lty = 2)
dev.off()

# LISA
lisa   <- spdep::localmoran(var_int, listw_knn)
Ii     <- lisa[,"Ii"]; pval <- lisa[,"Pr(z != E(Ii))"]
alpha  <- .05
x_mean <- mean(var_int)
cluster <- rep("Non signif", length(var_int))
cluster[Ii>0 & var_int>x_mean & pval<alpha] <- "High-High"
cluster[Ii>0 & var_int<x_mean & pval<alpha] <- "Low-Low"
cluster[Ii<0 & var_int>x_mean & pval<alpha] <- "High-Low"
cluster[Ii<0 & var_int<x_mean & pval<alpha] <- "Low-High"

imprese_sf_jittered$LISA_cluster <- factor(cluster,
  levels=c("High-High","Low-Low","High-Low","Low-High","Non signif"))

p_lisa <- plot_lisa(imprese_sf_jittered, ripgeo)
save_plot(p_lisa, "lisa_cluster_map")

# Gi*
Gi <- spdep::localG(var_int, listw_knn, zero.policy=TRUE)
imprese_sf_jittered$Gi <- as.numeric(Gi)

cutGi <- cut(imprese_sf_jittered$Gi, breaks=c(-Inf,-1.96,-1.65,1.65,1.96,Inf),
             labels=c("Coldspot 99%","Coldspot 95%","Non signif","Hotspot 95%","Hotspot 99%"))

p_gi <- plot_gi(imprese_sf_jittered, ripgeo)
save_pdf(p_gi, "gi_star_hotspot")

# 3.7 Modelli Regressivi ------------------------------------------------

open_log("output_modelli_spaziali")

set.seed(2025)
nsample <- min(5000, nrow(imprese_sf_jittered))
imp_sample <- imprese_sf_jittered[sample(seq_len(nrow(imprese_sf_jittered)), nsample), ]
coords_sample <- st_coordinates(imp_sample)
nb_s  <- spdep::knn2nb(spdep::knearneigh(coords_sample, k=k_opt))
lw_s  <- spdep::nb2listw(nb_s, style="W",zero.policy=TRUE)

fml <- as.formula(paste0("ISP_bn ~ ", paste(vars_model, collapse=" + "))) 

sar <- spatialreg::errorsarlm(fml, data=imp_sample, listw=lw_s, zero.policy=TRUE)
print(summary(sar))
print(spdep::moran.test(residuals(sar), lw_s, zero.policy=TRUE))

sdm <- spatialreg::lagsarlm(fml,data= imp_sample,listw = lw_s,type = "mixed",zero.policy = TRUE)

print(summary(sdm))

imp_sdm <- spatialreg::impacts(sdm, listw = lw_s, R = 1000)
print(imp_sdm)              # oppure summary(imp_sdm)
print(spdep::moran.test(residuals(sdm), lw_s, zero.policy=TRUE))

gm  <- spatialreg::GMerrorsar(fml,data=imprese_sf_jittered,listw=listw_knn,zero.policy=TRUE)
print(summary(gm))

close_log()

# GWR – bandwidth e plotting coeff ------------------------------------------------

bw <- spgwr::gwr.sel(fml, data = imp_sample, coords = coords_sample)
cat("Bandwidth ottimale GWR:", bw, "\n")

gwr_mod <- spgwr::gwr(fml,data   = imp_sample,coords = coords_sample,bandwidth = bw,hatmatrix = TRUE)
coeff_loc <- as.data.frame(gwr_mod$SDF)
df_gwr <- cbind(imp_sample, coeff_loc)

# singoli pdf GWR
# for(v in vars_model){
#   p <- plot_gwr_coef(df_gwr, v, ripgeo)
#   save_pdf(p, paste0("grafico_GWR_", v))
# }

# pdf unico Gwr
pdf(file.path("02_maps", "gwr_coefficients_collection.pdf"), 8, 6)
for (v in vars_model) print(
  plot_gwr_coef(df_gwr, v, ripgeo, legend = TRUE)    # niente legenda, più compatto
)
dev.off()

save_rds(gwr_mod,"gwr_model")

# 3.8 Loop per dimensioni & output ------------------------------------

dimensioni <- list(Micro=imprese_m_sf_tr, Piccola=imprese_P_sf_tr,
                   Media=imprese_M_sf_tr, Grande=imprese_G_sf_tr)
risultati_spaziali <- list()

# loop principale sulle 4 classi dimensionali --------------------------
open_log("analisi_dimensioni")          # (opzionale) log diagnostico
# ---------- prima del ciclo ----------
for (dn in names(dimensioni)){
  # 1. sottocartelle per questa dimensione
  dir.create(file.path("02_maps",  dn), showWarnings = FALSE)
  dir.create(file.path("03_tables",dn), showWarnings = FALSE)
  dir.create(file.path("01_rds",   dn), showWarnings = FALSE)
  
  # ---------- analisi punti ----------
  cat("\n--- Analisi dimensione:", dn, "---\n")
  res <- tryCatch(
    analizza_spaziale_punti(dimensioni[[dn]], ripgeo, dn),
    error = function(e){message(e); NULL}
  )
  if (is.null(res)) next                # passa alla dimensione successiva
  
  # 2. salvataggi “spaziali”
  save_rds(res$imprese_cat,   file.path(dn, "sf_pts"))
  sf::st_write(res$imprese_cat,
               file.path("01_rds", dn, paste0("sf_", dn, ".gpkg")),
               delete_dsn = TRUE, quiet = TRUE)
  
  save_pdf (res$lisa_map, file.path(dn, "lisa_map"  ))
  save_pdf (res$map_gi,   file.path(dn, "gi_map"    ))
  
  # ---------- analisi regressiva ----------
  reg <- analisi_regressiva_spaziale(res$imprese_cat,
                                     res$W_knn,
                                     vars_model)
  
  # 3. salvataggi regressione
  save_rds(reg,                file.path(dn, "reg_models"))
  capture.output(summary(reg$sar),
                 file = file.path("03_tables", dn, "sar_summary.txt"))
  capture.output(summary(reg$sdm),
                 file = file.path("03_tables", dn, "sdm_summary.txt"))
  if(!is.null(reg$impacts))
    capture.output(summary(reg$impacts),
                   file = file.path("03_tables", dn, "sdm_impacts.txt"))
  # ---------- GWR LOCAL --------------------------------------------------
  gwr_res <- run_gwr(res$imprese_cat,
                     vars_model = vars_model,
                     listw      = res$W_knn,
                     samp_size = Inf)   # listw resta per simmetria
  
  # 1. Salva l’oggetto modello (.rds)
  save_rds(gwr_res, file.path(dn, "gwr_results"))   # 01_rds/<dn>/gwr_results.rds
  
  # 2. PDF “collezione” – tutte le variabili in un unico file
  pdf(file.path("02_maps", dn, "gwr_coefficients_collection.pdf"), 8, 6)
  
  for (v in vars_model) {
    
    # disegna il coefficiente locale
    p <- plot_gwr_coef(gwr_res$coeff_sf, v, ripgeo, legend = TRUE)
    print(p)                                       
  }
  
  dev.off()
}

close_log()      # chiude eventuale sink()


# 3.9 Dashboard / Leaflet summary --------------------------------------

leaf_all <- leaflet(imprese_sf_jittered) %>% addTiles() %>%
  addCircleMarkers(color = ~case_when(LISA_cluster=="High-High" ~ "red",
                                      LISA_cluster=="Low-Low"  ~ "blue",
                                      TRUE ~ "grey"), radius=3, stroke=FALSE, fillOpacity=.6)

save_html(leaf_all,"mappa_cluster_completa")

# 3.10 Export tabelle riassuntive --------------------------------------

conteggio_totale <- imprese_sf_jittered |> st_drop_geometry() |> count(LISA_cluster)
save_table(conteggio_totale,"conteggio_totale_cluster")

# fine script -----------------------------------------------------------
cat("\nSCRIPT COMPLETATO – tutti gli output nella gerarchia delle cartelle create.\n")
