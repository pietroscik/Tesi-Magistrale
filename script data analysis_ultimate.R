###############################################################################
############ TESI MAGISTRALE - UN'ANALISI SPAZIALE DELLE ######################
############ PERFORMANCE ECONOMICHE E FINANZIARIE DELLE #######################
########################## IMPRESE ITALIANE ###################################
###############################################*X*PIETRO MAIETTA*X*############
###############################################################################
# Carica le librerie
# Vettore di pacchetti richiesti
packages <- c(
  "leaflet.extras", "leaflet", "ggplot2", "psych", "openxlsx", "dplyr", 
  "readxl", "plotly", "pheatmap", "tidyr", "purrr", "MASS", "reshape2", 
  "DT", "nortest", "bestNormalize", "e1071", "moments", "rstatix", 
  "car", "lmtest", "corrplot", "tidyverse","irr","glmnet","kableExtra",
  "sandwich","caret"
)

# Funzione per installare e caricare i pacchetti
install_and_load <- function(pkg) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  } else {
    library(pkg, character.only = TRUE)
  }
}

# Applica la funzione a tutti i pacchetti
invisible(lapply(packages, install_and_load))

##### Script Unione ####
# Lista dei percorsi dei file
file_paths <- c(
  "C:/Users/39329/Downloads/aida tesi nord-est3.xlsx",
  "C:/Users/39329/Downloads/aida tesi nord-est2.xlsx",
  "C:/Users/39329/Downloads/aida tesi nord-est1.xlsx",
  "C:/Users/39329/Downloads/aida tesi isole.xlsx",
  "C:/Users/39329/Downloads/aida tesi sud2.xlsx",
  "C:/Users/39329/Downloads/aida tesi monza.xlsx",
  "C:/Users/39329/Downloads/aida tesi sud.xlsx",
  "C:/Users/39329/Downloads/aida tesi milano2.xlsx",
  "C:/Users/39329/Downloads/aida tesi milano1.xlsx",
  "C:/Users/39329/Downloads/aziende tesi piemonte2.xlsx",
  "C:/Users/39329/Downloads/aziende tesi piemonte1.xlsx",
  "C:/Users/39329/Downloads/aziende tesi piemonte.xlsx",
  "C:/Users/39329/Downloads/aida centro1.xlsx",
  "C:/Users/39329/Downloads/aida tesi centro.xlsx"
)

# Funzione per leggere la seconda scheda di ciascun file
read_second_sheet <- function(file) {
  read_excel(file, sheet = 2)
}

# Leggi tutti i file dalla seconda scheda
datasets <- lapply(file_paths, read_second_sheet)

# Funzione aggiornata per standardizzare i tipi di dati
standardize_columns <- function(datasets) {
  datasets[] <- lapply(datasets, as.character)  # Converti ogni colonna in character
  return(datasets)
}

# Applica la funzione a tutti i dataset
datasets <- lapply(datasets, standardize_columns)

# Unisci i dataset
combined_data <- bind_rows(datasets)

# Verifica la struttura del dataset unito
dim(combined_data)
summary(combined_data)
head(combined_data)
colSums(is.na(combined_data))
#Conversione dei valori n.d e n.s. in NA
combined_data[combined_data == "n.d." | combined_data == "n.s."] <- NA

# Rimuove righe con tutti i valori NA
combined_data <- combined_data[rowSums(is.na(combined_data)) != ncol(combined_data), ]
#rimuovi tutte le righe che non presentano valore sulla colonna Ragione sociale
combined_data <- combined_data[!is.na(combined_data$`Ragione sociale`) & combined_data$`Ragione sociale`!= "", ]

# Verifica le dimensioni dopo la rimozione
dim(combined_data)
# Scrivi il dataset su file Excel
write.xlsx(combined_data,"dataset.xlsx")

##### Pulizia e formattazione ####
dataset <- combined_data[, -1]  # Rimuove la prima colonna
View(dataset)
str(dataset)
# Contare i valori mancanti per colonna
colSums(!is.na(dataset))

# Visualizzare le prime righe del dataset
head(dataset)

# Dimensioni del dataset
dim(dataset)
# Rimuovi le colonne con nomi NA
dataset <- dataset[, !is.na(names(dataset))]

#Ridefinizione variabili e suddivisione settore
#di interesse per l'analisi
class(dataset)
summary(dataset)
nomi_colonne<-colnames(dataset)
nomi_colonne
colnames(dataset) <- c(
  "ragione_sociale", 
  "numero_CCIAA", 
  "provincia", 
  "codice_fiscale", 
  "chiusura_bilancio_ultimo_anno", 
  "ricavi_vendite_migl_EUR", 
  "dipendenti", 
  "CAP", 
  "regione", 
  "ISTAT_regione", 
  "ISTAT_provincia", 
  "ISTAT_comune", 
  "longitudine", 
  "latitudine", 
  "stato_giuridico", 
  "anno_fondazione", 
  "forma_giuridica", 
  "ricavi_vendite_ultimo_anno_migl_EUR", 
  "utile_netto_ultimo_anno_migl_EUR", 
  "totale_attivita_ultimo_anno_migl_EUR", 
  "dipendenti_ultimo_anno", 
  "dipendenti_fonte_ultimo_anno", 
  "capitalizzazione_corrente_migl_EUR", 
  "data_capitalizzazione_corrente", 
  "fatturato_estimato_migl_EUR", 
  "capitale_sociale_migl_EUR", 
  "indicatore_indipendenza_BvD", 
  "num_societa_gruppo", 
  "num_azionisti", 
  "num_partecipate", 
  "ateco_2007_codice", 
  "ateco_2007_descrizione", 
  "nace_rev2_codice", 
  "nace_rev2_descrizione", 
  "nome_gruppo_pari", 
  "descrizione_gruppo_pari", 
  "dimensione_gruppo_pari", 
  "overview_completa", 
  "storico", 
  "linea_business_principale", 
  "linea_business_secondaria", 
  "attivita_principale", 
  "attivita_secondaria", 
  "prodotti_servizi_principali", 
  "stima_dimensione", 
  "strategia_organizzazione_policy", 
  "alleanze_strategiche", 
  "socio_network", 
  "marchi_principali", 
  "stato_principale", 
  "stati_reg_stranieri_principali", 
  "siti_produzione_principali", 
  "siti_distribuzione_principali", 
  "siti_rappresentanza_principali", 
  "clienti_principali", 
  "indicatore_info_concise", 
  "societa_artigiana", 
  "startup_innovativa", 
  "PMI_innovativa", 
  "operatore_estero", 
  "EBITDA_migl_EUR", 
  "utile_netto_migl_EUR", 
  "totale_attivita_migl_EUR", 
  "patrimonio_netto_migl_EUR", 
  "posizione_finanziaria_netta_migl_EUR", 
  "EBITDA_su_vendite_percentuale", 
  "ROS_percentuale", 
  "ROA_percentuale", 
  "ROE_percentuale", 
  "debt_equity_ratio_percentuale", 
  "debiti_banche_fatt_percentuale", 
  "debt_EBITDA_ratio_percentuale", 
  "rotazione_cap_investito_volte", 
  "indice_liquidita", 
  "indice_corrente", 
  "indice_indebitamento_breve_percentuale", 
  "indice_indebitamento_lungo_percentuale", 
  "indice_copertura_immob_patrimoniale_percentuale", 
  "grado_ammortamento_percentuale", 
  "rapporto_indebitamento", 
  "indice_copertura_immob_finanziario_percentuale", 
  "debiti_banche_fatt_percentuale", 
  "costo_denaro_prestito_percentuale", 
  "grado_copertura_interessi_percentuale", 
  "oneri_finanziari_fatt_percentuale", 
  "indice_indipendenza_finanziaria_percentuale", 
  "grado_indipendenza_terzi_percentuale", 
  "posizione_finanziaria_netta", 
  "debt_equity_ratio_percentuale", 
  "debt_EBITDA_ratio_percentuale", 
  "rotazione_cap_investito", 
  "rotazione_cap_circolante_lordo", 
  "incidenza_circolante_operativo_percentuale", 
  "giacenza_media_scorte_gg", 
  "giorni_copertura_scorte_gg", 
  "durata_media_crediti_lordo_IVA_gg", 
  "durata_media_debiti_lordo_IVA_gg", 
  "durata_ciclo_commerciale_gg", 
  "EBITDA", 
  "EBITDA_su_vendite_percentuale", 
  "ROA_percentuale", 
  "ROI_percentuale", 
  "ROS_percentuale", 
  "ROE_percentuale", 
  "incidenza_oneri_extrag_percentuale", 
  "dipendenti_1", 
  "ricavi_pro_capite_EUR", 
  "valore_aggiunto_pro_capite_EUR", 
  "costo_lavoro_addetto_EUR", 
  "rendimento_dipendenti", 
  "capitale_circolante_netto_migl_EUR", 
  "margine_consumi_migl_EUR", 
  "margine_tesoreria_migl_EUR", 
  "margine_struttura_migl_EUR", 
  "flusso_cassa_gestione_migl_EUR",
  "azioni_quote_titoli_migl_EUR",
  "crediti_verso enti_creditizi_migl_EUR",
  "totale_attivo_migl_EUR",
  "debiti_verso enti_creditizi_migl_EUR",
  "capitale_mgl_EUR",
  "commissioni_attive_migl_EUR",
  "commissionipassive_migl_EUR",
  "profitti_perdite_opfin_migl_EUR",
  "utile_perdita_esercizio_migl_EUR",
  "margine_interesse_migl_EUR",
  "margine_intermediazione_migl_EUR",
  "risultato_lordo_gestione_migl_EUR",
  "dipendenti_2",
  "dm_compenso_salario",
  "dm_compenso_totale",
  "dm_data_compenso_salario",
  "dm_data_compenso_totale"
)
colnames(dataset)
# Identifica le colonne con nomi duplicati ed elimina quelle 
nomi_colonne <- names(dataset)
# Calcola la percentuale di valori mancanti per colonna
na_percentage <- sapply(dataset, function(x) mean(is.na(x)) * 100)

# Filtra le colonne con più del 30% di dati mancanti
colonne_da_rimuovere <- names(na_percentage)[na_percentage >= 30]
dataset_ridotto <- dataset[, !(names(dataset) %in% colonne_da_rimuovere)]

comparison_matrix <- matrix(FALSE, nrow = ncol(dataset_ridotto), ncol = ncol(dataset_ridotto))
colnames(comparison_matrix) <- colnames(dataset_ridotto)
rownames(comparison_matrix) <- colnames(dataset_ridotto)


# Confronta tutte le colonne, escludendo il confronto con se stessa
for (i in 1:ncol(dataset_ridotto)) {
  for (j in 1:ncol(dataset_ridotto)) {
      comparison_matrix[i, j] <- all(dataset_ridotto[[i]] == dataset_ridotto[[j]])
    }
  }


# Visualizza le colonne uguali
equal_columns <- which(comparison_matrix == TRUE, arr.ind = TRUE)
if (length(equal_columns) > 0) {
  cat("Colonne uguali:\n")
  for (idx in 1:nrow(equal_columns)) {
    # Aggiungi il filtro per escludere i confronti tra la stessa colonna
    if (equal_columns[idx, 1] != equal_columns[idx, 2]) {
      cat(colnames(dataset_ridotto)[equal_columns[idx, 1]], "<->", colnames(dataset_ridotto)[equal_columns[idx, 2]], "\n")
    }
  }
} else {
  cat("Nessuna colonna uguale trovata.\n")
}
#Esclusione colonne uguali manualmente su indicazione equal_colums
colonne_da_escludere <- c("dipendenti_ultimo_anno","dipendenti_1",
                          "dipendenti_2","totale_attivita_ultimo_anno_migl_EUR",
                          "totale_attivo_migl_EUR",
                          "utile_perdita_esercizio_migl_EUR",
                          "ricavi_vendite_ultimo_anno_migl_EUR",
                          "utile_netto_ultimo_anno_migl_EUR")

# Escludi le colonne dal dataset_ridotto
dataset_ridotto <- dataset_ridotto[, !(colnames(dataset_ridotto) %in% colonne_da_escludere)]

# Visualizza il dataset ridotto con le colonne escludenti
print(dataset_ridotto)
colnames(dataset_ridotto)

# Calcola la percentuale di dati mancanti per ciascuna colonna
na_percentage <- sapply(dataset_ridotto, function(x) mean(is.na(x)) * 100)
na_percentage <- sort(na_percentage, decreasing = TRUE)

# Stampare la percentuale di valori mancanti
cat("Percentuale di valori mancanti per colonna:\n")
print(na_percentage)


# Rimuovi le righe che contengono almeno un NA
dataset_completo <- na.omit(dataset_ridotto)
# Rimuovi le colonne duplicate per nome specifico
any(duplicated(names(dataset_completo)))
duplicated_cols <- names(dataset_completo)[duplicated(names(dataset_completo))]
print(duplicated_cols)
dataset_completo <- dataset_completo[, !duplicated(colnames(dataset_completo))]
colnames(dataset_completo)
# Cambio tipo variabili
dataset_completo <- dataset_completo %>%
  mutate(
    ragione_sociale=as.character(ragione_sociale),
    numero_CCIAA=as.character(numero_CCIAA),
    provincia = factor(provincia),
    codice_fiscale=as.character(codice_fiscale),
    chiusura_bilancio_ultimo_anno=as.character(chiusura_bilancio_ultimo_anno),
    ricavi_vendite_migl_EUR=as.numeric(ricavi_vendite_migl_EUR),
    dipendenti=as.numeric(dipendenti),
    CAP = factor(CAP),
    regione = factor(regione),
    ISTAT_regione = factor(ISTAT_regione),
    ISTAT_provincia = factor(ISTAT_provincia),
    ISTAT_comune = factor(ISTAT_comune),
    longitudine=as.numeric(longitudine),
    latitudine=as.numeric(latitudine),
    stato_giuridico = factor(stato_giuridico),
    forma_giuridica = factor(forma_giuridica),
    ricavi_vendite_migl_EUR = as.numeric(ricavi_vendite_migl_EUR),
    utile_netto_migl_EUR = as.numeric(utile_netto_migl_EUR),
    dipendenti_fonte_ultimo_anno=as.character(dipendenti_fonte_ultimo_anno),
    capitale_sociale_migl_EUR=as.numeric(capitale_sociale_migl_EUR),
    indicatore_indipendenza_BvD = factor(indicatore_indipendenza_BvD),
    num_societa_gruppo = factor(num_societa_gruppo),
    num_azionisti = factor(num_azionisti),
    num_partecipate = factor(num_partecipate),
    ateco_2007_codice = factor(ateco_2007_codice),
    ateco_2007_descrizione = as.character(ateco_2007_descrizione),
    nace_rev2_codice = factor(nace_rev2_codice),
    nace_rev2_descrizione = as.character(nace_rev2_descrizione),
    nome_gruppo_pari   = factor(nome_gruppo_pari),
    descrizione_gruppo_pari=as.character(descrizione_gruppo_pari),
    dimensione_gruppo_pari = as.numeric(dimensione_gruppo_pari),
    overview_completa = as.character(overview_completa),
    linea_business_principale = as.character(linea_business_principale),
    attivita_principale=as.character(attivita_principale),
    prodotti_servizi_principali=as.character(prodotti_servizi_principali),
    stato_principale = factor(stato_principale),
    societa_artigiana = factor(societa_artigiana),
    startup_innovativa = factor(startup_innovativa),
    PMI_innovativa = factor(PMI_innovativa),
    operatore_estero=factor(operatore_estero),
    EBITDA_migl_EUR = as.numeric(EBITDA_migl_EUR),
    utile_netto_migl_EUR=as.numeric(utile_netto_migl_EUR),
    totale_attivita_migl_EUR=as.numeric(totale_attivita_migl_EUR),
    patrimonio_netto_migl_EUR=as.numeric(patrimonio_netto_migl_EUR),
    posizione_finanziaria_netta_migl_EUR=as.numeric(posizione_finanziaria_netta_migl_EUR),
    EBITDA_su_vendite_percentuale=as.numeric(EBITDA_su_vendite_percentuale),
    ROE_percentuale = as.numeric(ROE_percentuale),
    ROS_percentuale = as.numeric(ROS_percentuale),
    ROA_percentuale = as.numeric(ROA_percentuale),
    debiti_banche_fatt_percentuale = as.numeric(debiti_banche_fatt_percentuale),
    debt_equity_ratio_percentuale = as.numeric(debt_equity_ratio_percentuale),
    debt_EBITDA_ratio_percentuale = as.numeric(debt_EBITDA_ratio_percentuale),
    rotazione_cap_investito_volte = as.numeric(rotazione_cap_investito_volte),
    indice_liquidita = as.numeric(indice_liquidita),
    indice_corrente = as.numeric(indice_corrente),
    indice_indebitamento_breve_percentuale=as.numeric(indice_indebitamento_breve_percentuale),
    indice_indebitamento_lungo_percentuale=as.numeric(indice_indebitamento_lungo_percentuale),
    indice_copertura_immob_patrimoniale_percentuale = as.numeric(indice_copertura_immob_patrimoniale_percentuale),
    rapporto_indebitamento=as.numeric(rapporto_indebitamento),
    indice_copertura_immob_finanziario_percentuale = as.numeric(indice_copertura_immob_finanziario_percentuale),
    grado_copertura_interessi_percentuale = as.numeric(grado_copertura_interessi_percentuale),
    oneri_finanziari_fatt_percentuale = as.numeric(oneri_finanziari_fatt_percentuale),
    indice_indipendenza_finanziaria_percentuale = as.numeric(indice_indipendenza_finanziaria_percentuale),
    grado_indipendenza_terzi_percentuale = as.numeric(grado_indipendenza_terzi_percentuale),
    posizione_finanziaria_netta=as.numeric(posizione_finanziaria_netta),
    rotazione_cap_investito = as.numeric(rotazione_cap_investito),
    rotazione_cap_circolante_lordo = as.numeric(rotazione_cap_circolante_lordo),
    incidenza_circolante_operativo_percentuale = as.numeric(incidenza_circolante_operativo_percentuale),
    giacenza_media_scorte_gg = as.numeric(giacenza_media_scorte_gg),
    durata_media_crediti_lordo_IVA_gg = as.numeric(durata_media_crediti_lordo_IVA_gg),
    durata_media_debiti_lordo_IVA_gg = as.numeric(durata_media_debiti_lordo_IVA_gg),
    EBITDA=as.numeric(EBITDA),
    ROI_percentuale = as.numeric(ROI_percentuale),
    incidenza_oneri_extrag_percentuale = as.numeric(incidenza_oneri_extrag_percentuale),
    ricavi_pro_capite_EUR = as.numeric(ricavi_pro_capite_EUR),
    valore_aggiunto_pro_capite_EUR = as.numeric(valore_aggiunto_pro_capite_EUR),
    costo_lavoro_addetto_EUR = as.numeric(costo_lavoro_addetto_EUR),
    rendimento_dipendenti = as.numeric(rendimento_dipendenti),
    capitale_circolante_netto_migl_EUR=as.numeric(capitale_circolante_netto_migl_EUR),
    margine_consumi_migl_EUR=as.numeric(margine_consumi_migl_EUR),
    margine_tesoreria_migl_EUR=as.numeric(margine_tesoreria_migl_EUR),
    margine_struttura_migl_EUR = as.numeric(margine_struttura_migl_EUR),
    flusso_cassa_gestione_migl_EUR=as.numeric(flusso_cassa_gestione_migl_EUR),
    commissioni_attive_migl_EUR=as.numeric(commissioni_attive_migl_EUR)
    ) %>%
  drop_na()  # Rimuovi righe con NA se necessario

# Verifica il risultato
colnames(dataset_completo)
# Visualizza il dataset dopo la rimozione delle righe
summary(dataset_completo)
#write.csv(dataset_completo, "dataset_completo.csv", row.names = FALSE)
write.xlsx(dataset_completo, "dataset_completo.xlsx")

# Rimuovi tutte le variabili tranne dataset, dataset_ridotto e dataset_completo
rm(list = setdiff(ls(), c( "dataset_completo")))

# Verifica che solo le variabili desiderate siano rimaste
ls()


##### Divisione del set per categoria####
#circa dataset dopo pulizia ambiente
dataset_completo <- read_excel("dataset_completo.xlsx")
#suddivisione dataset per area geografica
dataset_completo <- dataset_completo %>%
  mutate(macroarea = case_when(
    regione %in% c("Valle d'Aosta/Vallée d'Aoste", "Piemonte", "Liguria", "Lombardia") ~ "Nord-Ovest",
    regione %in% c("Veneto", "Trentino-Alto Adige", "Friuli-Venezia Giulia", "Emilia-Romagna") ~ "Nord-Est",
    regione %in% c("Toscana", "Umbria", "Marche", "Lazio", "Abruzzo") ~ "Centro",
    regione %in% c("Campania", "Molise", "Puglia", "Basilicata", "Calabria", "Sicilia", "Sardegna") ~ "Sud",
    TRUE ~ "Errore" # Per evidenziare eventuali regioni non classificate
  ))%>%
  mutate(macroarea=as.factor(macroarea))
#visualizzazione delle statistiche per macroarea
dataset_completo %>%
  group_by(macroarea) %>%
  summarise(n_aziende = n())
dataset_completo %>%
  group_by(regione, macroarea) %>%
  summarise(n_aziende = n(), .groups = 'drop') %>%
  arrange(macroarea) 
summary(dataset_completo)

# Subset Completo di tutte le variabili
aree_di_interesse <- list(
  Identificativa = c("ragione_sociale", "numero_CCIAA", "codice_fiscale",
                     "ateco_2007_codice", "ateco_2007_descrizione",
                     "nome_gruppo_pari", "descrizione_gruppo_pari", "dimensione_gruppo_pari", "overview_completa"),
  
  Finanziaria = c("ricavi_vendite_migl_EUR", "utile_netto_migl_EUR", 
                  "EBITDA_migl_EUR", "patrimonio_netto_migl_EUR", 
                  "posizione_finanziaria_netta_migl_EUR",
                  "ROE_percentuale", "ROI_percentuale", "ROS_percentuale", 
                  "ROA_percentuale", "debiti_banche_fatt_percentuale",
                  "debt_equity_ratio_percentuale", "debiti_banche_fatt_percentuale",
                  "debt_EBITDA_ratio_percentuale"),
  
  Operativa = c("dipendenti", "giacenza_media_scorte_gg", 
                "durata_media_crediti_lordo_IVA_gg", "durata_media_debiti_lordo_IVA_gg",
                "incidenza_circolante_operativo_percentuale", "flusso_cassa_gestione_migl_EUR"),
  
  Geografica = c("macroarea","regione", "provincia", "ISTAT_regione", "ISTAT_provincia",
                 "ISTAT_comune", "longitudine", "latitudine", "CAP"),
  
  Giuridica = c("stato_giuridico", "forma_giuridica", "societa_artigiana", 
                "startup_innovativa", "PMI_innovativa", "operatore_estero"),
  
  Attività = c("linea_business_principale", "attivita_principale", "prodotti_servizi_principali",
               "nace_rev2_codice", "nace_rev2_descrizione"),
  
  Indicatori = c("indice_liquidita", "indice_corrente", "indice_indebitamento_breve_percentuale",
                 "indice_indebitamento_lungo_percentuale", "indice_copertura_immob_patrimoniale_percentuale",
                 "rapporto_indebitamento", "indice_copertura_immob_finanziario_percentuale",
                 "grado_indipendenza_terzi_percentuale", "posizione_finanziaria_netta",
                 "grado_copertura_interessi_percentuale", "oneri_finanziari_fatt_percentuale",
                 "indice_indipendenza_finanziaria_percentuale")
)

#Redditività e Performance Operativa
variabili_redditivita <- c("ROE_percentuale", "ROI_percentuale", "ROS_percentuale", 
                           "ROA_percentuale", "EBITDA_migl_EUR", "utile_netto_migl_EUR")
#Solidità Finanziaria e Indebitamento
variabili_solidita <- c("patrimonio_netto_migl_EUR", "posizione_finanziaria_netta_migl_EUR", 
                        "debt_equity_ratio_percentuale", "debt_EBITDA_ratio_percentuale")
#Produttività e Gestione del Capitale Umano
variabili_produttivita <- c("dipendenti", "ricavi_pro_capite_EUR", "valore_aggiunto_pro_capite_EUR", 
                            "rendimento_dipendenti")
# Liquidità e Gestione Finanziaria a Breve Termine
variabili_liquidita <- c("indice_liquidita", "indice_corrente", "flusso_cassa_gestione_migl_EUR")
#Capitale Circolante e Gestione delle Scorte
variabili_circolante <- c("giacenza_media_scorte_gg", "durata_media_crediti_lordo_IVA_gg", 
                          "durata_media_debiti_lordo_IVA_gg", "incidenza_circolante_operativo_percentuale")


# Lista delle categorie di variabili
categorie_variabili <- list(
  Redditività = variabili_redditivita,
  Solidità = variabili_solidita,
  Produttività = variabili_produttivita,
  Liquidità = variabili_liquidita,
  Capitale_Circolante = variabili_circolante
)
##### Classificazione per settore (ATECO)####
#Suddivisione codice ATECO xx.xx.xx
dataset_completo <- dataset_completo %>%
  mutate(
    sezione = substr(ateco_2007_codice, 1, 1),
    divisione = substr(ateco_2007_codice, 1, 2),
    gruppo = substr(ateco_2007_codice, 1, 3),
    classe = substr(ateco_2007_codice, 1, 5),
    categoria = substr(ateco_2007_codice, 1, 6)
  )
dataset_completo <- dataset_completo %>%
  mutate(divisione_ateco = as.numeric(substr(as.character(ateco_2007_codice), 1, 2))) %>%
  mutate(sezione = case_when(
    divisione_ateco >= 1  & divisione_ateco <= 3  ~ "A - Agricoltura, silvicoltura e pesca",
    divisione_ateco >= 5  & divisione_ateco <= 9  ~ "B - Estrazione di minerali da cave e miniere",
    divisione_ateco >= 10 & divisione_ateco <= 33 ~ "C - Attività manifatturiere",
    divisione_ateco == 35                        ~ "D - Fornitura di energia elettrica, gas, vapore e aria condizionata",
    divisione_ateco >= 36 & divisione_ateco <= 39 ~ "E - Fornitura di acqua; gestione rifiuti e risanamento",
    divisione_ateco >= 41 & divisione_ateco <= 43 ~ "F - Costruzioni",
    divisione_ateco >= 45 & divisione_ateco <= 47 ~ "G - Commercio all’ingrosso e al dettaglio",
    divisione_ateco >= 49 & divisione_ateco <= 53 ~ "H - Trasporto e magazzinaggio",
    divisione_ateco >= 55 & divisione_ateco <= 56 ~ "I - Servizi di alloggio e ristorazione",
    divisione_ateco >= 58 & divisione_ateco <= 63 ~ "J - Servizi di informazione e comunicazione",
    divisione_ateco >= 64 & divisione_ateco <= 66 ~ "K - Attività finanziarie e assicurative",
    divisione_ateco == 68                         ~ "L - Attività immobiliari",
    divisione_ateco >= 69 & divisione_ateco <= 75 ~ "M - Attività professionali, scientifiche e tecniche",
    divisione_ateco >= 77 & divisione_ateco <= 82 ~ "N - Noleggio, agenzie di viaggio, servizi di supporto",
    divisione_ateco == 84                         ~ "O - Amministrazione pubblica e difesa",
    divisione_ateco == 85                         ~ "P - Istruzione",
    divisione_ateco >= 86 & divisione_ateco <= 88 ~ "Q - Sanità e assistenza sociale",
    divisione_ateco >= 90 & divisione_ateco <= 93 ~ "R - Attività artistiche, sportive, di intrattenimento",
    divisione_ateco >= 94 & divisione_ateco <= 96 ~ "S - Altre attività di servizi",
    divisione_ateco >= 97 & divisione_ateco <= 98 ~ "T - Attività di famiglie e convivenze",
    divisione_ateco == 99                         ~ "U - Organizzazioni ed organismi extraterritoriali",
    TRUE ~ "Non classificato"
  ))


##### Classificazione dimensione Impresa####
dataset_completo <- dataset_completo %>%
  mutate(
    dimensione_impresa = case_when(
      # Microimpresa: deve rispettare almeno 2 criteri su 3
      (dipendenti < 10 & ricavi_vendite_migl_EUR <= 2500 & totale_attivita_migl_EUR <= 2500) |
        (dipendenti < 10 & ricavi_vendite_migl_EUR <= 2500) |
        (dipendenti < 10 & totale_attivita_migl_EUR <= 2500) ~ "Micro",
      
      # Piccola impresa: almeno 2 criteri su 3
      (dipendenti < 50 & ricavi_vendite_migl_EUR <= 12500 & totale_attivita_migl_EUR <= 12500) |
        (dipendenti < 50 & ricavi_vendite_migl_EUR <= 12500) |
        (dipendenti < 50 & totale_attivita_migl_EUR <= 12500) ~ "Piccola",
      
      # Media impresa: almeno 2 criteri su 3
      (dipendenti < 250 & ricavi_vendite_migl_EUR <= 62500 & totale_attivita_migl_EUR <= 53750) |
        (dipendenti < 250 & ricavi_vendite_migl_EUR <= 62500) |
        (dipendenti < 250 & totale_attivita_migl_EUR <= 53750)  ~ "Media",
      
      # Grande impresa: tutte le altre
      TRUE ~ "Grande"
    )
  ) %>%
  mutate(
    dimensione_impresa = factor(dimensione_impresa, 
                                levels = c("Micro", "Piccola", "Media", "Grande")
    )
  )

# Visualizzare il dataset con la nuova colonna dimensione_impresa
head(dataset_completo)

# Verifica risultati
n_dimensione_imprese <- dataset_completo %>%
  group_by(dimensione_impresa) %>%
  summarise(
    n_aziende = n(),
    media_dipendenti = mean(dipendenti),
    media_fatturato = mean(ricavi_vendite_migl_EUR),
    media_valore_bilancio = mean(totale_attivita_migl_EUR),
    .groups = "drop"
  ) %>%
  mutate(percentuale = n_aziende/sum(n_aziende)*100)

print(n_dimensione_imprese)
# Salvataggio
write.xlsx(n_dimensione_imprese, "classificazione_UE_2023.xlsx")

# Numero di aziende per settore
n_aziende_settore <- dataset_completo %>%
  group_by(sezione) %>%
  summarise(n_aziende = n(), .groups = 'drop') %>%
  mutate(percentuale = (n_aziende / sum(n_aziende)) * 100) %>%
  arrange(desc(n_aziende))  # Ordina per numero di aziende

# Numero di aziende per settore, dimensione e macroarea
n_aziende_settore_dimensione <- dataset_completo %>%
  group_by(sezione, dimensione_impresa, macroarea) %>%
  summarise(n_aziende = n(), .groups = "drop") %>%
  group_by(sezione) %>%
  mutate(percentuale = (n_aziende / sum(n_aziende)) * 100) %>%
  pivot_wider(
    names_from = dimensione_impresa,
    values_from = c(n_aziende, percentuale),
    names_glue = "{dimensione_impresa}_{.value}"
  ) 

# Numero di aziende per settore e macroarea
n_aziende_macroarea_settore <- dataset_completo %>%
  group_by(sezione, macroarea) %>%
  summarise(n_aziende = n(), .groups = 'drop') %>%
  mutate(percentuale = (n_aziende / sum(n_aziende)) * 100) %>%
  arrange(desc(n_aziende))

# Numero di aziende per macroarea
n_aziende_macroarea <- dataset_completo %>%
  group_by(macroarea) %>%
  summarise(n_aziende = n(), .groups = "drop") %>%
  mutate(percentuale = (n_aziende / sum(n_aziende)) * 100) %>%
  arrange(desc(n_aziende))

# Numero di aziende per regione
n_aziende_regione <- dataset_completo %>%
  group_by(regione) %>%
  summarise(n_aziende = n(), .groups = "drop") %>%
  mutate(percentuale = (n_aziende / sum(n_aziende)) * 100) %>%
  arrange(desc(n_aziende))

# Stampa i risultati in console
print(n_aziende_settore)
print(n_aziende_settore_dimensione)
print(n_aziende_macroarea_settore)
print(n_aziende_macroarea)
print(n_aziende_regione)

# Conversione il dataset da wide a long
n_aziende_settore_dimensione_long <- n_aziende_settore_dimensione %>%
  pivot_longer(cols = ends_with("_n_aziende"), 
               names_to = "dimensione_impresa", 
               values_to = "n_aziende") %>%
  mutate(dimensione_impresa = gsub("_n_aziende", "", dimensione_impresa)) # Puliamo i nomi

# Creo il grafico corretto
ggplot(n_aziende_settore_dimensione_long, aes(x = sezione, y = n_aziende, fill = dimensione_impresa)) +
  geom_bar(stat = "identity", position = "fill") +  # Stack normalizzato (percentuale)
  coord_flip() +
  theme_minimal() +
  labs(title = "Distribuzione delle aziende per settore e dimensione",
       x = "Settore", y = "Percentuale")  +
  theme(legend.position = "bottom")

#Numero di aziende per settore (bar chart)
ggplot(n_aziende_settore, aes(x = reorder(sezione, n_aziende), y = n_aziende, fill = sezione)) +
  geom_bar(stat = "identity") +
  coord_flip() +  # Ruota l'asse per migliorare leggibilità
  theme_minimal() +
  labs(title = "Numero di aziende per settore", x = "Settore", y = "Numero di aziende") +
  theme(legend.position = "none")

# Convertire macroarea in un fattore con ordine specifico
n_aziende_settore_dimensione_long$macroarea <- factor(n_aziende_settore_dimensione_long$macroarea, 
                                                      levels = c("Nord-Ovest", "Nord-Est", "Centro", "Sud"))
# Percentuale di aziende per dimensione in ogni settore (bar chart stacked)
ggplot(n_aziende_settore_dimensione_long, aes(x = reorder(sezione, -n_aziende), y = n_aziende, fill = dimensione_impresa)) +
  geom_bar(stat = "identity", position = "stack") +  
  coord_flip() +
  theme_minimal() +
  labs(title = "Distribuzione delle aziende per settore, dimensione e macroarea",
       x = "Settore", y = "Numero di aziende")+  
  theme(legend.position = "bottom") +
  facet_wrap(~ macroarea, ncol = 2)  # Mantiene l'ordine definito sopra

#Numero di aziende per macroarea (bar chart)
ggplot(n_aziende_macroarea, aes(x = reorder(macroarea, -n_aziende), y = n_aziende, fill = macroarea)) +
  geom_bar(stat = "identity") +
  coord_flip() +
  theme_minimal() +
  labs(title = "Numero di aziende per macroarea", x = "Macroarea", y = "Numero di aziende") +
  theme(legend.position = "none")

###Salva n_aziende
write.xlsx(n_aziende_regione, "numero_aziende_regione.xlsx")
write.xlsx(n_aziende_macroarea, "numero_aziende_macroarea.xlsx")
write.xlsx(n_aziende_macroarea_settore, "numero_aziende_macroarea_settore.xlsx")
write.xlsx(n_aziende_settore,"numero_aziende_per_settore.xlsx")
write.xlsx(n_aziende_settore_dimensione,"numero_dimensione_aziende_per_settore&area.xlsx")
write.xlsx(dataset_completo, "dataset.xlsx")

##### Indicatore Sintetico Posizione(media ponderata degli indici di categoria)#####
dataset_completo <- readRDS("matrice.rds")
# Creazione PFN_EBITDA_migl_EUR  nel dataset_completo
dataset_completo$PFN_EBITDA_migl_EUR <- ifelse(
  dataset_completo$EBITDA_migl_EUR == 0,0,
  dataset_completo$posizione_finanziaria_netta_migl_EUR / dataset_completo$EBITDA_migl_EUR
)

# Funzione per normalizzare con range [0, 1000]
# Funzione per normalizzare in [0, 1] o [0, 1000]
normalize <- function(x, scale = 1) {
  rng <- range(x, na.rm = TRUE)
  if (diff(rng) == 0) return(rep(0, length(x)))  # evita divisione per zero
  (x - rng[1]) / (rng[2] - rng[1]) * scale
}

# 1. Normalizzazione delle variabili numeriche
dataset_normalizzato <- dataset_completo
dataset_normalizzato[, unlist(categorie_variabili)] <- lapply(
  dataset_completo[, unlist(categorie_variabili)],
  function(x) if (is.numeric(x)) normalize(x, scale = 1000) else x
)


# Variabili da normalizzare singolarmente (in [0,1])
variabili_da_normalizzare <- c(
  "capitale_sociale_migl_EUR", "utile_netto_migl_EUR","PFN_EBITDA_migl_EUR",
  "debt_equity_ratio_percentuale", "debiti_banche_fatt_percentuale",
  "rotazione_cap_investito_volte", "indice_indebitamento_breve_percentuale",
  "indice_indebitamento_lungo_percentuale",
  "indice_copertura_immob_patrimoniale_percentuale",
  "rapporto_indebitamento", "indice_copertura_immob_finanziario_percentuale",
  "grado_copertura_interessi_percentuale", "oneri_finanziari_fatt_percentuale",
  "indice_indipendenza_finanziaria_percentuale",
  "grado_indipendenza_terzi_percentuale",
  "posizione_finanziaria_netta", "rotazione_cap_circolante_lordo",
  "costo_lavoro_addetto_EUR","margine_consumi_migl_EUR",
  "margine_tesoreria_migl_EUR","margine_struttura_migl_EUR",
  "commissioni_attive_migl_EUR","ricavi_vendite_migl_EUR", 
  "capitale_sociale_migl_EUR","dimensione_gruppo_pari", 
  "totale_attivita_migl_EUR","rotazione_cap_investito"
)
variabili_non_trovate <- variabili_da_normalizzare[!variabili_da_normalizzare %in% names(dataset_completo)]
print(variabili_non_trovate)

for (v in variabili_da_normalizzare) {
  dataset_normalizzato[[v]] <- normalize(dataset_completo[[v]],scale = 1000)
}


# Calcola ISP per il gruppo a (Economico-Reddituali)
ISP_a <- (dataset_normalizzato$ROE_percentuale* 0.2727) + 
  (dataset_normalizzato$EBITDA_su_vendite_percentuale * 0.3560) + 
  (dataset_normalizzato$ROI_percentuale * 0.2386) + 
  (dataset_normalizzato$rotazione_cap_investito * 0.1327)

# Calcola ISP per il gruppo b (Patrimoniale-Finanziari)
ISP_b<-(dataset_normalizzato$debt_equity_ratio_percentuale * 0.3162) +
 (dataset_normalizzato$debt_EBITDA_ratio_percentuale * 0.2703) + 
 (dataset_normalizzato$totale_attivita_migl_EUR * 0.1583) +
 (dataset_normalizzato$PFN_EBITDA_migl_EUR * 0.2552)

hist(ISP_a * 0.4311) 
hist(ISP_b * 0.5689)

# Calcolare l'ISP
dataset_completo$ISP<- (ISP_a * 0.4311) + (ISP_b * 0.5689)
dataset_normalizzato$ISP<- ((ISP_a * 0.4311) + (ISP_b * 0.5689))
saveRDS(dataset_completo, "matrice.rds")

#aggiunta categoria variabili performance
##### Inferenza pesi ISP ####
dataset_completo<-readRDS("matrice.rds")
vars_model <- unname(unlist(categorie_variabili))

dataset_z_score<-as.data.frame(scale(dataset_completo[c(vars_model,"ISP","EBITDA_su_vendite_percentuale","rotazione_cap_investito",
                                                        "totale_attivita_migl_EUR","PFN_EBITDA_migl_EUR","indice_indipendenza_finanziaria_percentuale",
                                                        "grado_copertura_interessi_percentuale","indice_indebitamento_breve_percentuale",
                                                        "rapporto_indebitamento","oneri_finanziari_fatt_percentuale")]))
dataset_z_score$codice_fiscale<-dataset_completo$codice_fiscale
dataset_z_score$sezione<-dataset_completo$sezione
# Variabili del gruppo ISP_a
vars_a <- dataset_z_score[, c("ROE_percentuale", 
                              "EBITDA_su_vendite_percentuale", 
                              "ROI_percentuale", 
                              "rotazione_cap_investito")]

# PCA
pca_a <- prcomp(vars_a, scale. = TRUE)
summary(pca_a)

# Pesi dalla prima componente
pesi_pca_a <- pca_a$rotation[, 1]
pesi_pca_a_norm <- pesi_pca_a / sum(abs(pesi_pca_a))  # normalizzati in somma assoluta

# Calcolo del nuovo ISP_a con pesi ottimizzati
ISP_a_pca <- as.matrix(vars_a) %*% pesi_pca_a_norm
head(ISP_a_pca)

# Istogramma del nuovo ISP_a
hist(ISP_a_pca, main = "ISP_a - Pesi PCA", col = "skyblue", xlab = "ISP_a PCA")
# Regressione lineare su ISP
lm_a <- lm(ISP ~ ROE_percentuale + EBITDA_su_vendite_percentuale + ROI_percentuale + rotazione_cap_investito,
           data = dataset_z_score)

summary(lm_a)  # visualizza i coefficienti stimati

# Estrazione dei coefficienti (escludendo l'intercetta)
pesi_lm_a <- coef(lm_a)[-1]
pesi_lm_a_norm <- pesi_lm_a / sum(abs(pesi_lm_a))  # normalizzazione in somma assoluta

# Calcolo ISP_a ottimizzato
vars_a <- dataset_z_score[, names(pesi_lm_a)]
ISP_a_lm <- as.matrix(vars_a) %*% pesi_lm_a_norm
summary(ISP_a_lm)
head(ISP_a_lm)

# Istogramma del nuovo ISP_a
hist(ISP_a_lm, main = "ISP_a - Pesi Regressione", col = "tomato", xlab = "ISP_a LM")

vars_b <- dataset_z_score[, c(
  "debt_equity_ratio_percentuale", 
  "debt_EBITDA_ratio_percentuale", 
  "totale_attivita_migl_EUR", 
  "PFN_EBITDA_migl_EUR"
)]
# PCA
pca_b <- prcomp(vars_b, scale. = TRUE)
summary(pca_b)

# Pesi dalla prima componente
pesi_pca_b <- pca_b$rotation[, 1]
pesi_pca_b_norm <- pesi_pca_b / sum(abs(pesi_pca_b))  # normalizzati in somma assoluta

# Calcolo del nuovo ISP_a con pesi ottimizzati
ISP_b_pca <- as.matrix(vars_b) %*% pesi_pca_b_norm
head(ISP_b_pca)

# Istogramma del nuovo ISP_a
hist(ISP_b_pca, main = "ISP_b - Pesi PCA", col = "skyblue", xlab = "ISP_b PCA")
# Regressione lineare su ISP
lm_b <- lm(ISP ~debt_equity_ratio_percentuale+debt_EBITDA_ratio_percentuale+
             totale_attivita_migl_EUR+PFN_EBITDA_migl_EUR,
           data = dataset_z_score)

summary(lm_b)  # visualizza i coefficienti stimati

# Estrazione dei coefficienti (escludendo l'intercetta)
pesi_lm_b <- coef(lm_b)[-1]
pesi_lm_b_norm <- pesi_lm_b / sum(abs(pesi_lm_b))  # normalizzazione in somma assoluta

# Calcolo ISP_a ottimizzato
vars_b <- dataset_z_score[, names(pesi_lm_b)]
ISP_b_lm <- as.matrix(vars_b) %*% pesi_lm_b_norm
summary(ISP_b_lm)
head(ISP_b_lm)
# Istogramma del nuovo ISP_a
hist(ISP_b_lm, main = "ISP_b - Pesi Regressione", col = "tomato", xlab = "ISP_b LM")

# Min-max normalization (range 0–1000)
ISP_a_pca_norm <- normalize(ISP_a_pca,scale=1000)
ISP_b_pca_norm <- normalize(ISP_b_pca,scale=1000)

# ISP prima componente principale
ISP_pca_totale <- ISP_a_pca_norm * 0.5 + ISP_b_pca_norm * 0.5  # oppure pesi diversi
hist(ISP_pca_totale)

# Z-score standardization
standardize <- function(x) scale(x)[, 1]

ISP_a_lm_std <- standardize(ISP_a_lm)
ISP_b_lm_std <- standardize(ISP_b_lm)

ISP_lm_totale <- ISP_a_lm_std * 0.5 + ISP_b_lm_std * 0.5
hist(ISP_lm_totale)

# Normalizzazione
ISP_a_lm_norm  <- normalize(ISP_a_lm,scale = 1000)
ISP_b_lm_norm  <- normalize(ISP_b_lm,scale=1000)

# Combinazioni con pesi originali
pesa <- 0.4311
pesb <- 0.5689

ISP_totale_pca <- ISP_a_pca_norm * pesa + ISP_b_pca_norm * pesb
ISP_totale_lm  <- ISP_a_lm_norm  * pesa + ISP_b_lm_norm  * pesb

# Standardizzazione (z-score)
ISP_a_pca_std <- standardize(ISP_a_pca)
ISP_b_pca_std <- standardize(ISP_b_pca)

ISP_totale_pca_std <- ISP_a_pca_std * pesa + ISP_b_pca_std * pesb
ISP_totale_lm_std  <- ISP_a_lm_std  * pesa + ISP_b_lm_std  * pesb

# CONFRONTO VISIVO
par(mfrow = c(2, 2))  # 2x2 layout

hist(ISP_totale_pca, main = "ISP Totale (PCA normalizzato)", col = "skyblue", xlab = "ISP PCA Norm")
hist(ISP_totale_lm,  main = "ISP Totale (Regressione normalizzato)", col = "tomato", xlab = "ISP LM Norm")
hist(ISP_totale_pca_std, main = "ISP Totale (PCA standardizzato)", col = "skyblue3", xlab = "ISP PCA Std")
hist(ISP_totale_lm_std,  main = "ISP Totale (Regressione standardizzato)", col = "tomato3", xlab = "ISP LM Std")

par(mfrow = c(1, 1))  # ripristina layout

# CONFRONTO NUMERICO
df_isp <- data.frame(
  ISP=dataset_completo$ISP,
  ISP_pca_norm = ISP_totale_pca,
  ISP_lm_norm = ISP_totale_lm,
  ISP_std= dataset_z_score$ISP,
  ISP_pca_std = ISP_totale_pca_std,
  ISP_lm_std = ISP_totale_lm_std
)

summary(df_isp)

##### Stima di ulteriori variabili nel gruppo B####
# Funzione di normalizzazione 
normalize_1000 <- function(x) {
  rng <- range(x, na.rm = TRUE)
  if (diff(rng) == 0) return(rep(0, length(x)))  # evita divisione per zero se tutti i valori uguali
  (x - rng[1]) / (rng[2] - rng[1]) * 1000
}


# --- Variabili di base ---
vars_a <- c("ROE_percentuale", "EBITDA_su_vendite_percentuale", "ROI_percentuale", "rotazione_cap_investito")
vars_b <- c("debt_equity_ratio_percentuale", "debt_EBITDA_ratio_percentuale", "totale_attivita_migl_EUR", "PFN_EBITDA_migl_EUR")

# --- Variabili estese per gruppo B ---
vars_b_esteso <- c(vars_b,
                   "indice_liquidita", "indice_corrente", "indice_indebitamento_breve_percentuale",
                   "grado_copertura_interessi_percentuale", "indice_indipendenza_finanziaria_percentuale",
                   "rapporto_indebitamento", "oneri_finanziari_fatt_percentuale")
# Vedi tutte le variabili originali selezionate
names(dataset_z_score)

# --- Funzione per stima pesi LM e LASSO ---
stima_pesi <- function(data, y_var, x_vars) {
  # Linear model
  formula <- as.formula(paste(y_var, "~", paste(x_vars, collapse = "+")))
  lm_mod <- lm(formula, data=data)
  coefs_lm <- coef(lm_mod)[x_vars]
  coefs_lm_norm <- coefs_lm / sum(abs(coefs_lm))
  ISP_lm <- as.matrix(data[, x_vars]) %*% coefs_lm_norm
  
  # LASSO
  X <- model.matrix(formula, data)[, -1]
  y <- data[[y_var]]
  cv_lasso <- cv.glmnet(X, y, alpha=1)
  best_lambda <- cv_lasso$lambda.min
  model_lasso <- glmnet(X, y, alpha=1, lambda=best_lambda)
  coefs_lasso <- as.vector(coef(model_lasso))[-1]
  coefs_lasso_norm <- coefs_lasso / sum(abs(coefs_lasso))
  ISP_lasso <- as.matrix(data[, x_vars]) %*% coefs_lasso_norm
  
  list(
    lm_mod = lm_mod,
    coefs_lm_norm = coefs_lm_norm,
    ISP_lm = ISP_lm,
    model_lasso = model_lasso,
    coefs_lasso_norm = coefs_lasso_norm,
    ISP_lasso = ISP_lasso,
    r2_lm = summary(lm_mod)$r.squared
  )
}

# --- Stima pesi e ISP gruppo A ---
res_a <- stima_pesi(dataset_z_score, "ISP", vars_a)
ISP_a_lm_norm <- normalize_1000(res_a$ISP_lm)
ISP_a_lasso_norm <- normalize_1000(res_a$ISP_lasso)

# --- Stima pesi e ISP gruppo B base ---
res_b <- stima_pesi(dataset_z_score, "ISP", vars_b)
ISP_b_lm_norm <- normalize_1000(res_b$ISP_lm)
ISP_b_lasso_norm <- normalize_1000(res_b$ISP_lasso)

# --- Stima pesi e ISP gruppo B esteso ---
res_b_ext <- stima_pesi(dataset_z_score, "ISP", vars_b_esteso)
ISP_b_ext_lm_norm <- normalize_1000(res_b_ext$ISP_lm)
ISP_b_ext_lasso_norm <- normalize_1000(res_b_ext$ISP_lasso)

# --- Calcolo pesi combinati basati su R² ---
peso_a <- res_a$r2_lm / (res_a$r2_lm + res_b_ext$r2_lm)
peso_b <- res_b_ext$r2_lm / (res_a$r2_lm + res_b_ext$r2_lm)

# --- ISP totali combinati gruppo B ext ---
ISP_tot_lm_ext <- peso_a * ISP_a_lm_norm + peso_b * ISP_b_ext_lm_norm
ISP_tot_lasso_ext <- peso_a * ISP_a_lasso_norm + peso_b * ISP_b_ext_lasso_norm

# --- ISP totali combinati  ---
ISP_tot_lm <- peso_a * ISP_a_lm_norm + peso_b * ISP_b_lm_norm
ISP_tot_lasso <- peso_a * ISP_a_lasso_norm + peso_b * ISP_b_lasso_norm

# --- Visualizzazione confronti ---
par(mfrow=c(2,2))
hist(ISP_tot_lm, main="ISP Totale - Linear Model", col="blue", xlab="ISP LM Normalizzato")
hist(ISP_tot_lasso, main="ISP Totale - LASSO", col="green", xlab="ISP LASSO Normalizzato")
hist(ISP_tot_lm_ext, main="ISP Totale - Linear Model", col="lightblue", xlab="ISP LM Normalizzato ext")
hist(ISP_tot_lasso_ext, main="ISP Totale - LASSO", col="lightgreen", xlab="ISP LASSO Normalizzato ext")
par(mfrow=c(1,1))

# --- Output riepilogativo ---
cat("R² gruppo A (LM):", round(res_a$r2_lm, 4), "\n")
cat("R² gruppo B esteso (LM):", round(res_b_ext$r2_lm, 4), "\n")
cat("Peso gruppo A:", round(peso_a, 4), "Peso gruppo B:", round(peso_b, 4), "\n")

# CONFRONTO NUMERICO
# Aggiungo i vettori ISP_tot_lm e ISP_tot_lasso normalizzati e standardizzati al dataframe

ISP_tot_lm_std <- scale(ISP_tot_lm)[,1]        # standardizzazione z-score (media=0, sd=1)
ISP_tot_lasso_std <- scale(ISP_tot_lasso)[,1]
ISP_tot_lm_std_ext <- scale(ISP_tot_lm_ext)[,1]        # standardizzazione z-score (media=0, sd=1)
ISP_tot_lasso_std_ext <- scale(ISP_tot_lasso_ext)[,1]

# Creo il dataframe di confronto
df_isp <- data.frame(
  ISP = dataset_completo$ISP,
  ISP_pca_norm = ISP_totale_pca,
  ISP_lm_norm = ISP_totale_lm,
  ISP_lm_ext_norm = ISP_tot_lm_ext,
  ISP_lm_tot_norm = ISP_tot_lm,
  ISP_lasso_ext_norm = ISP_tot_lasso,
  ISP_std = dataset_z_score$ISP,
  ISP_pca_std = ISP_totale_pca_std,
  ISP_lm_std = ISP_totale_lm_std,
  ISP_lm_ext_std = ISP_tot_lm_std_ext,
  ISP_lm_tot_std = ISP_tot_lm_std,
  ISP_lasso_ext_std = ISP_tot_lasso_std
)

# Riepilogo statistico sintetico
summary(df_isp)

# Visualizzazione grafica comparativa (optional)
par(mfrow=c(1,4))
hist(df_isp$ISP_pca_norm, main="ISP PCA Normalizzato", col="skyblue", xlab="ISP PCA Norm")
hist(df_isp$ISP_lm_norm, main="ISP LM Normalizzato", col="lightgreen", xlab="ISP LM Norm")
hist(df_isp$ISP_lasso_ext_norm, main="ISP LASSO Normalizzato", col="salmon", xlab="ISP LASSO Norm")
hist(df_isp$ISP, main="ISP Base Normalizzato", col="lightgray", xlab="ISP Base Norm")
par(mfrow=c(1,1))
# Visualizzazione grafica comparativa aggiornata
par(mfrow = c(2, 4), mar = c(4, 4, 2, 1))

hist(df_isp$ISP_pca_norm, main = "ISP PCA (norm)", col = "skyblue", xlab = "", breaks = 30)
hist(df_isp$ISP_lm_norm, main = "ISP LM base (norm)", col = "lightgreen", xlab = "", breaks = 30)
hist(df_isp$ISP_lm_ext_norm, main = "ISP LM esteso (norm)", col = "khaki", xlab = "", breaks = 30)
hist(df_isp$ISP_lasso_ext_norm, main = "ISP LASSO esteso (norm)", col = "salmon", xlab = "", breaks = 30)

hist(df_isp$ISP_pca_std, main = "ISP PCA (std)", col = "steelblue", xlab = "", breaks = 30)
hist(df_isp$ISP_lm_std, main = "ISP LM base (std)", col = "forestgreen", xlab = "", breaks = 30)
hist(df_isp$ISP_lm_ext_std, main = "ISP LM esteso (std)", col = "goldenrod", xlab = "", breaks = 30)
hist(df_isp$ISP_lasso_ext_std, main = "ISP LASSO esteso (std)", col = "firebrick", xlab = "", breaks = 30)

par(mfrow = c(1, 1))  # Reset layout

##### Analisi settoriale con LASSO e LM ####
results_settoriali <- list()
settori <- unique(dataset_z_score$sezione)

for (settore in unique(dataset_z_score$sezione)) {
  dati <- subset(dataset_z_score, sezione == settore)
  dati_test <- dati[, unique(c("ISP", vars_a, vars_b_esteso))]
  cat(settore, ": totale =", nrow(dati), 
      " -> complete.cases =", sum(complete.cases(dati_test)), "\n")
}

for (settore in settori) {
  dati <- subset(dataset_z_score, sezione == settore)
  
  if (nrow(dati) < 30) {
    message(paste("Settore", settore, "saltato: solo", nrow(dati), "osservazioni"))
    next
  }
  
  res_a_sett <- stima_pesi(dati, "ISP", vars_a)
  res_b_sett <- stima_pesi(dati, "ISP", vars_b_esteso)
  
  peso_a_sett <- res_a_sett$r2_lm / (res_a_sett$r2_lm + res_b_sett$r2_lm)
  peso_b_sett <- res_b_sett$r2_lm / (res_a_sett$r2_lm + res_b_sett$r2_lm)
  
  ISP_sett_lm <- peso_a_sett * normalize_1000(res_a_sett$ISP_lm) + peso_b_sett * normalize_1000(res_b_sett$ISP_lm)
  ISP_sett_lasso <- peso_a_sett * normalize_1000(res_a_sett$ISP_lasso) + peso_b_sett * normalize_1000(res_b_sett$ISP_lasso)
  
  results_settoriali[[settore]] <- list(
    ISP_sett_lm = ISP_sett_lm,
    ISP_sett_lasso = ISP_sett_lasso,
    peso_a_sett = peso_a_sett,
    peso_b_sett = peso_b_sett,
    r2_a = res_a_sett$r2_lm,
    r2_b = res_b_sett$r2_lm
  )
}

# Costruzione tabella riassuntiva dei risultati settoriali
tabella_risultati <- do.call(rbind, lapply(names(results_settoriali), function(settore) {
  r <- results_settoriali[[settore]]
  data.frame(
    sezione = settore,
    r2_gruppo_a = round(r$r2_a, 4),
    r2_gruppo_b = round(r$r2_b, 4),
    peso_gruppo_a = round(r$peso_a_sett, 4),
    peso_gruppo_b = round(r$peso_b_sett, 4)
  )
}))

# Visualizza i risultati
print(tabella_risultati)

# Facoltativo: salva su file
write.csv(tabella_risultati, "risultati_settoriali_pesi.csv", row.names = FALSE)

# Riorganizza per plot
tabella_plot <- pivot_longer(tabella_risultati, 
                             cols = starts_with("peso_gruppo"), 
                             names_to = "gruppo", 
                             values_to = "peso")

# Rinomina gruppi per leggibilità
tabella_plot$gruppo <- factor(tabella_plot$gruppo, 
                              levels = c("peso_gruppo_a", "peso_gruppo_b"),
                              labels = c("Redditività/Performance", "Patrimoniale/Finanziario"))

# Ordina le sezioni in base al peso del gruppo A (decrescente)
ordine_sezioni <- tabella_risultati %>%
  arrange(desc(peso_gruppo_a)) %>%
  pull(sezione)

tabella_plot$sezione <- factor(tabella_plot$sezione, levels = ordine_sezioni)

# Grafico
ggplot(tabella_plot, aes(x = sezione, y = peso, fill = gruppo)) +
  geom_bar(stat = "identity") +
  coord_flip() +
  labs(title = "Pesi stimati per gruppo di variabili per settore (ordinati per peso A decrescente)",
       x = "Sezione ATECO",
       y = "Peso stimato",
       fill = "Gruppo") +
  theme_minimal()


for (settore in names(results_settoriali)) {
  index <- dataset_z_score$sezione == settore
  dati_settore <- subset(dataset_z_score, sezione == settore)
  
  ISP_lm <- results_settoriali[[settore]]$ISP_sett_lm
  ISP_lasso <- results_settoriali[[settore]]$ISP_sett_lasso
  
  dataset_z_score$ISP_sett_lm[index] <- ISP_lm
  dataset_z_score$ISP_sett_lasso[index] <- ISP_lasso
}
ls()
# Standardizzazione z-score dei modelli settoriali
ISP_sett_lm_std <- scale(dataset_z_score$ISP_sett_lm)[,1]
ISP_sett_lasso_std <- scale( dataset_z_score$ISP_sett_lasso)[,1]

# Aggiunta al dataframe di confronto
df_isp$ISP_sett_lm_norm <- dataset_z_score$ISP_sett_lm
df_isp$ISP_sett_lm_std <- ISP_sett_lm_std
df_isp$ISP_sett_lasso_norm <-  dataset_z_score$ISP_sett_lasso
df_isp$ISP_sett_lasso_std <- ISP_sett_lasso_std

summary(df_isp)

# Visualizzazione grafica degli ulterio ISP
par(mfrow = c(3, 4), mar = c(4, 4, 2, 1))

hist(df_isp$ISP_sett_lm_norm, main = "ISP LM sett (norm)", col = "lightblue", xlab = "", breaks = 30)
hist(df_isp$ISP_sett_lasso_norm, main = "ISP LASSO sett (norm)", col = "lightsalmon", xlab = "", breaks = 30)
hist(df_isp$ISP_sett_lm_std, main = "ISP LM sett (std)", col = "dodgerblue", xlab = "", breaks = 30)
hist(df_isp$ISP_sett_lasso_std, main = "ISP LASSO sett (std)", col = "indianred", xlab = "", breaks = 30)

hist(df_isp$ISP_pca_norm, main = "ISP PCA (norm)", col = "skyblue", xlab = "", breaks = 30)
hist(df_isp$ISP_lm_norm, main = "ISP LM base (norm)", col = "lightgreen", xlab = "", breaks = 30)
hist(df_isp$ISP_pca_std, main = "ISP PCA (std)", col = "steelblue", xlab = "", breaks = 30)
hist(df_isp$ISP_lm_std, main = "ISP LM base (std)", col = "forestgreen", xlab = "", breaks = 30)

hist(df_isp$ISP_lm_ext_norm, main = "ISP LM esteso (norm)", col = "khaki", xlab = "", breaks = 30)
hist(df_isp$ISP_lasso_ext_norm, main = "ISP LASSO esteso (norm)", col = "salmon", xlab = "", breaks = 30)
hist(df_isp$ISP_lm_ext_std, main = "ISP LM esteso (std)", col = "goldenrod", xlab = "", breaks = 30)
hist(df_isp$ISP_lasso_ext_std, main = "ISP LASSO esteso (std)", col = "firebrick", xlab = "", breaks = 30)

par(mfrow = c(1, 1))

isp_mat <- df_isp[, grep("_std$", names(df_isp))]
cor_isp <- cor(isp_mat, use = "complete.obs", method = "spearman")
corrplot(cor_isp, method = "color", type = "upper", tl.cex = 0.8)

# inversione di PCA
df_isp$ISP_pca_norm<-1000-df_isp$ISP_pca_norm
df_isp$ISP_pca_std<-scale(df_isp$ISP_pca_norm)
summary(df_isp)

# CORRELAZIONI
cor(df_isp, use = "complete.obs")

df_isp_norm <- df_isp[, grep("_norm$", names(df_isp))]
df_isp_norm$ISP <- df_isp$ISP
# Calcola la matrice dei rank decrescenti
rank_matrix <- apply(df_isp_norm, 2, function(x) rank(-x))

# Usa solo le colonne dei ranking
rank_matrix <- apply(df_isp_norm, 2, function(x) rank(-x))
kendall_result <- kendall(rank_matrix)
print(kendall_result)

# Aggiungi media e deviazione standard del rank per ciascuna impresa
rank_summary <- as.data.frame(rank_matrix)
rank_summary$media_rank <- rowMeans(rank_matrix)
rank_summary$sd_rank <- apply(rank_matrix, 1, sd)
rank_summary$media_rank <- apply(rank_matrix, 1, mean, na.rm = TRUE)

# Definiamo soglie
top_soglia <- 1000   # top 100 aziende su tutte
outlier_sd_soglia <- 5000  # soglia empirica per forte distorsione

# Filtra top performer e outlier
top_aziende <- rank_summary[rank_summary$media_rank <= top_soglia, ]
outlier_aziende <- rank_summary[rank_summary$sd_rank >= outlier_sd_soglia, ]

# Visualizza
cat("== Migliori aziende (rank medio <= 100) ==\n")
print(head(top_aziende[order(top_aziende$media_rank), ], 10))

cat("\n== Aziende distorte (sd rank >= 5000) ==\n")
print(head(outlier_aziende[order(-outlier_aziende$sd_rank), ], 10))

# Valutazione Mappa rischio-performance

ggplot(rank_summary, aes(x = sd_rank, y = media_rank)) +
  geom_point(alpha = 0.5) +
  geom_hline(yintercept = median(rank_summary$media_rank), linetype = "dashed", color = "blue") +
  geom_vline(xintercept = median(rank_summary$sd_rank), linetype = "dashed", color = "red") +
  labs(title = "Mappa Rischio-Performance delle Imprese",
       x = "Incertezza tra ISP (SD rank)",
       y = "Posizione media (rank)") +
  theme_minimal()

# Clusterizzazione
kmeans_input <- rank_summary %>% select(media_rank, sd_rank)
set.seed(2025)
kmeans_input<-scale(kmeans_input)
cluster <- kmeans(kmeans_input, centers = 4)
rank_summary$cluster <- cluster$cluster

# Visualizzare i cluster nella mappa rischio-performance
ggplot(rank_summary, aes(x = sd_rank, y = media_rank, color = as.factor(cluster))) +
  geom_point(alpha = 0.6, size = 2) +
  geom_hline(yintercept = median(rank_summary$media_rank), linetype = "dashed", color = "blue") +
  geom_vline(xintercept = median(rank_summary$sd_rank), linetype = "dashed", color = "red") +
  labs(
    title = "Mappa Rischio-Performance con Cluster",
    x = "Deviazione Standard dei Rank (Incertezza)",
    y = "Media dei Rank (Performance)",
    color = "Cluster"
  ) +
  theme_minimal()


ggplot(rank_summary, aes(x = media_rank, y = sd_rank, color = factor(cluster))) +
  geom_point(alpha = 0.7, size = 3) +
  geom_vline(xintercept = median(rank_summary$media_rank), linetype = "dashed") +
  geom_hline(yintercept = median(rank_summary$sd_rank), linetype = "dashed") +
  scale_color_brewer(palette = "Set1") +
  labs(
    title = "Cluster di imprese su media e deviazione dei ranking",
    x = "Media ranking (performance)",
    y = "Deviazione standard (incertezza)",
    color = "Cluster"
  ) +
  theme_minimal()

# Seleziona e ristruttura i dati per i boxplot
df_plot <- df_isp %>%
  select(ends_with("_norm"), ends_with("_std")) %>%
  pivot_longer(
    everything(),
    names_to = "versione",
    values_to = "valore"
  ) %>%
  mutate(
    tipo = ifelse(str_detect(versione, "_norm$"), "Normalizzato", "Standardizzato"),
    versione = str_remove(versione, "_norm$|_std$")
  )

# 2. Boxplot con facet per tipo
ggplot(df_plot, aes(x = versione, y = valore, fill = versione)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  coord_flip() +
  facet_wrap(~ tipo, scales = "free_x") +
  labs(
    title = "Boxplot per versione normalizzata e standardizzata",
    x = "Versione ISP",
    y = "Valore",
    fill = "Versione"
  ) +
  theme_minimal(base_size = 12) +
  theme(legend.position = "none")

# Sostituzione della variabile dipendente con la versione settoriale 
# in quanto specifica della multisettorialità della matrice
dataset_completo$ISP <- df_isp$ISP_sett_lasso_norm

# Aggiorno dataset_completo
saveRDS(dataset_completo, "matrice.rds")

##### Indicatore di integrazione verticale à la buzzell####
dataset_completo %>%
  group_by(divisione) %>%
  summarise(
    ROA_medio = mean(ROA_percentuale, na.rm = TRUE),
    ROA_mediano = median(ROA_percentuale, na.rm = TRUE),
    sd_ROA = sd(ROA_percentuale, na.rm = TRUE),
    n = n()
  )

# Calcolo del ROA medio per settore
dataset_completo <- dataset_completo %>%
  group_by(divisione) %>%  # Usa direttamente la colonna "settore" (ATECO 2 digit)
  mutate(ROA_medio_settore = mean(ROA_percentuale, na.rm = TRUE)) %>%  
  ungroup()
# Sostituzione alla variabile dipendenti valore 0 con 1
dataset_completo$dipendenti[dataset_completo$dipendenti == 0] <- 1
# Calcolare il valore aggiunto in migl di EUR
dataset_completo <- dataset_completo %>%
  mutate(valore_aggiunto_migl_EUR = (valore_aggiunto_pro_capite_EUR * dipendenti) / 1e3)

# Calcolo dell'indicatore di integrazione verticale à la Buzzell
dataset_completo <- dataset_completo %>%
  mutate(
    numeratore = valore_aggiunto_migl_EUR - utile_netto_migl_EUR + (ROA_medio_settore * totale_attivita_migl_EUR),
    denominatore = ricavi_vendite_migl_EUR - utile_netto_migl_EUR + (ROA_medio_settore * totale_attivita_migl_EUR),
    integrazione_verticale = numeratore / denominatore
  ) %>%
  mutate(
    integrazione_verticale = ifelse(is.nan(integrazione_verticale) | is.infinite(integrazione_verticale), NA, integrazione_verticale)
  )
# aggiunta di integrazione-verticale dataset_norm
dataset_normalizzato$integrazione_verticale <- (dataset_completo$integrazione_verticale - 
                                                         min(dataset_completo$integrazione_verticale, na.rm = TRUE)) / 
  (max(dataset_completo$integrazione_verticale, na.rm = TRUE) - min(dataset_completo$integrazione_verticale, na.rm = TRUE))


# Esportazione dei risultati
write.xlsx(dataset_completo, "dataset_con_integrazione_verticale_completo.xlsx")
saveRDS(dataset_completo, "matrice.rds")

#visualizzazione
dataset_completo <- dataset_completo[order(dataset_completo$integrazione_verticale, decreasing = TRUE), ]

int<-dataset_completo %>%
  filter(integrazione_verticale > 1) %>%
  dplyr::select(ragione_sociale, integrazione_verticale, dimensione_impresa, macroarea, sezione) %>%
  head(6)
# Filtro e selezione colonne
osservazioni_filtrate <- dataset_completo %>%
  filter(integrazione_verticale > 1) %>%
  dplyr::select(ragione_sociale, integrazione_verticale, dimensione_impresa, macroarea, sezione)

# Tabella interattiva
datatable(osservazioni_filtrate, 
          options = list(pageLength = 10, scrollX = TRUE),
          caption = "Imprese con integrazione verticale > 1")
par(mfrow=c(1,3))
hist(dataset_completo$integrazione_verticale)
boxplot(dataset_completo$integrazione_verticale)
summary(dataset_completo$integrazione_verticale)
qqnorm(dataset_completo$integrazione_verticale, main = "Q-Q Plot: Integrazione Verticale")
qqline(dataset_completo$integrazione_verticale, col = "red", lwd = 2)

#confronti tra indicatori
par(mfrow = c(1, 1))

plot(dataset_completo$ISP, dataset_completo$integrazione_verticale,
     xlab = "ISP", ylab = "Integrazione verticale",
     main = "Scatterplot: ISP vs Integrazione verticale")

# Salvataggio dell'istogramma per l'Indice Sintetico
p1 <- ggplot(dataset_completo, aes(x = ISP)) +
  geom_histogram(binwidth = 0.05, fill = "steelblue", color = "black", alpha = 0.7) +
  labs(title = "Distribuzione dell'Indice Sintetico", x = "Indice Sintetico", y = "Frequenza")
ggsave("distribuzione_indice_sintetico.png", plot = p1, width = 8, height = 6)

# Salvataggio dell'istogramma per l'Integrazione Verticale
p2 <- ggplot(dataset_completo, aes(x = integrazione_verticale)) +
  geom_histogram(binwidth = 0.05, fill = "green", color = "black", alpha = 0.7) +
  labs(title = "Distribuzione Integrazione Verticale", x = "Integrazione Verticale", y = "Frequenza")
ggsave("distribuzione_integrazione_verticale.png", plot = p2, width = 8, height = 6)

# Boxplot per l'Indice Sintetico per Settore
p3 <- ggplot(dataset_completo, aes(x = divisione, y = ISP)) +
  geom_boxplot() +
  labs(title = "Distribuzione dell'Indice Sintetico per Settore", x = "Settore", y = "Indice Sintetico")
ggsave("boxplot_indice_sintetico_per_settore.png", plot = p3, width = 10, height = 6)

# Boxplot per l'Integrazione Verticale per Settore
p4 <- ggplot(dataset_completo, aes(x = divisione, y = integrazione_verticale)) +
  geom_boxplot() +
  labs(title = "Distribuzione dell'Integrazione Verticale per Settore", x = "Settore", y = "Integrazione Verticale")
ggsave("boxplot_integrazione_verticale_per_settore.png", plot = p4, width = 10, height = 6)

# Test di correlazione tra Indice Sintetico e Integrazione Verticale
cor_test_result <- cor.test(scale(dataset_completo$ISP),scale( dataset_completo$integrazione_verticale), use = "complete.obs")
# Estrazione dei risultati dal test di correlazione
cor_test_result_df <- data.frame(
  correlazione = cor_test_result$estimate,
  p_value = cor_test_result$p.value,
  conf_int_lower = cor_test_result$conf.int[1],
  conf_int_upper = cor_test_result$conf.int[2]
)

# Calcolo e salvataggio della heatmap delle correlazioni
cor_matrix <- cor(dataset_completo[, c("ISP", "integrazione_verticale", "valore_aggiunto_migl_EUR", "ROA_percentuale")], use = "complete.obs")
heatmap_data <- melt(cor_matrix)

p5 <- ggplot(heatmap_data, aes(Var1, Var2, fill = value)) +
  geom_tile() +
  scale_fill_gradient2(low = "blue", mid = "white", high = "red", midpoint = 0) +
  labs(title = "Heatmap delle Correlazioni")
ggsave("heatmap_correlazioni.png", plot = p5, width = 8, height = 6)
##### Statistiche descrittive & trasformazioni####
dataset_completo<-readRDS("matrice.rds")

# Controllo della struttura
print(categorie_variabili)
par(mfrow = c(2, 3))
variabili <- c("ISP","integrazione_verticale")

for (variabile in variabili) {
  cat("\n### Variabile:", variabile, "\n")
  
  x <- na.omit(dataset_completo[[variabile]])
  
  
  unique_vals <- unique(x)
  is_discrete <- length(unique_vals) <= 10 && all(unique_vals == floor(unique_vals))
  
  if (is_discrete) {
    breaks <- seq(min(x) - 0.5, max(x) + 0.5, by = 1)
  } else {
    breaks <- "FD"
  }
  
  h <- hist(x,
            breaks = breaks,
            freq = TRUE,
            col = "lightblue",
            border = "white",
            main = paste("Istogramma di", variabile),
            xlab = variabile)
  
  if (!is_discrete) {
    d <- density(x)
    bin_width <- if (is.numeric(h$breaks)) diff(h$breaks[1:2]) else 1
    y_scaled <- d$y * length(x) * bin_width
    lines(d$x, y_scaled, col = "red", lwd = 2)
    legend("topright", legend = "Densità (scala conteggi)", col = "red", lwd = 2, bty = "n")
  }
  
  boxplot(x,
          horizontal = TRUE,
          main = paste("Boxplot di", variabile),
          col = "orange")
  
  qqnorm(x, main = paste("QQ Plot di", variabile))
  qqline(x, col = "blue", lwd = 2)
}

for (categoria in names(categorie_variabili)) {
  cat("\n## Categoria:", categoria, "\n")
  
  for (variabile in categorie_variabili[["Capitale_Circolante"]]) {
    cat("\n### Variabile:", variabile, "\n")
    
    x <- dataset_completo[[variabile]]
    
    
    # 1. Riconosci variabili discrete (<= 10 valori unici interi)
    unique_vals <- unique(x)
    is_discrete <- length(unique_vals) <= 10 && all(unique_vals == floor(unique_vals))
    
    if (is_discrete) {
      breaks <- seq(min(x) - 0.5, max(x) + 0.5, by = 1)
    } else {
      breaks <- "FD"  # Freedman–Diaconis
    }
    
    # 2. Istogramma
    h <- hist(x,
              breaks = breaks,
              freq = TRUE,
              col = "lightblue",
              border = "white",
              main = paste("Istogramma di", variabile),
              xlab = variabile)
    
    # 3. Densità: solo se variabile continua
    if (!is_discrete) {
      d <- density(x)
      bin_width <- if (is.numeric(h$breaks)) diff(h$breaks[1:2]) else 1
      y_scaled <- d$y * length(x) * bin_width  # Scala la densità
      
      lines(d$x, y_scaled, col = "red", lwd = 2)
      legend("topright", legend = "Densità (scala conteggi)", col = "red", lwd = 2, bty = "n")
    }
    
    # 4. Boxplot
    boxplot(x,
            horizontal = TRUE,
            main = paste("Boxplot"),
            col = "orange")
    
    # 5. QQ plot
    qqnorm(x)
    qqline(x, col = "blue", lwd = 2)
  }
}

summary(dataset_completo$ISP)

sum(is.na(dataset_completo$ISP))

dataset_completo <- na.omit(dataset_completo)

bn_result <- bestNormalize(dataset_completo$ISP, standardize = FALSE)
dataset_completo$ISP_bn <- predict(bn_result)
print(bn_result)

par(mfrow = c(1, 3))
hist(dataset_completo$ISP, main = "Originale", xlab = "ISP")
hist(dataset_completo$ISP_bn, main = "OrderNorm", xlab = "ISP ordernorm")

qqnorm(dataset_completo$ISP_bn)
qqline(dataset_completo$ISP_bn, col = "red")
summary(dataset_completo$ISP)
sd(dataset_completo$ISP_bn)
skewness(dataset_completo$ISP)   # asimmetria
kurtosis(dataset_completo$ISP)   # curtosi (eccesso)
##### Matrice correlazione tra variabili e nei settori####
# Ripristina le impostazioni grafiche
par(mfrow = c(1, 1)) 
categorie_variabili$Performance <- c("ISP_bn", "integrazione_verticale")
# Standardizza le variabili selezionate nel dataset (per ogni categoria)
dataset_z_score <- as.data.frame(scale(dataset_completo[, unlist(categorie_variabili)]))
# Matrice di correlazione
cor_matrix <- cor(dataset_z_score[, unlist(categorie_variabili)])
# Crea la heatmap della matrice di correlazione
pheatmap(cor_matrix, 
         clustering_distance_rows = "euclidean", 
         clustering_distance_cols = "euclidean", 
         clustering_method = "complete", 
         display_numbers = TRUE, 
         main = "Matrice di Correlazione")

#aggiunta della colonna settore
dataset_normalizzato <- dataset_normalizzato %>%               # 25 167 righe
  left_join(
    dataset_completo %>%                                       # 25 148 righe
      select(codice_fiscale, sezione) %>%                      # chiave + variabile da ereditare
      distinct(),                                              # precauzione: niente duplicati
    by = "codice_fiscale"
  )
dataset_z_score$sezione <- dataset_completo$sezione  # Aggiunta la colonna 

# Filtra i gruppi con almeno 2 osservazioni per il calcolo della correlazione
cor_settori_interesse <- dataset_z_score %>%
  group_by(sezione) %>%
  filter(n() > 2) %>%  # Mantieni solo i gruppi con più di 2 osservazione
  summarise(
    cor_matrix = list(
      cor(across(where(is.numeric)), use = "pairwise.complete.obs")
    ),
    .groups = "drop"
  )

# Visualizza il risultato
print(cor_settori_interesse)



##### Controllo la significatività delle correlazioni tra variabili e settore####
# Lista delle variabili da testare
vars_model <- unname(unlist(categorie_variabili))

# Funzione principale
confronto_settori_variabili <- function(dataset, variabili, variabile_settore = "sezione", soglia_normalita = 0.05, soglia_significativita = 0.05) {
  
  risultati <- map_dfr(variabili, function(var) {
    
    # Costruisci dataset temporaneo
    dati_var <- dataset %>%
      select(all_of(variabile_settore), !!sym(var)) %>%
      filter(!is.na(.[[var]])) %>%
      group_by(!!sym(variabile_settore)) %>%
      mutate(n_settore = n()) %>%
      ungroup() %>%
      filter(n_settore >= 3)  # il test di Shapiro richiede almeno 3 osservazioni
    
    # Se ci sono abbastanza dati
    if (n_distinct(dati_var[[variabile_settore]]) < 2) {
      return(tibble(variabile = var, test = NA, p_value = NA, significativo = NA))
    }
    
    # Test di normalità per ogni settore
    normalita <- dati_var %>%
      group_by(!!sym(variabile_settore)) %>%
      summarise(p_shapiro = tryCatch(shapiro.test(.[[var]])$p.value, error = function(e) NA), .groups = "drop")
    
    tutti_normali <- all(normalita$p_shapiro > soglia_normalita, na.rm = TRUE)
    
    # Test globale (ANOVA o Kruskal-Wallis)
    formula_test <- as.formula(paste(var, "~", variabile_settore))
    test_risultato <- tryCatch({
      if (tutti_normali) {
        res <- aov(formula_test, data = dati_var)
        p_val <- summary(res)[[1]][["Pr(>F)"]][1]
        list(test = "ANOVA", p_value = p_val)
      } else {
        res <- kruskal.test(formula_test, data = dati_var)
        list(test = "Kruskal-Wallis", p_value = res$p.value)
      }
    }, error = function(e) list(test = NA, p_value = NA))
    
    # Ritorna riga di risultati
    tibble(
      variabile = var,
      test = test_risultato$test,
      p_value = test_risultato$p_value,
      significativo = ifelse(!is.na(test_risultato$p_value) & test_risultato$p_value < soglia_significativita, "Significativo", "Non significativo")
    )
  })
  
  return(risultati)
}

#  Applicazione funzione
risultati_test_settore <- confronto_settori_variabili(dataset = dataset_completo, variabili = vars_model)

# Esporta in Excel
write.xlsx(risultati_test_settore, "risultati_test_settore.xlsx")

# Visualizza risultati significativi
risultati_test_settore %>%
  filter(significativo == "Significativo") %>%
  ggplot(aes(x = reorder(variabile, p_value), y = p_value, fill = test)) +
  geom_col() +
  coord_flip() +
  geom_hline(yintercept = 0.05, linetype = "dashed", color = "red") +
  labs(title = "Variabili con differenze significative tra settori",
       x = "Variabile", y = "p-value", fill = "Test usato") +
  theme_minimal()


##### Visualizzazione grafica dei dati ####
## Creazione Mappa Iterattiva 
# Verifico se ci sono NA nelle colonne latitudine e longitudine
dataset_map <- dataset_completo[!is.na(dataset_completo$longitudine) & !is.na(dataset_completo$latitudine), ]
attach(dataset_map)

breaks <- c(min(ISP_bn), mean(ISP_bn), max(ISP_bn))  # Classi personalizzate

# Creiamo una palette con transizione graduale tra i colori
pal <- colorBin(
  palette = c("darkblue", "blue", "white", "pink", "darkred"), 
  domain = dataset_map$ISP_bn, 
  bins = breaks
)
# Crea la mappa con i marker colorati in base ai ricavi, e popup con informazioni aggiuntive
mappa <- leaflet(dataset_map) %>%
  addTiles() %>%
  addCircleMarkers(
    lng = ~longitudine, 
    lat = ~latitudine, 
    radius = ISP_bn, 
    color = ~pal(ISP_bn),  # Applica la palette di colori 
    stroke = FALSE, 
    fillOpacity = 0.8,
    popup = ~paste(
      "Azienda: ", ragione_sociale, "<br>", 
      "Ricavi: ", ricavi_vendite_migl_EUR, " Migliaia EUR", "<br>",
      "EBITDA: ", EBITDA_migl_EUR, " Migliaia EUR", "<br>",
      "Indice sintetico Performance: ", ISP_bn, "<br>",
      "Settore: ", ateco_2007_descrizione  
    )
  ) %>%
  addLegend("bottomright", pal = pal, values = ~ricavi_vendite_migl_EUR,
            title = "ISP", opacity = 1)

# Visualizza la mappa
mappa

breaks <- c(min(utile_netto_migl_EUR), -5000, -1000, -100, 0, 100, 1000, 5000, max(utile_netto_migl_EUR))  # Classi personalizzate

# Creiamo una palette con transizione graduale tra i colori
pal <- colorBin(
  palette = c("darkblue", "blue", "lightblue", "white", "pink", "red", "darkred"), 
  domain = dataset_map$utile_netto_migl_EUR, 
  bins = breaks
)
# Creiamo la mappa
mappa <- leaflet(dataset_map) %>%
  addTiles() %>%
  addCircleMarkers(
    lng = ~longitudine, 
    lat = ~latitudine, 
    radius = 1,  # Aumentato per visibilità
    color = ~pal(utile_netto_migl_EUR),  # Colore in base all'utile netto
    stroke = FALSE, 
    fillOpacity = 0.8,
    popup = ~paste(
      "Azienda: ", ragione_sociale, "<br>", 
      "Ricavi: ", ricavi_vendite_migl_EUR, " Migliaia EUR", "<br>",
      "EBITDA: ", EBITDA_migl_EUR, " Migliaia EUR", "<br>",
      "Utile Netto: ", utile_netto_migl_EUR, " Migliaia EUR", "<br>",  # Qui ho corretto il nome della variabile
      "Settore: ", ateco_2007_descrizione  
    )
  ) %>%
  addLegend("bottomright", 
            pal = pal, 
            values = dataset_map$utile_netto_migl_EUR,  # Uso l'utile netto per la legenda
            title = "Utile/Perdita (in migliaia EUR)", 
            opacity = 1)

# Mostra la mappa
mappa

# Mappa clustering
# Crea una palette basata su bin
pal_custom <- colorBin(
  palette = "RdYlBu", 
  domain = dataset_map$ricavi_vendite_migl_EUR,
  bins = c(0, 100000, 500000, 1000000, 5000000, max(dataset_map$ricavi_vendite_migl_EUR))  # Limiti personalizzati
)
# Creazione della mappa con i marker colorati per bin e clustering per regione
mappa_bin <- leaflet(dataset_map) %>%
  addTiles() %>%
  addCircleMarkers(
    lng = ~longitudine, 
    lat = ~latitudine, 
    radius = 1,  # Modifica la dimensione dei marker
    color = ~pal_custom(ricavi_vendite_migl_EUR),
    stroke = FALSE, 
    fillOpacity = 0.8,
    popup = ~paste(
      "<b>Azienda: </b>", ragione_sociale, "<br>", 
      "<b>Settore: </b>", ateco_2007_descrizione, "<br>",
      "<b>Ricavi: </b>", ricavi_vendite_migl_EUR, " Migliaia EUR", "<br>",
      "<b>EBITDA: </b>", EBITDA_migl_EUR, " Migliaia EUR", "<br>",
      "<b>Dimensione Aziendale: </b>", dimensione_impresa
    )) %>%
  addLegend("bottomright", pal = pal_custom, values = ~ricavi_vendite_migl_EUR,
            title = "Ricavi (in migliaia EUR)", opacity = 1) %>%
  setView(lng = mean(dataset_map$longitudine, na.rm = TRUE), 
          lat = mean(dataset_map$latitudine, na.rm = TRUE), zoom = 6)

# Visualizza la mappa
mappa_bin

# Rimuovi tutte le variabili tranne dataset, dataset_ridotto e dataset_completo
rm(list = setdiff(ls(), c( "dataset_completo","categorie_variabili","vars_model")))

# Verifica che solo le variabili desiderate siano rimaste
ls()

saveRDS(dataset_completo, "matrice.rds")



##### Preprocessing iniziale #####
dataset_completo<-readRDS("matrice.rds")

# Trasformazioni logaritmiche
dataset_completo <- dataset_completo %>%
  mutate(
    log_ricavi_vendite_migl_EUR = log(ricavi_vendite_migl_EUR + 1),
    log_EBITDA_migl_EUR = log(EBITDA_migl_EUR + 1)
  )
categorie_variabili$Log <- c("log_ricavi_vendite_migl_EUR", "log_EBITDA_migl_EUR")
# Categorizzazione dell'integrazione verticale
dataset_completo <- dataset_completo %>%
  mutate(
    integrazione_verticale_fatt = cut(
      integrazione_verticale,
      breaks = quantile(integrazione_verticale, probs = seq(0, 1, 0.2), na.rm = TRUE),
      include.lowest = TRUE,
      labels = c("Molto Basso", "Basso", "Medio", "Alto", "Molto Alto")
    )
  )

##### Campionamento per la costruzione del modello 
set.seed(2025)
dataset_sample <- sample_n(dataset_completo, 5000)

##### Selezione delle variabili indipendenti #####

variabili_ind_sample <- dataset_sample %>% dplyr::select(all_of(vars_model))
variabili_scaled <- as.data.frame(scale(
  variabili_ind_sample[, setdiff(names(variabili_ind_sample), c("ISP_bn", "EBITDA_migl_EUR", "ricavi_vendite_migl_EUR"))]
))
variabili_scaled$ISP_bn <- variabili_ind_sample$ISP_bn
# Vedi tutte le variabili originali selezionate
names(variabili_ind_sample)

# Vedi tutte le variabili presenti nel dataset scalato
names(variabili_scaled)

setdiff(names(variabili_ind_sample), names(variabili_scaled))
# Riordina se vuoi mantenere l'ordine originale
variabili_scaled <- variabili_scaled[, c("ISP_bn", setdiff(names(variabili_scaled), "ISP_bn"))]
##### Modello stepwise su campione #####
model_null <- lm(ISP_bn ~ 1, data = variabili_scaled)
model_full <- lm(ISP_bn ~ ., data = variabili_scaled)
model_stepwise <- stepAIC(model_null, scope = list(upper = model_full), direction = "both")

##### Output modello stepwise (campione) #####
summary(model_stepwise)
print(vif(model_stepwise))

filter_vif <- function(model, threshold = 5) {
  # Copia della formula e dei dati
  data <- model$model
  formula <- formula(model)
  
  # Estrai variabili indipendenti (escludendo risposta)
  response <- all.vars(formula)[1]
  predictors <- all.vars(formula)[-1]
  
  # Ciclo iterativo di rimozione variabili col VIF alto
  repeat {
    # Fit modello con variabili correnti
    formula_curr <- as.formula(paste(response, "~", paste(predictors, collapse = "+")))
    model_curr <- lm(formula_curr, data = data)
    
    # Calcola VIF
    vifs <- vif(model_curr)
    
    # Controlla se tutti sotto soglia
    if (all(vifs < threshold)) {
      break
    }
    
    # Trova variabile con VIF massimo
    var_to_remove <- names(which.max(vifs))
    cat("Rimuovo variabile per VIF alto:", var_to_remove, " (VIF =", max(vifs), ")\n")
    
    # Rimuovi variabile dai predittori
    predictors <- setdiff(predictors, var_to_remove)
    
    # Se non rimane nulla, esci
    if (length(predictors) == 0) {
      stop("Tutte le variabili rimosse per VIF alto.")
    }
  }
  
  # Ritorna il modello finale e le variabili selezionate
  list(model_final = model_curr, variables = predictors, vifs = vif(model_curr))
}

result <- filter_vif(model_stepwise, threshold = 3)

summary(result$model_final)
print(result$vifs)

remove_high_pval <- function(model, pval_thresh = 0.1) {
  data_curr <- model$model
  formula_curr <- formula(model)
  
  repeat {
    coefs <- summary(model)$coefficients
    coefs_no_intercept <- coefs[rownames(coefs) != "(Intercept)", , drop = FALSE]
    vars_pval_high <- rownames(coefs_no_intercept)[coefs_no_intercept[, "Pr(>|t|)"] > pval_thresh]
    
    if (length(vars_pval_high) > 0) {
      message("Rimuovo per p-value alto: ", paste(vars_pval_high, collapse = ", "))
      vars_curr <- attr(terms(formula_curr), "term.labels")
      vars_new <- setdiff(vars_curr, vars_pval_high)
      
      if (length(vars_new) == 0) {
        message("Attenzione: nessuna variabile rimasta dopo filtro p-value. Esco.")
        break
      }
      
      response_var <- all.vars(formula_curr)[1]
      formula_curr <- reformulate(vars_new, response = response_var)
      model <- lm(formula_curr, data = data_curr)
      next
    }
    break
  }
  
  return(list(model_final = model, formula = formula_curr))
}
result_pval <- remove_high_pval(model = result$model_final, pval_thresh = 0.1)

# Risultato finale
summary(result_pval$model_final)
print(car::vif(result_pval$model_final))
model_reduced<-result_pval$model_final


# Plot diagnostici
png("diagnostic_plots_sample.png", width = 2000, height = 2000, res = 300)
par(mfrow = c(2, 2))
plot(model_reduced)
dev.off()

# Test diagnostici
test_results_sample <- data.frame(
  Test = c("Shapiro-Wilk", "Anderson-Darling", "Breusch-Pagan", "Durbin-Watson"),
  Statistic = c(
    shapiro.test(residuals(model_reduced))$statistic,
    ad.test(residuals(model_reduced))$statistic,
    bptest(model_reduced)$statistic,
    durbinWatsonTest(model_reduced)$dw
  ),
  P_value = c(
    shapiro.test(residuals(model_reduced))$p.value,
    ad.test(residuals(model_reduced))$p.value,
    bptest(model_reduced)$p.value,
    durbinWatsonTest(model_reduced)$p
  )
)

# Metriche del modello
summary_sample <- summary(model_reduced)
model_metrics_sample <- data.frame(
  Metric = c("R-squared", "Adjusted R-squared", "AIC", "BIC", "F-statistic", "Model p-value", "N osservazioni"),
  Value = c(
    round(summary_sample$r.squared, 4),
    round(summary_sample$adj.r.squared, 4),
    round(AIC(model_reduced), 2),
    round(BIC(model_reduced), 2),
    round(summary_sample$fstatistic[1], 2),
    signif(pf(summary_sample$fstatistic[1], summary_sample$fstatistic[2], summary_sample$fstatistic[3], lower.tail = FALSE), 4),
    nobs(model_reduced)
  )
)

##### Tabelle riassuntive (campione) 
kable(tidy(model_reduced), digits = 4, caption = "Coefficiente e p-value del modello stepwise (campione)") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))

kable(test_results_sample, digits = 4, caption = "Risultati dei test diagnostici (campione)") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))

kable(model_metrics_sample, digits = 4, caption = "Statistiche riassuntive del modello stepwise (campione)") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))

##### Controllo Eteroschedasticità
coeftest(model_reduced, vcov. = vcovHC(model_reduced, type = "HC3"))
data_curr <- model_reduced$model   # ricreo l’oggetto con lo stesso nome

mod_hc <- update(model_reduced, . ~ . - debt_equity_ratio_percentuale - debt_EBITDA_ratio_percentuale)

coeftest(mod_hc, vcov. = vcovHC(mod_hc, type = "HC3"))
summary(mod_hc)             # R², Adj R², F-stat
AIC(model_reduced, mod_hc)   # confronto informativo

##### Controllo outlear
train(ISP_bn ~ ., data = mod_hc$model, method = "lm", trControl = trainControl(method = "cv", number = 10))
broom::tidy(mod_hc, conf.int = TRUE, vcov = vcovHC(mod_hc, type = "HC3"))

cooks <- cooks.distance(mod_hc)
which(cooks > 4/length(cooks))    # circa soglia 0.0008 con n=5000

cooks <- cooks.distance(mod_hc)

# 1) Osservazioni con Cook’s > 0.02  (~25× la soglia 4/n)
inf_02  <- which(cooks > 0.02)

# 2) Punti “molto” influenti (Cook’s > 0.04  ≃ 4× mediana classica 0.01)
inf_04  <- which(cooks > 0.04)

length(inf_02);  length(inf_04)      # di solito poche decine e pochi singoli casi

plot(cooks,  pch = 20, main = "Cook's distance (mod_hc)")
abline(h = c(0.02, 0.04), lty = 2, col = "red")
text(which(cooks > 0.04), cooks[cooks > 0.04],
     labels = names(cooks[cooks > 0.04]), pos = 4, cex = .7)

# Rimuovo osservazioni Cook’s > 0.04
mod_trim <- update(mod_hc, subset = cooks <= 0.04)

# Confronto rapido delle stime robuste
coefs_robust <- function(m) broom::tidy(
  m, conf.int = TRUE, vcov = sandwich::vcovHC(m, type = "HC3")
)

bind_rows(
  baseline = coefs_robust(mod_hc),
  trimmed  = coefs_robust(mod_trim),
  .id = "modello"
) |> 
  select(modello, term, estimate, conf.low, conf.high)

mod_final <- update(mod_trim, . ~ . - indice_corrente)

broom::tidy(mod_final, conf.int = TRUE,
            vcov = sandwich::vcovHC(mod_final, type = "HC3")) |>
  kable(digits = 4, caption = "Coefficienti robusti (HC3)") |>
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))


##### Modello stepwise su dataset completo #####
variabili_ind_full <- dataset_completo %>% dplyr::select(all_of(vars_model))
variabili_scaled_full <- as.data.frame(scale(
  variabili_ind_full[, setdiff(names(variabili_ind_sample), c("ISP_bn", "EBITDA_migl_EUR", "ricavi_vendite_migl_EUR"))]
))
variabili_scaled_full$ISP_bn <- variabili_ind_full$ISP_bn

model_null_full <- lm(ISP_bn ~ 1, data = variabili_scaled_full)
model_full_full <- lm(ISP_bn ~ ., data = variabili_scaled_full)
model_stepwise_full <- stepAIC(model_null_full, scope = list(upper = model_full_full), direction = "both")

##### Output modello stepwise (completo) #####
summary(model_stepwise_full)
print(vif(model_stepwise_full))

result_full <- filter_vif(model_stepwise_full, threshold = 3)

summary(result_full$model_final)
print(result_full$vifs)

result_pval_full <- remove_high_pval(model = result_full$model_final, pval_thresh = 0.1)

summary(result_pval_full$model_final)
print(car::vif(result_pval_full$model_final))
model_reduced_full<-result_pval_full$model_final

# Plot diagnostici
png("diagnostic_plots_full.png", width = 2000, height = 2000, res = 300)
par(mfrow = c(2, 2))
plot(model_reduced_full)
dev.off()


# Metriche del modello completo
summary_full <- summary(model_reduced_full)
model_metrics_full <- data.frame(
  Metric = c("R-squared", "Adjusted R-squared", "AIC", "BIC", "F-statistic", "Model p-value", "N osservazioni"),
  Value = c(
    round(summary_full$r.squared, 4),
    round(summary_full$adj.r.squared, 4),
    round(AIC(model_reduced_full), 2),
    round(BIC(model_reduced_full), 2),
    round(summary_full$fstatistic[1], 2),
    signif(pf(summary_full$fstatistic[1], summary_full$fstatistic[2], summary_full$fstatistic[3], lower.tail = FALSE), 4),
    nobs(model_reduced_full)
  )
)

# Tabelle riassuntive (completo) 
kable(tidy(model_reduced_full), digits = 4, caption = "Coefficiente e p-value del modello stepwise (completo)") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))

kable(model_metrics_full, digits = 4, caption = "Statistiche riassuntive del modello stepwise (completo)") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))
##### Confronto tra modelli: campione vs completo #####
model_comparison <- data.frame(
  `Metrica` = model_metrics_sample$Metric,
  `Modello su Campione` = model_metrics_sample$Value,
  `Modello su Dataset Completo` = model_metrics_full$Value
)

kable(model_comparison, digits = 4, caption = "Confronto delle metriche dei modelli stepwise: campione vs completo") %>%
  kable_styling(bootstrap_options = c("striped", "hover", "condensed"))
#####  CONFRONTO COEFFICIENTI #####

coef_campione <- tidy(model_reduced) %>% mutate(Modello = "Campione")
coef_completo <- tidy(model_reduced_full) %>% mutate(Modello = "Completo")

coef_comparison <- bind_rows(coef_campione, coef_completo) %>%
  filter(term != "(Intercept)") %>%
  dplyr::select(Modello, term, estimate, std.error, p.value) %>%
  arrange(term)

kable(coef_comparison, digits = 4, caption = "Confronto dei coefficienti stimati (senza intercetta)") %>%
  kable_styling(bootstrap_options = c("striped", "hover")) %>%
  group_rows("Modello su Campione", 1, nrow(coef_campione) - 1) %>%
  group_rows("Modello su Completo", nrow(coef_campione), nrow(coef_comparison))

##### Analisi per sottogruppo di dimensione #####
# Split per dimensione
dataset_split <- split(dataset_completo, dataset_completo$dimensione_impresa)

# Funzione modello per ciascun subset con selezione stepwise su variabili specifiche
library(car)  # per vif()
library(broom)

modello_dimensionale_stepwise_vif <- function(data) {
  variabili_numeriche <- vars_model[2:length(vars_model)]
  data[variabili_numeriche] <- scale(data[variabili_numeriche])
  
  formula_completa <- as.formula(paste("ISP_bn ~", paste(variabili_numeriche, collapse = " + "), 
                                       "+ as.factor(integrazione_verticale_fatt)*as.factor(macroarea)"))
  modello_stepwise <- step(lm(formula_completa, data = data), direction = "both", trace = 0)
  
  repeat {
    # Calcolo VIF solo per le numeriche
    current_vars <- names(coef(modello_stepwise))[-1]  # tolgo l'intercetta
    numeriche_presenti <- intersect(current_vars, variabili_numeriche)
    
    if (length(numeriche_presenti) < 2) break  # niente VIF da calcolare
    
    formula_vif <- as.formula(paste("ISP_bn ~", paste(numeriche_presenti, collapse = " + ")))
    modello_vif <- lm(formula_vif, data = data)
    
    vifs <- vif(modello_vif)
    
    if (all(vifs < 5)) break  # OK
    
    var_alta_vif <- names(which.max(vifs))
    
    # Rimuovi la variabile con VIF alto dal modello stepwise
    current_terms <- attr(terms(modello_stepwise), "term.labels")
    new_terms <- setdiff(current_terms, var_alta_vif)
    
    if (length(new_terms) == 0) break
    
    nuova_formula <- as.formula(paste("ISP_bn ~", paste(new_terms, collapse = " + ")))
    modello_stepwise <- lm(nuova_formula, data = data)
  }
  
  return(modello_stepwise)
}


# Modelli stimati per ciascuna dimensione
modelli_per_dim_stepwise <- lapply(dataset_split, modello_dimensionale_stepwise_vif)
nomi_dimensioni <- names(modelli_per_dim_stepwise)

# Coefficienti
coef_df_stepwise <- map2_df(modelli_per_dim_stepwise, nomi_dimensioni, ~ tidy(.x) %>% mutate(dimensione_impresa = .y))

# Tabelle riassuntive coefficienti e p-value
coef_wide_stepwise <- coef_df_stepwise %>%
  dplyr::select(term, estimate, dimensione_impresa) %>%
  pivot_wider(names_from = dimensione_impresa, values_from = estimate)

pval_wide_stepwise <- coef_df_stepwise %>%
  dplyr::select(term, p.value, dimensione_impresa) %>%
  pivot_wider(names_from = dimensione_impresa, values_from = p.value, names_prefix = "pval_")

coef_summary_stepwise <- left_join(coef_wide_stepwise, pval_wide_stepwise, by = "term")
print(coef_summary_stepwise)

# Indici di bontà di adattamento
model_fit_stepwise <- map2_df(modelli_per_dim_stepwise, nomi_dimensioni, ~ glance(.x) %>% mutate(dimensione_impresa = .y)) %>%
  dplyr::select(dimensione_impresa, r.squared, adj.r.squared, AIC, BIC, n = df.residual)

print(model_fit_stepwise)

# Visualizzazione coefficienti per dimensione
coef_df_stepwise %>%
  filter(term != "(Intercept)") %>%
  ggplot(aes(x = term, y = estimate, fill = dimensione_impresa)) +
  geom_col(position = "dodge") +
  labs(title = "Confronto coefficienti stimati per dimensione impresa (Stepwise)",
       y = "Valore coefficiente", x = "Variabile") +
  theme_minimal() +
  coord_flip()

#####  CODICE PER ESPORTARE I RISULTATI #####
# Se vuoi salvare in formato Excel
if (!requireNamespace("writexl", quietly = TRUE)) install.packages("writexl")
library(writexl)

write_xlsx(list(
  "Coefficienti_modelli_stepwise" = coef_summary_stepwise,
  "Bonta_adattamento_stepwise" = model_fit_stepwise
), path = "risultati_modelli_dimensioni_stepwise.xlsx")

# Oppure, se preferisci CSV
write.csv(coef_summary_stepwise, "coef_summary_stepwise.csv", row.names = FALSE)
write.csv(model_fit_stepwise, "model_fit_stepwise.csv", row.names = FALSE)
# Scegli il nome del file di output
sink("risultati_modelli_dimensioni_stepwise.txt")

# Scrivi intestazione
cat("===== ANALISI MODELLI PER DIMENSIONE CON SELEZIONE STEPWISE =====\n\n")
# 1. MODELLO SU DATASET COMPLETO
cat("----- MODELLO COMPLETO -----\n\n")
print(summary(model_reduced_full))  # Presumi che questo oggetto esista
cat("\n--- Statistiche di adattamento ---\n")
print(glance(model_reduced_full))
cat("\n\n")

# 2. MODELLO SU CAMPIONE
cat("----- MODELLO CAMPIONE -----\n\n")
print(summary(model_reduced))  # Anche questo deve esistere
cat("\n--- Statistiche di adattamento ---\n")
print(glance(model_reduced))
cat("\n\n")
# Loop su ciascun modello stimato
for (i in seq_along(modelli_per_dim_stepwise)) {
  nome_dim <- nomi_dimensioni[i]
  modello <- modelli_per_dim_stepwise[[i]]
  
  cat("---- MODELLO PER:", nome_dim, "----\n\n")
  print(summary(modello))  # Stampa il sommario del modello (coeff, p-value, R², etc.)
  cat("\n--- Statistiche di adattamento ---\n")
  print(glance(modello))
  cat("\n===========================================\n\n")
}

# Chiudi la connessione al file
sink()
# Rimuovi tutte le variabili tranne dataset, dataset_ridotto e dataset_completo
rm(list = setdiff(ls(), c( "dataset_completo","categorie_variabili","vars_model","normalize")))

# Verifica che solo le variabili desiderate siano rimaste
ls()

saveRDS(dataset_completo, "matrice.rds")

