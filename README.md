# Tesi Magistrale – Analisi Spaziale delle Performance delle Imprese Italiane

Questo repository contiene gli script, i dati di esempio e i risultati principali utilizzati nella tesi magistrale di Pietro Maietta, discussa presso l’Università degli Studi di Napoli “Parthenope”.

## Contenuti
- `scripts/data_analysis_ultimate.R`: costruzione del dataset, pulizia, classificazioni, costruzione e validazione dell’indicatore sintetico di performance (ISP).
- `scripts/analisi_spaziale_isp(2).R`: analisi spaziale delle performance tramite Moran’s I, LISA, Getis-Ord, modelli SAR/SDM e GWR.
- `output/`: tabelle e grafici principali, non inclusi integralmente nella tesi per motivi di spazio.
- `doc/`: documentazione aggiuntiva, inclusa l’Appendice C della tesi.

## Requisiti
- R (>= 4.2.0)
- RStudio
- Librerie R principali: `dplyr`, `ggplot2`, `sf`, `spdep`, `spatialreg`, `glmnet`, `pheatmap`, `leaflet`, `MASS`, `car`, `lmtest`, `caret`.

Puoi installare i pacchetti richiesti con:

```R
packages <- c("dplyr","ggplot2","sf","spdep","spatialreg","glmnet",
              "pheatmap","leaflet","MASS","car","lmtest","caret")
install.packages(packages)
