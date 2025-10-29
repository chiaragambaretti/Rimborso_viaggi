# Rimborso Viaggi
Progetto realizzato durante il corso **Data management con python**
Questo progetto permette di automatizzare la gestione dei rimborsi viaggio di un collaboratore. Lo script estrae i dati dai biglietti PDF, genera un file Excel di riepilogo e crea una ricevuta in formato Word pronta per firma.

## Funzionalità principali

1. Legge tutti i file **PDF** presenti nella cartella `Viaggi da rimborsare`.
2. Estrae automaticamente: Stazione di partenza, Stazione di arrivo, Data del viaggio, Prezzo del biglietto
3. Inserisce i dati in un file Excel `Riepilogo rimborsi.xlsx` con:
   - Foglio nuovo per ogni esecuzione
   - Colonne: Data, Partenza, Arrivo, Prezzo, Nome file
   - Calcolo totale rimborsi
5. Genera una ricevuta **Word** con la somma totale e l’elenco dei viaggi, pronta per la firma del collaboratore.

## Requisit

- Python ≥ 3.8
- Librerie Python:
  - `PyMuPDF` (`fitz`)
  - `openpyxl`
  - `python-docx`
