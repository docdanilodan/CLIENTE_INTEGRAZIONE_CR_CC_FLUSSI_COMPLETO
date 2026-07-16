# FinancePlus - Clienti e Pratiche

## File principale
`app_financeplus_clienti_pratiche.py`

## Funzioni
- Lettura visure e report PDF
- Estrazione automatica denominazione, P.IVA, Comune, indirizzo, CAP e Provincia
- Più moduli pratica per ogni cliente
- Prodotti: Chiro, Factoring, Leasing, Invoice Trading, Mutuo
- Gestori: Paolo, Sergio, Zeno, Nick
- Stato semaforo verde/giallo
- Report PDF per singolo cliente
- Report PDF generale
- Archivio SQLite

## Streamlit Cloud
Caricare nella cartella principale:
- `app_financeplus_clienti_pratiche.py`
- `requirements.txt`

Main file da selezionare:
`app_financeplus_clienti_pratiche.py`

## Collaboratori
- Nuova voce di menu `Inserisci collaboratore`
- Campi Nome e Cognome
- Salvataggio anagrafica collaboratore
- Menu a tendina Collaboratore nella scheda cliente
- Collaboratore riportato in archivio e report cliente

## Nuova area CLIENTI - Analisi / Value
- Voce `CLIENTI` sotto `Inserisci cliente`
- Selezione cliente da menu a tendina
- Pulsanti: `CR`, `CC`, `Bilancio`, `Bozza Bilancio`
- Modulo CR operativo con caricamento `.xlsx`, `.csv`, `.txt`, `.pdf` testuale
- Normalizzazione automatica colonne: periodo, banca, categoria, accordato, utilizzato, sconfino, garanzia, sofferenza, perdita
- Report CR avanzato strutturato su 45 pagine, coerente con il fac-simile fornito
- Anteprima PDF integrata
- Download PDF con denominazione cliente e data in italiano
- Download HTML
- Pacchetto ZIP con PDF, HTML, grafici PNG, tabelle CSV e dati JSON strutturati

## Nota tecnica
L'estrazione dai PDF testuali dipende dalla struttura del prospetto. I dati riclassificati devono essere verificati rispetto al documento ufficiale della Centrale dei Rischi.
