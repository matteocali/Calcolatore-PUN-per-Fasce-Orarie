# Calcolatore PUN (Prezzo Unico Nazionale) per Fasce Orarie

Questo script Python analizza i dati orari del PUN (Prezzo Unico Nazionale dell'energia elettrica) e calcola i prezzi medi aggregati per le fasce di consumo **F1, F2 e F3**, generando report dettagliati in Excel e grafici in PDF.

## 🎯 Funzionalità

- **Analisi Automatica**: Legge i file Excel contenenti i dati orari del PUN.
- **Calcolo Fasce**: Applica la logica delle fasce orarie (F1, F2, F3) tenendo conto di:
  - giorni feriali vs festivi/weekend;
  - orari specifici;
  - festività nazionali italiane (pacchetto `holidays`).
- **Aggregazione Flessibile**:
  - **Giornaliera**: Media dei prezzi per ogni giorno divisa per fasce.
  - **Mensile**: Media mensile aggregata per fasce.
- **Output Completi**:
  - 📊 Grafico PDF vettoriale dell'andamento prezzi.
  - 📈 Tabella Excel con i dati elaborati.
  - Calcolo del prezzo medio generale (senza distinzione di fascia).

## 📥 Dove scaricare i dati

I dati ufficiali del PUN possono essere scaricati dal sito del **GME (Gestore Mercati Energetici)**:

🔗 [https://gme.mercatoelettrico.org/it-it/Home/Esiti/Elettricita/MGP/Esiti/PUN](https://gme.mercatoelettrico.org/it-it/Home/Esiti/Elettricita/MGP/Esiti/PUN)

1. Selezionare l'intervallo di date di interesse.
2. Scegliere la granularità **Oraria** (è fondamentale per il calcolo corretto delle fasce).
3. Scaricare il file Excel (`.xlsx` o `.xls`).
4. Posizionare il file nella cartella del progetto (opzionale).

## ⚙️ Installazione e Utilizzo

Il progetto è configurato per utilizzare **[uv](https://github.com/astral-sh/uv)**, un tool estremamente veloce per la gestione dei progetti Python.

### 1. Prerequisiti

Assicurati di avere `uv` installato sul tuo sistema.

### 2. Setup Ambiente

Esegui il sync per creare l'ambiente virtuale e installare tutte le dipendenze necessarie (`pandas`, `openpyxl`, `matplotlib`, `holidays`):

```bash
uv sync
```

### 3. Esecuzione

Per eseguire lo script principale (`pun_calculator.py`):

```bash
uv run pun_calculator.py
```

Lo script cercherà automaticamente il primo file Excel disponibile nella cartella corrente o in `data/` e ti guiderà con un menu interattivo.

### 4. Opzioni Avanzate (Riga di Comando)

Puoi velocizzare l'esecuzione specificando il file di input o la modalità direttamente da terminale:

- **Specificare un file specifico:**

  ```bash
  uv run pun_calculator.py data/2026_dati_pun.xlsx
  ```

- **Forzare la modalità di calcolo (giornaliero/mensile):**

  ```bash
  uv run pun_calculator.py --mode mensile
  ```

- **Combinazione:**

  ```bash
  uv run pun_calculator.py data/gennaio_2026.xlsx -m giornaliero
  ```

## 🕒 Definizione delle Fasce Orarie

Il calcolo segue lo standard ARERA per le tariffe biorarie/multiorarie:

| Fascia | Descrizione | Orari | Giorni |
| :--- | :--- | :--- | :--- |
| **F1 (Punta)** | Ore di maggior carico | **8:00 - 19:00** | Lunedì - Venerdì (esclusi festivi) |
| **F2 (Intermedia)** | Ore intermedie | **7:00 - 8:00** <br> **19:00 - 23:00** <br> **7:00 - 23:00** (Sab) | Lunedì - Venerdì <br> Sabato (esclusi festivi) |
| **F3 (Fuori Punta)** | Ore di minor carico | **23:00 - 7:00** <br> **Tutto il giorno** | Lunedì - Sabato <br> Domenica e Festivi Nazionali |

*Nota: Le festività nazionali vengono rilevate automaticamente anno per anno grazie alla libreria `holidays`.*
