import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import holidays
import glob
import os
import sys
import argparse


def load_data(file_path=None):
    """
    Carica i dati dal file Excel specificato o cerca il primo .xlsx/.xls nella directory corrente.
    """
    if file_path:
        if not os.path.exists(file_path):
            print(f"Errore: Il file '{file_path}' non esiste.")
            sys.exit(1)
        print(f"\nCaricamento file specificato: {file_path}")
    else:
        # Cerca il file Excel nella directory corrente
        patterns = ["*.xlsx", "*.xls"]
        files = []
        for p in patterns:
            files.extend(glob.glob(p))

        if not files:
            # Fallback: cerca in data/ se esiste
            patterns_data = [
                os.path.join("data", "*.xlsx"),
                os.path.join("data", "*.xls"),
            ]
            for p in patterns_data:
                files.extend(glob.glob(p))

            if not files:
                print(
                    "Nessun file Excel trovato nella directory corrente o in 'data/'."
                )
                sys.exit(1)

        file_path = files[0]
        print(f"\nCaricamento file automatico: {file_path}")

    try:
        # Leggi Excel con Pandas
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Errore nella lettura del file: {e}")
        sys.exit(1)

    # Standardizza i nomi delle colonne rimuovendo spazi
    df.columns = [str(c).strip() for c in df.columns]

    # Identifica le colonne necessarie
    price_col = next((c for c in df.columns if "MWh" in c or "Prezzo" in c), None)
    date_col = next((c for c in df.columns if "Data" in c or "Date" in c), None)
    hour_col = next((c for c in df.columns if "Ora" in c or "Hour" in c), None)

    if not all([price_col, date_col, hour_col]):
        print(f"\nColonne trovate: {df.columns.tolist()}")
        print(
            "ERRORE: Impossibile identificare le colonne necessarie (Data, Ora, Prezzo/MWh)."
        )
        sys.exit(1)

    # Pulisci e converti la colonna Prezzo
    df[price_col] = (
        df[price_col].astype(str).str.replace(",", ".", regex=False).str.strip()
    )
    df[price_col] = pd.to_numeric(df[price_col], errors="coerce")

    # Converti MWh a kWh (dividi per 1000)
    df["Prezzo_kWh"] = df[price_col] / 1000.0

    # Parsing Data (DD/MM/YYYY)
    df["Date"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")

    if df["Date"].isnull().any():
        print(
            "Attenzione: alcune date non sono state riconosciute e verranno ignorate."
        )
        df = df.dropna(subset=["Date"])

    # Creazione indice Datetime completo
    # L'ora nel file PUN va da 1 a 24.
    df["Hour_Index"] = df[hour_col] - 1
    df["Datetime"] = df.apply(
        lambda row: row["Date"] + pd.Timedelta(hours=row["Hour_Index"]), axis=1
    )

    return df


def get_fascia(datetime_val, it_holidays):
    """
    Restituisce la fascia (F1, F2, F3) per una data ora.
    F1: Lun-Ven 8-19
    F2: Lun-Ven 7-8/19-23, Sab 7-23
    F3: Lun-Sab 23-7, Dom/Festivi tutto il giorno
    """
    day_of_week = datetime_val.weekday()  # 0=Lun, 6=Dom
    hour = datetime_val.hour  # 0-23
    date_val = datetime_val.date()

    is_holiday = date_val in it_holidays
    is_sunday = day_of_week == 6

    # F3: Domenica e Festivi tutto il giorno, oppure Notte (23-7)
    if is_holiday or is_sunday or (hour >= 23 or hour < 7):
        return "F3"

    # Sabato (non festivo, e non notte) -> F2
    if day_of_week == 5:
        return "F2"

    # Lunedì - Venerdì (non festivi)
    if 8 <= hour < 19:
        return "F1"
    else:
        return "F2"


def main():
    parser = argparse.ArgumentParser(description="Calcolatore Medio PUN per Fasce")
    parser.add_argument(
        "input_file", nargs="?", help="Percorso del file Excel di input (opzionale)"
    )
    parser.add_argument(
        "-m",
        "--mode",
        choices=["giornaliero", "mensile"],
        help="Modalità di aggregazione: giornaliero o mensile",
    )
    args = parser.parse_args()

    print("--- Calcolatore Medio PUN per Fasce ---")

    # 1. Caricamento Dati e Calcolo Fasce
    df = load_data(args.input_file)
    years = df["Date"].dt.year.unique()
    it_holidays = holidays.IT(years=years)

    print("Elaborazione fasce orarie in corso...")
    df["Fascia"] = df["Datetime"].apply(lambda x: get_fascia(x, it_holidays))

    # Determine date range for filename
    min_date = df["Date"].min().strftime("%Y%m%d")
    max_date = df["Date"].max().strftime("%Y%m%d")
    date_range_str = f"{min_date}-{max_date}"

    # 2. Scelta Intervallo
    scelta = None
    if args.mode:
        if args.mode == "giornaliero":
            scelta = "1"
        elif args.mode == "mensile":
            scelta = "2"
    else:
        print("\nScegli l'intervallo di aggregazione:")
        print("1. Giornaliero")
        print("2. Mensile")
        try:
            input_val = input("Inserisci il numero (1 o 2): ").strip()
            if input_val in ["1", "2"]:
                scelta = input_val
            else:
                print("Scelta non valida. Default a Mensile.")
                scelta = "2"
        except EOFError:
            scelta = "2"
            print("Default a Mensile")

    aggregation_name = ""
    result_df = None

    if scelta == "1":
        aggregation_name = "Giornaliero"
        # Raggruppa per Giorno e Fascia -> Media Prezzo
        # Pre-calcolo media globale giornaliera (su tutte le ore)
        daily_avg = df.groupby(df["Datetime"].dt.date)["Prezzo_kWh"].mean()

        grouped = df.groupby([df["Datetime"].dt.date, "Fascia"])["Prezzo_kWh"].mean()
        result_df = grouped.unstack()  # Pivot per avere F1, F2, F3 come colonne
        result_df.index.name = "Data"

        # Aggiungo colonna media complessiva
        result_df["Media"] = daily_avg

    elif scelta == "2":
        aggregation_name = "Mensile"
        # Raggruppa per Mese e Fascia -> Media Prezzo
        df["Mese"] = df["Datetime"].dt.to_period("M")

        # Pre-calcolo media globale mensile
        monthly_avg = df.groupby("Mese")["Prezzo_kWh"].mean()

        grouped = df.groupby(["Mese", "Fascia"])["Prezzo_kWh"].mean()
        result_df = grouped.unstack()

        # Aggiungo colonna media complessiva
        result_df["Media"] = monthly_avg

    # 3. Esportazione Tabella Excel
    base_filename = f"pun_{aggregation_name.lower()}_{date_range_str}"
    excel_filename = f"tabella_{base_filename}.xlsx"

    try:
        export_df = result_df.copy()

        if scelta == "1":  # Giornaliero
            # Formatta indice come stringa GG/MM/AAAA
            export_df.index = pd.to_datetime(export_df.index).strftime("%d/%m/%Y")
        else:  # Mensile
            export_df.index = export_df.index.astype(str)

        export_df.to_excel(excel_filename)
        print(f"\nTabella esportata in '{excel_filename}'")
    except Exception as e:
        print(f"Errore nell'esportazione Excel: {e}")

    # 4. Generazione Grafico
    print("\nGenerazione grafico...")
    plt.figure(figsize=(12, 6))

    colors = {"F1": "red", "F2": "orange", "F3": "green"}

    # Disegno grafico a linee con pallini
    # Gestione diversa solo per asse X (Date vs Stringhe) ma logica plot simile

    if scelta == "1":  # Giornaliero (asse X temporale)
        # Convertiamo indice in datetime per correttezza su asse X
        result_df.index = pd.to_datetime(result_df.index)

        # Plot delle Fasce
        for col in ["F1", "F2", "F3"]:
            if col in result_df.columns:
                series = result_df[col].dropna()
                if not series.empty:
                    # Plot serie pulita da NaN per avere linee congiunte
                    plt.plot(
                        series.index,
                        series.values,
                        marker="o",
                        linestyle="-",
                        label=col,
                        color=colors.get(col, "blue"),
                    )

        # Plot della Media Generale
        if "Media" in result_df.columns:
            series = result_df["Media"].dropna()
            if not series.empty:
                plt.plot(
                    series.index,
                    series.values,
                    marker="s",
                    linestyle="--",
                    label="Media",
                    color="gray",
                    alpha=0.7,
                    linewidth=1.5,
                )

        # Formattazione asse X
        ax = plt.gca()
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m/%Y"))
        # Imposta locator automatico per gestire la densità delle date
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())

        plt.gcf().autofmt_xdate()  # Ruota date

    else:  # Mensile (asse X categorico/stringa)
        x_values = result_df.index.astype(str)

        for col in ["F1", "F2", "F3"]:
            if col in result_df.columns:
                # Per plot categorico, dobbiamo assicurarci di plottare solo i valori validi
                valid_mask = result_df[col].notna()
                if valid_mask.any():
                    plt.plot(
                        x_values[valid_mask],
                        result_df.loc[valid_mask, col],
                        marker="o",
                        linestyle="-",
                        label=col,
                        color=colors.get(col, "blue"),
                    )

        # Plot della Media Generale (Mensile)
        if "Media" in result_df.columns:
            valid_mask = result_df["Media"].notna()
            if valid_mask.any():
                plt.plot(
                    x_values[valid_mask],
                    result_df.loc[valid_mask, "Media"],
                    marker="s",
                    linestyle="--",
                    label="Media",
                    color="gray",
                    alpha=0.7,
                    linewidth=1.5,
                )

        plt.xticks(rotation=45)

    # Stile comune
    plt.title(f"Prezzo Medio Energia per Fascia - {aggregation_name}")
    plt.ylabel("Prezzo Medio (€/kWh)")
    plt.xlabel("Data/Periodo")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend(title="Fascia")

    plt.tight_layout()

    # Salvataggio PDF
    output_filename = f"grafico_{base_filename}.pdf"
    plt.savefig(output_filename, format="pdf")
    print(f"Grafico salvato come '{output_filename}'")

    try:
        plt.show()
    except Exception as e:
        print(f"Errore durante la visualizzazione del grafico: {e}")


if __name__ == "__main__":
    main()
