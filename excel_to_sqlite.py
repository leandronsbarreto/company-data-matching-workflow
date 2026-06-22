from pathlib import Path
import sqlite3
import re

import pandas as pd


# ============================================================
# SETTINGS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DATA_DIR = BASE_DIR / "sample_data"

MATCH_FILE = SAMPLE_DATA_DIR / "Match_Data.xlsx"
SEARCH_EXPORT_FILE = SAMPLE_DATA_DIR / "Search_Export_001.xlsx"
EXTRACTED_BATCH_FILE = SAMPLE_DATA_DIR / "Extracted_Batch.xlsx"

DATABASE_FILE = BASE_DIR / "company_matching_demo.db"


# ============================================================
# HELPERS
# ============================================================

def normalize_column_name(column_name):
    """
    Convert Excel headers into SQL-friendly column names.

    Example:
    "Interne ID" -> "interne_id"
    "Datenbank-ID" -> "datenbank_id"
    "NACE Rev. 2 - Haupttätigkeit - Code" -> "nace_rev_2_haupttaetigkeit_code"
    """
    text = str(column_name).strip().lower()

    replacements = {
        "ä": "ae",
        "ö": "oe",
        "ü": "ue",
        "ß": "ss",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text)
    text = text.strip("_")

    return text


def clean_dataframe_columns(df):
    df = df.copy()
    df.columns = [normalize_column_name(col) for col in df.columns]
    return df


def load_excel_sheet(file_path, sheet_name):
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
    df = clean_dataframe_columns(df)

    return df


# ============================================================
# MAIN
# ============================================================

def main():
    if not SAMPLE_DATA_DIR.exists():
        raise FileNotFoundError(
            f"Folder not found: {SAMPLE_DATA_DIR}. "
            "Please make sure the sample_data folder exists."
        )

    print("Loading Excel files...")

    source_data = load_excel_sheet(MATCH_FILE, "Source_Data")
    not_found = load_excel_sheet(MATCH_FILE, "Not_Found")
    search_export = load_excel_sheet(SEARCH_EXPORT_FILE, "Search_Export")
    extracted_batch = load_excel_sheet(EXTRACTED_BATCH_FILE, "Extracted_Batch")

    print(f"Creating SQLite database: {DATABASE_FILE.name}")

    with sqlite3.connect(DATABASE_FILE) as connection:
        source_data.to_sql("source_data", connection, if_exists="replace", index=False)
        not_found.to_sql("not_found", connection, if_exists="replace", index=False)
        search_export.to_sql("search_export", connection, if_exists="replace", index=False)
        extracted_batch.to_sql("extracted_batch", connection, if_exists="replace", index=False)

    print("")
    print("Done.")
    print(f"Database created: {DATABASE_FILE}")
    print("")
    print("Tables created:")
    print("- source_data")
    print("- not_found")
    print("- search_export")
    print("- extracted_batch")


if __name__ == "__main__":
    main()