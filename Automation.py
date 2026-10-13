
from pathlib import Path
import pandas as pd

# Project folder
BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "sample_data.xlsx"
OUTPUT_FILE = BASE_DIR / "cleaned_data.xlsx"
OUTPUT_CSV = BASE_DIR / "cleaned_data.csv"


def create_sample_data():
    """Create a sample Excel file for practice."""

    data = {
        " Customer Name ": [
            " Rahul Sharma ",
            "Priya Singh",
            "Rahul Sharma",
            " Amit Verma ",
            "Neha Gupta",
            "Priya Singh",
            "  "
        ],
        "Order Date": [
            "12-10-2026",
            "13-10-2026",
            "12-10-2026",
            "14-10-2026",
            None,
            "13-10-2026",
            None
        ],
        "Product": [
            " Laptop ",
            "Mouse",
            " Laptop ",
            " Keyboard ",
            "Monitor",
            "Mouse",
            None
        ],
        "Amount": [
            55000,
            800,
            55000,
            1500,
            None,
            800,
            None
        ]
    }

    df = pd.DataFrame(data)
    df.to_excel(INPUT_FILE, index=False)

    print(f"Sample Excel file created: {INPUT_FILE.name}")


def clean_data():
    """Clean the sample Excel data and export results."""

    # Read Excel file
    df = pd.read_excel(INPUT_FILE)

    original_rows = len(df)
    original_duplicates = int(df.duplicated().sum())

    # Clean column names
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", "_", regex=True)
    )

    # Clean text columns
    for column in df.select_dtypes(
        include=["object", "string"]
    ).columns:
        df[column] = (
            df[column]
            .astype("string")
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )

    # Convert blank text to missing values
    df = df.replace(r"^\s*$", pd.NA, regex=True)

    # Remove empty rows and duplicate rows
    df = df.dropna(how="all")
    df = df.drop_duplicates()

    # Convert order dates
    if "order_date" in df.columns:
        df["order_date"] = pd.to_datetime(
            df["order_date"],
            format="%d-%m-%Y",
            errors="coerce"
        )

    # Export cleaned files
    df.to_excel(OUTPUT_FILE, index=False)
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    # Print report
    print("\n===== DATA CLEANING REPORT =====")
    print(f"Original rows: {original_rows}")
    print(f"Duplicate rows found: {original_duplicates}")
    print(f"Cleaned rows: {len(df)}")
    print(f"Rows removed: {original_rows - len(df)}")
    print(f"Missing cells remaining: {int(df.isna().sum().sum())}")
    print(f"\nExcel output: {OUTPUT_FILE.name}")
    print(f"CSV output: {OUTPUT_CSV.name}")
    print("\nCleaning completed successfully!")


if __name__ == "__main__":
    create_sample_data()
    clean_data()
