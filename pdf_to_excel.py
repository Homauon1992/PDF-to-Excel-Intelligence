import os
import sys

import numpy as np
import pandas as pd
import pdfplumber
from colorama import Fore, init


def extract_tables(pdf_path: str) -> list[tuple[str, pd.DataFrame]]:
    if not pdf_path or not os.path.isfile(pdf_path):
        raise FileNotFoundError(pdf_path)

    extracted: list[tuple[str, pd.DataFrame]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            for table_index, table in enumerate(tables, start=1):
                if not table:
                    continue
                df = pd.DataFrame(table)
                df = df.replace(r"^\s*$", np.nan, regex=True)
                df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
                if df.empty:
                    continue
                sheet_name = f"page{page_index}_table{table_index}"
                extracted.append((sheet_name, df))

    if not extracted:
        raise ValueError("No Tables Found")

    return extracted


def save_to_excel(tables: list[tuple[str, pd.DataFrame]], output_path: str) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in tables:
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])


def main() -> None:
    init(autoreset=True)
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else input("Enter PDF path: ").strip()

    try:
        tables = extract_tables(pdf_path)
        save_to_excel(tables, "converted_data.xlsx")
        print(
            Fore.GREEN
            + "Success! All tables have been extracted and saved to Excel."
        )
    except FileNotFoundError:
        print(Fore.RED + f"File Not Found: {pdf_path}")
    except ValueError as exc:
        if str(exc) == "No Tables Found":
            print(Fore.RED + "No Tables Found in the provided PDF.")
        else:
            raise


if __name__ == "__main__":
    main()