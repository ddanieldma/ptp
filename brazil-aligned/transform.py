import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import pandas as pd

from extract import *


def prepare_columns_to_excel(df: pd.DataFrame):
    # Capitalizing all columns
    df.columns = df.columns.str.capitalize()

    dict_rename_columns = {
        "Mean_conc_parc": "% Mean Conceded and Partially",
        "conc_parc": "% Conceded and Partially",
        "% conceded": "% Conceded",
        "% partially conceded": "% Partially Conceded",
        "% denied": "% Denied",
        "% others": "% Others",
    }

    # Renaming only aggregated columns that exist in the dataframe
    rename_columns = {col: new_col for col, new_col in dict_rename_columns.items() if col in df.columns}
    df = df.rename(columns=rename_columns)

    return df

def save_dataframe_to_excel_tab(df: pd.DataFrame, file_path: Path | str, sheet_name: str):
    # Save the DataFrame to the specific sheet
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    DATA_DIR = Path("data")
    RAW_DIR = DATA_DIR / "raw"

    # Constants and setup
    FILE_NAME = 'Brazil-Aligned and Non-Aligned All Presidents.xlsx'
    FILE_PATH = RAW_DIR / FILE_NAME
    SHEET_NAME = 'Cabinet & Bureaucracy'

    # Loading workbook and sheet
    ws = None
    try:
        wb = load_workbook(FILE_PATH)
        ws = wb[SHEET_NAME]
    except FileNotFoundError:
        raise FileExistsError(f"Original excel file '{FILE_NAME}' expected!")

    AGENCY_COLUMN = 'D'
    
    colors_dict = {
        "branco": get_color_index(ws["D1"]),
        "amarelo": get_color_index(ws["D16"]),
        "vermelho": get_color_index(ws["D257"]),
    }

    categorized_rows_list = categorize_rows(ws, AGENCY_COLUMN, colors_dict)