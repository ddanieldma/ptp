from pathlib import Path
from openpyxl import load_workbook  
import openpyxl.worksheet.worksheet as Worksheet
import pandas as pd
from collections.abc import Sequence

DATA_DIR = Path("data")
RAW_DIR = DATA_DIR / "raw"
INTERIM_DIR = DATA_DIR / "interim"

# Constants and setup
FILE_PATH = RAW_DIR / 'Brazil-Aligned and Non-Aligned All Presidents.xlsx'
SHEET_NAME = 'Cabinet & Bureaucracy'
AGENCY_COLUMN = 'D'

PARQUET_PATH = RAW_DIR / 'data-editada-parquet-raw.parquet'

def get_ws():
    """Returns the worksheet for testing and function use."""
    wb = load_workbook(FILE_PATH)
    return wb[SHEET_NAME]

def get_df_from_excel(file_path: str | Path = None, sheet_name: str = None) -> pd.DataFrame:
    """
    Gets the dataframe from the excel file.

    Parameters
    ----------
    file_path: str | Path, optional
        The path to the excel file . If None, uses the default.
    
    sheet_name: str
        The name of the sheet with the dataframe. If none, uses default.
    """

    if not file_path:
        file_path = FILE_PATH
    
    if not sheet_name:
        sheet_name = SHEET_NAME

    try:
        return pd.read_excel(file_path, sheet_name)
    except FileNotFoundError:
        raise(f"File {file_path} doesn't exist!")
    
get_df_from_excel(None, "oi")