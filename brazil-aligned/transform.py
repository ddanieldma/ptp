import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import pandas as pd
from collections.abc import Sequence

from extract import *

def get_ws():
    """Returns the worksheet for testing and function use."""
    wb = load_workbook(FILE_PATH)
    return wb[SHEET_NAME]

def get_color_index(cell: openpyxl.cell) -> str:
    """
    Returns the color index of the cell.
    """
    return cell.fill.start_color.index if cell.fill.start_color else None

def categorize_rows(ws: openpyxl.worksheet, column_letter: str, colors_dict: dict) -> list:
    """
    Categorizes rows based on cell colors in the specified column.
    """

    categories = []
    
    for row in ws.iter_rows():
        cell = row[ord(column_letter.upper()) - ord('A')] # Gets the column number using ord
        color = get_color_index(cell) 
        
        if color == colors_dict["amarelo"]:
            category = "contra"
        elif color == colors_dict["vermelho"]:
            category = "alinhada"
        else:
            category = "neutra"
        
        categories.append(category)
    
    return categories


def transform_dataset(
    df: pd.DataFrame,
    categorized_rows_list: Sequence,
    columns: list[str] = [
        'President',
        'Year',
        'Agency',
        '% Concedico e Parcialmente',
    ]) -> pd.DataFrame:
    """
    Transforms the dataset by selecting columns, filling missing values, 
    adding categories, and renaming columns.
    """
    
    # Removing unnecessary columns
    df = df[columns].copy()

    # Propagating last valid observation in the columns (this will correctly
    # fill the President and Year columns)
    df = df.ffill()

    # Adding column with category based in the color
    categorized_rows_list = categorized_rows_list[1:] # The first row is the header of the df
    df["category"] = pd.Series(categorized_rows_list, index=df.index)

    # Setting year column to int
    df["Year"] = df["Year"].astype(int)

    # Renaming columns
    df = df.rename(columns={"% Concedico e Parcialmente": "conc_parc"})
    df.columns = df.columns.str.lower()

    return df

def transform_dataset_from_excel(file_path: Path = Path("data") / "raw" / 'Brazil-Aligned and Non-Aligned All Presidents.xlsx', sheet_name: str = "Cabinet & Bureaucracy", agency_column: str = 'D', columns: list[str] = [
        'President',
        'Year',
        'Agency',
        '% Concedico e Parcialmente',
    ]) -> pd.DataFrame:
    try:
        wb = load_workbook(file_path)
        ws = wb[sheet_name]
    except FileNotFoundError:
        raise FileExistsError(f"Original excel file '{file_path}' expected!")
    
    colors_dict = {
        "branco": get_color_index(ws["D1"]),
        "amarelo": get_color_index(ws["D16"]),
        "vermelho": get_color_index(ws["D257"]),
    }

    categorized_rows_list = categorize_rows(ws, agency_column, colors_dict)

    df = pd.read_excel(file_path, sheet_name=sheet_name)

    return transform_dataset(df, categorized_rows_list, columns)

def prepare_columns_to_excel(df: pd.DataFrame):
    # Capitalizing all columns
    df.columns = df.columns.str.capitalize()

    dict_rename_columns = {
        "Mean_conc_parc": "% Mean Conceded Partially",
    }

    # Renaming only aggregated columns that exist in the dataframe
    rename_columns = {col: new_col for col, new_col in dict_rename_columns.items() if col in df.columns}
    df = df.rename(columns=rename_columns)


    return df


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