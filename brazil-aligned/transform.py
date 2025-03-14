import openpyxl
from openpyxl import load_workbook
from pathlib import Path
import pandas as pd

from extract import *


def prepare_columns_to_excel(df: pd.DataFrame):
    # Dropping category column if it exists
    if 'category' in list(df.columns):
        df = df.drop(columns=['category'])
    
    # Changing names back to the original
    df.columns = df.columns.str.replace("_", " ")
    df.columns = df.columns.str.strip().str.replace("perc", "%", case=False)
    df.columns = df.columns.str.title()

    return df

def save_dataframe_to_excel_tab(df: pd.DataFrame, file_path: Path | str, sheet_name: str):
    # Save the DataFrame to the specific sheet
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    from extract import get_interim_data

    df = get_interim_data()
    print(df.head(5))

    excel_df = prepare_columns_to_excel(df)
    print(excel_df.head(5))