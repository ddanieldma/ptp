import pandas as pd
from pathlib import Path

def print_equals(num: int) -> None:
    print(num*"=")

def write_to_excel(df: pd.DataFrame, file_path: Path | str, sheet_name: str) -> None:
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)