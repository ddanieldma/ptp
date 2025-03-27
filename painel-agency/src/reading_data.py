import pandas as pd
from pathlib import Path


def exists_parquet(parquet_path: Path | str):
    return Path(parquet_path).exists()

def _save_data_to_parquet(df: pd.DataFrame, parquet_file_path: Path | str):
    if not exists_parquet(parquet_file_path):
        df.to_parquet(parquet_file_path)
    else:
        raise ValueError(f"Parquet file already exists!")
    
def read_data_excel_parquet(excel_file_path: Path | str, sheet_name: str, parquet_file_path: Path | str = None) -> pd.DataFrame:
    if exists_parquet(parquet_file_path):
        df = pd.read_parquet(parquet_file_path)
    else:
        df = pd.read_excel(excel_file_path, sheet_name)
        _save_data_to_parquet(df, parquet_file_path)

    return df



ROOT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT_DIR / "data"
RAW_DIR = DATA_DIR / "raw" 

# Inventário flaviense
BRA_RES_PARQUET_PATH = RAW_DIR / "brazil-resp-rates-all-pres-min-bur.parquet"
BRA_RES_EXCEL_PATH = RAW_DIR / "brazil-resp-rates-all-pres-min-bur.xlsx"

# Planilha de produtos classificados
PAINEL_AGENCY_PARQUET_PATH = RAW_DIR / "painel-agency.parquet"
PAINEL_AGENCY_EXCEL_PATH = RAW_DIR / "painel-agency.xlsx"

# Lendo dados
# Reading inventario flaviense
df_brazil_responses = read_data_excel_parquet(BRA_RES_EXCEL_PATH, "Cabinet & Bureaucracy", BRA_RES_PARQUET_PATH)

# Reading inventario flaviense
df_painel_agency = read_data_excel_parquet(PAINEL_AGENCY_EXCEL_PATH, "Sheet 1", PAINEL_AGENCY_PARQUET_PATH)


if __name__ == "__main__":
    print("df_brazil_responses.head(5)")
    print(df_brazil_responses.head(5))
    print("df_painel_agency.head(5)")
    print(df_painel_agency.head(5))