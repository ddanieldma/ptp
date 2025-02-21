from pathlib import Path
from openpyxl import load_workbook

DATA_DIR = Path("data")
RAW_DIR = DATA_DIR / "raw"
INTERIM_DIR = DATA_DIR / "interim"

# Constants and setup
FILE_PATH = RAW_DIR / 'Brazil-Aligned and Non-Aligned All Presidents.xlsx'
SHEET_NAME = 'Cabinet & Bureaucracy'
AGENCY_COLUMN = 'D'

# Loading excel file
wb = load_workbook(FILE_PATH)
ws = wb[SHEET_NAME]