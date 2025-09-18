import pandas as pd
from read_file import read_excel_file

def load_sheets(file_path):
    """Loads all sheets into a dictionary of DataFrames."""
    excel_data = read_excel_file(file_path)
    if excel_data:
        sheet_dict = {sheet: excel_data.parse(sheet) for sheet in excel_data.sheet_names}
        print("Sheets loaded successfully!")
        return sheet_dict
    return {}

if __name__ == "__main__":
    file_path = r"C:\Users\janani\OneDrive\Documents\excel files\synthetic_ecommerce_dataset_multisheet.xlsx"
    sheets = load_sheets(file_path)
    print("Available sheets:", list(sheets.keys()))
