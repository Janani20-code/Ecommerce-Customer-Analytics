import pandas as pd

def read_excel_file(file_path):
    """Reads the Excel file with multiple sheets."""
    try:
        data = pd.ExcelFile(file_path)  # keeps structure
        print("File read successfully!")
        return data
    except Exception as e:
        print("Error reading file:", e)
        return None

if __name__ == "__main__":
    file_path = r"C:\Users\janani\OneDrive\Documents\excel files\synthetic_ecommerce_dataset_multisheet.xlsx"
    excel_data = read_excel_file(file_path)

