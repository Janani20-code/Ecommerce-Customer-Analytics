import pandas as pd
from clean_data import clean_data
from load_data import load_sheets


def save_cleaned_file(cleaned_dict, output_path):
    """Saves cleaned data into a new Excel file."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in cleaned_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)
    print(f"Cleaned file saved as: {output_path}")


if __name__ == "__main__":
    input_file = r"C:\Users\janani\OneDrive\Documents\excel files\synthetic_ecommerce_dataset_multisheet.xlsx"
    output_file = r"C:\Users\janani\OneDrive\Documents\excel files\synthetic_ecommerce_dataset_cleaned.xlsx"

    sheets = load_sheets(input_file)
    cleaned = clean_data(sheets)
    save_cleaned_file(cleaned, output_file)
