import pandas as pd
from load_data import load_sheets


def clean_data(sheet_dict):
    """Cleans all sheets with column-specific cleaning rules (no processing)."""
    cleaned_data = {}

    for name, df in sheet_dict.items():
        print(f"Cleaning sheet: {name}")

        # Remove duplicates
        df = df.drop_duplicates()

        # Sheet-specific cleaning
        if name == 'customers':
            df = clean_customers_sheet(df)
        elif name == 'products':
            df = clean_products_sheet(df)
        elif name == 'orders':
            df = clean_orders_sheet(df)
        elif name == 'reviews':
            df = clean_reviews_sheet(df)

        # Reset index
        df = df.reset_index(drop=True)

        cleaned_data[name] = df
        print(f"Finished cleaning: {name}")

    return cleaned_data


def clean_customers_sheet(df):
    """Cleans the customers sheet with specific rules."""
    # Handle missing values
    df['customer_id'] = df['customer_id'].fillna(0).astype(int)
    df['gender'] = df['gender'].fillna('Unknown')
    df['age'] = df['age'].fillna(df['age'].median()).astype(int)
    df['city'] = df['city'].fillna('Unknown')
    df['loyalty_score'] = df['loyalty_score'].fillna(df['loyalty_score'].mean()).round(1)

    # Data validation
    df = df[df['age'] > 0]  # Remove negative ages
    df = df[df['age'] <= 120]  # Remove unrealistic ages

    # Clean gender values
    df['gender'] = df['gender'].str.strip().str.title()
    valid_genders = ['Male', 'Female', 'Other', 'Unknown']
    df['gender'] = df['gender'].apply(lambda x: x if x in valid_genders else 'Unknown')

    return df


def clean_products_sheet(df):
    """Cleans the products sheet with specific rules."""
    # Handle missing values
    df['product_id'] = df['product_id'].fillna(0).astype(int)
    df['product_name'] = df['product_name'].fillna('Unknown Product')
    df['category'] = df['category'].fillna('Uncategorized')
    df['price'] = df['price'].fillna(df['price'].mean()).round(2)
    df['stock'] = df['stock'].fillna(0).astype(int)

    # Data validation
    df = df[df['price'] >= 0]  # Remove negative prices
    df = df[df['stock'] >= 0]  # Remove negative stock

    return df


def clean_orders_sheet(df):
    """Cleans the orders sheet with specific rules."""
    # Handle missing values
    df['order_id'] = df['order_id'].fillna(0).astype(int)
    df['customer_id'] = df['customer_id'].fillna(0).astype(int)
    df['order_date'] = pd.to_datetime(df['order_date'], errors='coerce')
    df['product_category'] = df['product_category'].fillna('Unknown')
    df['order_value'] = df['order_value'].fillna(df['order_value'].mean()).round(2)
    df['payment_method'] = df['payment_method'].fillna('Unknown')
    df['delivered'] = df['delivered'].fillna('No')

    # Clean payment method values
    df['payment_method'] = df['payment_method'].str.strip().str.title()

    # Clean delivered column
    df['delivered'] = df['delivered'].astype(str).str.strip().str.title()
    df['delivered'] = df['delivered'].replace({'True': 'Yes', 'False': 'No', 'Yes': 'Yes', 'No': 'No'})
    df['delivered'] = df['delivered'].fillna('No')

    # Remove orders with invalid dates
    df = df[df['order_date'].notna()]

    return df


def clean_reviews_sheet(df):
    """Cleans the reviews sheet with specific rules."""
    # Handle missing values
    df['review_id'] = df['review_id'].fillna(0).astype(int)
    df['customer_id'] = df['customer_id'].fillna(0).astype(int)
    df['product_id'] = df['product_id'].fillna(0).astype(int)
    df['rating'] = df['rating'].fillna(df['rating'].mean()).round(1)
    df['review_text'] = df['review_text'].fillna('No review text')

    # Data validation for ratings
    df = df[df['rating'] >= 1]
    df = df[df['rating'] <= 5]

    # Clean review text
    df['review_text'] = df['review_text'].str.strip()

    return df


if __name__ == "__main__":
    file_path = r"C:\Users\janani\OneDrive\Documents\excel files\synthetic_ecommerce_dataset_multisheet.xlsx"
    sheets = load_sheets(file_path)
    cleaned_sheets = clean_data(sheets)

    print("\nCleaning completed! Ready for analysis.")
    for sheet_name, df in cleaned_sheets.items():
        print(f"{sheet_name}: {len(df)} rows, {len(df.columns)} columns")