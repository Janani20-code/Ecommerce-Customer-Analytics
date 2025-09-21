import pandas as pd
import os

# -------------------------------
# Ensure 'data' folder exists
# -------------------------------
os.makedirs("data", exist_ok=True)

# Path to cleaned Excel
file_path = r"C:\Users\janani\OneDrive\Documents\excel files\synthetic_ecommerce_dataset_cleaned.xlsx"

# Load all sheets
customers = pd.read_excel(file_path, sheet_name="customers")
products = pd.read_excel(file_path, sheet_name="products")
orders = pd.read_excel(file_path, sheet_name="orders")
reviews = pd.read_excel(file_path, sheet_name="reviews")

# -------------------------------
# 1. Customer Segmentation (RFM)
# -------------------------------
def customer_segmentation():
    rfm = orders.groupby("customer_id").agg({
        "order_date": lambda x: (orders["order_date"].max() - x.max()).days, # Recency
        "order_id": "count",                                                # Frequency
        "order_value": "sum"                                                # Monetary
    }).reset_index()

    rfm.columns = ["customer_id", "Recency", "Frequency", "Monetary"]

    # Simple segmentation rules
    rfm["Segment"] = "Regular"
    rfm.loc[(rfm["Recency"] < 30) & (rfm["Frequency"] > 3) & (rfm["Monetary"] > 500), "Segment"] = "VIP"
    rfm.loc[(rfm["Recency"] < 60) & (rfm["Frequency"] >= 2), "Segment"] = "Loyal"
    rfm.loc[rfm["Recency"] > 90, "Segment"] = "Lost"

    rfm.to_csv(os.path.join("data", "01_customer_segmentation.csv"), index=False)
    print("Customer Segmentation saved!")

# ---------------------------------------
# 2. Top Products & Categories Analysis
# ---------------------------------------
def top_products_categories():
    merged = orders.merge(products, left_on="product_category", right_on="category", how="left")
    top_products = merged.groupby("product_name")["order_value"].sum().reset_index().sort_values(by="order_value", ascending=False)
    top_categories = merged.groupby("category")["order_value"].sum().reset_index().sort_values(by="order_value", ascending=False)

    top_products.to_csv(os.path.join("data", "02_top_products.csv"), index=False)
    top_categories.to_csv(os.path.join("data", "02_top_categories.csv"), index=False)
    print("Top Products & Categories saved!")

# ---------------------------------------
# 3. Customer Demographics & Spending
# ---------------------------------------
def customer_demographics():
    customer_spend = orders.groupby("customer_id")["order_value"].sum().reset_index()
    merged = customers.merge(customer_spend, on="customer_id", how="left").fillna(0)

    # Age group buckets
    bins = [0, 25, 40, 60, 100]
    labels = ["Young", "Adult", "Middle-aged", "Senior"]
    merged["Age_Group"] = pd.cut(merged["age"], bins=bins, labels=labels, right=False)

    demographics = merged.groupby(["Age_Group", "gender", "city"])["order_value"].sum().reset_index()

    demographics.to_csv(os.path.join("data", "03_customer_demographics.csv"), index=False)
    print("Customer Demographics saved!")

# ---------------------------------------
# 4. Customer Loyalty & Retention
# ---------------------------------------
def customer_loyalty():
    customer_spend = orders.groupby("customer_id")["order_value"].sum().reset_index()
    merged = customers.merge(customer_spend, on="customer_id", how="left").fillna(0)

    loyalty = merged.groupby("loyalty_score")["order_value"].mean().reset_index()

    loyalty.to_csv(os.path.join("data", "04_customer_loyalty.csv"), index=False)
    print("Customer Loyalty saved!")

# ---------------------------------------
# 5. Review & Rating Insights
# ---------------------------------------
def review_rating_insights():
    merged = reviews.merge(orders, on="customer_id").merge(customers, on="customer_id")
    insights = merged.groupby("rating")["order_value"].mean().reset_index()

    insights.to_csv(os.path.join("data", "05_review_rating_insights.csv"), index=False)
    print("Review & Rating Insights saved!")

# ---------------------------------------
# 6. Customer Lifetime Value (CLV)
# ---------------------------------------
def customer_ltv():
    ltv = orders.groupby("customer_id").agg({
        "order_value": "sum",
        "order_id": "count"
    }).reset_index()

    # Round order_value
    ltv["order_value"] = ltv["order_value"].round(2)

    # Use total revenue as CLV score (rounded)
    ltv["CLV_Score"] = ltv["order_value"].round(2)

    # Rank customers into 3 categories
    ltv["CLV_Category"] = pd.qcut(ltv["CLV_Score"], 3, labels=["Low", "Medium", "High"])

    ltv.to_csv(os.path.join("data", "06_customer_ltv.csv"), index=False)
    print("Customer LTV saved!")

# ---------------------------------------
# 7. Churn Analysis
# ---------------------------------------
def churn_analysis():
    latest_date = orders["order_date"].max()
    churn = orders.groupby("customer_id")["order_date"].max().reset_index()
    churn["Days_Since_Last_Order"] = (latest_date - churn["order_date"]).dt.days
    churn["Status"] = churn["Days_Since_Last_Order"].apply(lambda x: "Inactive" if x > 90 else "Active")

    churn.to_csv(os.path.join("data", "07_churn_analysis.csv"), index=False)
    print("Churn Analysis saved!")

# ---------------------------------------
# Run all analyses
# ---------------------------------------
if __name__ == "__main__":
    customer_segmentation()
    top_products_categories()
    customer_demographics()
    customer_loyalty()
    review_rating_insights()
    customer_ltv()
    churn_analysis()
    print("\nAll project analyses completed and saved as CSV files in 'data/' folder!")
