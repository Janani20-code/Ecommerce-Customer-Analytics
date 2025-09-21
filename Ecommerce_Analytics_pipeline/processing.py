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
        "order_date": lambda x: (orders["order_date"].max() - x.max()).days,  # Recency
        "order_id": "count",                                                 # Frequency
        "order_value": "sum"                                                 # Monetary
    }).reset_index()
    rfm.columns = ["customer_id", "Recency", "Frequency", "Monetary"]

    # Round Monetary to 2 decimal places
    rfm["Monetary"] = rfm["Monetary"].round(2)

    # Dynamically assign VIP threshold using 75th percentile
    vip_threshold = rfm["Monetary"].quantile(0.75)

    # Simple segmentation rules
    rfm["Segment"] = "Regular"
    rfm.loc[(rfm["Recency"] < 30) & (rfm["Frequency"] > 3) & (rfm["Monetary"] >= vip_threshold), "Segment"] = "VIP"
    rfm.loc[(rfm["Recency"] < 60) & (rfm["Frequency"] >= 2), "Segment"] = "Loyal"
    rfm.loc[rfm["Recency"] > 90, "Segment"] = "Lost"

    # Optional: RFM Score
    rfm["RFM_Score"] = rfm["Recency"].rank(ascending=False) + rfm["Frequency"].rank() + rfm["Monetary"].rank()

    rfm.to_csv(os.path.join("data", "01_customer_segmentation.csv"), index=False)
    print("Customer Segmentation saved!")

# ---------------------------------------
# 2. Top Categories
# ---------------------------------------
def top_categories():
    top_categories = orders.groupby("product_category")["order_value"].sum().reset_index()
    top_categories["order_value"] = top_categories["order_value"].round(2)
    top_categories = top_categories.sort_values(by="order_value", ascending=False)

    # Optional: Percent of total revenue
    total_rev = top_categories["order_value"].sum()
    top_categories["Percent_of_Total"] = ((top_categories["order_value"] / total_rev) * 100).round(2)

    top_categories.to_csv(os.path.join("data", "02_top_categories.csv"), index=False)
    print("Top Categories saved!")

# ---------------------------------------
# ---------------------------------------
# 3. Customer Demographics & Spending
# ---------------------------------------
def customer_demographics():
    # Total spend per customer
    customer_spend = orders.groupby("customer_id")["order_value"].sum().reset_index()

    # Merge with customer info
    merged = customers.merge(customer_spend, on="customer_id", how="left").fillna(0)

    # Create Age group buckets
    bins = [0, 25, 40, 60, 100]
    labels = ["Young", "Adult", "Middle-aged", "Senior"]
    merged["Age_Group"] = pd.cut(merged["age"], bins=bins, labels=labels, right=False)

    # -----------------------
    # 1️⃣ Detailed demographics: Age + Gender + City
    # -----------------------
    demographics_detailed = merged.groupby(["Age_Group", "gender", "city"], observed=False)[
        "order_value"].sum().reset_index()
    demographics_detailed["order_value"] = demographics_detailed["order_value"].round(2)
    demographics_detailed.to_csv(os.path.join("data", "03_customer_demographics_detailed.csv"), index=False)

    # -----------------------
    # 2️⃣ Summary by Age_Group for charts
    # -----------------------
    demographics_summary = merged.groupby("Age_Group")["order_value"].sum().reset_index()
    demographics_summary["order_value"] = demographics_summary["order_value"].round(2)

    # Optional: Add percentage of total
    total_spend = demographics_summary["order_value"].sum()
    demographics_summary["Percent_of_Total"] = (demographics_summary["order_value"] / total_spend * 100).round(2)

    demographics_summary.to_csv(os.path.join("data", "03_customer_demographics_summary.csv"), index=False)

    print("Customer Demographics saved! Detailed & Summary CSVs created.")


# ---------------------------------------
# 4. Customer Loyalty & Retention
# ---------------------------------------
def customer_loyalty():
    customer_spend = orders.groupby("customer_id")["order_value"].sum().reset_index()
    merged = customers.merge(customer_spend, on="customer_id", how="left").fillna(0)

    loyalty = merged.groupby("loyalty_score")["order_value"].mean().reset_index()
    loyalty["order_value"] = loyalty["order_value"].round(2)

    # Optional: Loyalty Level
    loyalty["Loyalty_Level"] = pd.qcut(loyalty["order_value"], 3, labels=["Low", "Medium", "High"])

    loyalty.to_csv(os.path.join("data", "04_customer_loyalty.csv"), index=False)
    print("Customer Loyalty saved!")

# ---------------------------------------
# 5. Review & Rating Insights
# ---------------------------------------
def review_rating_insights():
    merged = reviews.merge(orders, on="customer_id").merge(customers, on="customer_id")
    insights = merged.groupby("rating")["order_value"].mean().reset_index()
    insights["order_value"] = insights["order_value"].round(2)

    # Optional: count of reviews per rating
    review_count = merged.groupby("rating")["review_id"].count().reset_index().rename(columns={"review_id": "review_count"})
    insights = insights.merge(review_count, on="rating")

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

    ltv["order_value"] = ltv["order_value"].round(2)
    ltv["CLV_Score"] = ltv["order_value"]

    # Rank customers into 3 categories
    ltv["CLV_Category"] = pd.qcut(ltv["CLV_Score"], 3, labels=["Low", "Medium", "High"])

    # Optional: Add Frequency
    ltv = ltv.rename(columns={"order_id": "Frequency"})

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
    customer_demographics()
    top_categories()
    customer_loyalty()
    review_rating_insights()
    customer_ltv()
    churn_analysis()
    print("\nAll project analyses completed and saved as CSV files in 'data/' folder!")
