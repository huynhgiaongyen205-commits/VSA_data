import pandas as pd
from pathlib import Path

FILE_PATH = "report_2026-01-30_18-03-01.xlsx"
output_dir = Path("output_data")
output_dir.mkdir(exist_ok=True)

# Load raw excel (no header)
raw = pd.read_excel(FILE_PATH, header=None)

print("Loaded shape:", raw.shape)

# =========================
# 1. BOOTH VISIT COUNTS
# =========================
booth_visits = raw.iloc[0:4, 0:2]
booth_visits.columns = ["booth_type", "visit_count"]
booth_visits = booth_visits.dropna()

# =========================
# 2. VISIT ORDER
# =========================
visit_order_row = raw[raw[0] == "Visit Order of Booths"].index[0]
visit_order = raw.iloc[visit_order_row + 1, 0]
visit_sequence = [x.strip() for x in visit_order.split("->")]

# =========================
# 3. PRODUCT TIME DATA 
# =========================
product_time_start = raw[raw[0] == "Product"].index[0] + 1
product_time_end = raw.iloc[product_time_start:].dropna(how="all").index[-1]

product_time = raw.iloc[product_time_start:product_time_end + 1, 0:2]
product_time.columns = ["product", "time_spent"]
product_time = product_time.dropna(subset=["product"])
product_time["time_spent"] = pd.to_numeric(
    product_time["time_spent"], errors="coerce"
)
MIN_PRODUCT_TIME_SEC = 60
EXCLUDE_PRODUCTS = [
    "product",
    "total items",
    "total price",
]

product_time_clean = product_time[
    product_time["product"].notna() &
    product_time["time_spent"].notna() &
    (product_time["time_spent"] >= MIN_PRODUCT_TIME_SEC) &
    (~product_time["product"].astype(str).str.lower().isin(EXCLUDE_PRODUCTS))
]

total_task_time_sec = int(product_time_clean["time_spent"].sum())
# =========================
# 4. SHOPPING TABLE
# =========================
shopping_start = raw[raw[0] == "Product Name"].index[0] + 1
shopping_table = raw.iloc[shopping_start:shopping_start + 5, 0:5]
shopping_table.columns = [
    "product_name",
    "total_count",
    "unit_price",
    "status",
    "subtotal"
]
shopping_table = shopping_table.dropna(subset=["product_name"])

# =========================
# 5. SUMMARY METRICS
# =========================
def find_value(keyword):
    matches = raw[
        raw.astype(str)
        .apply(lambda row: row.str.contains(keyword, case=False, na=False).any(), axis=1)
    ]
    if not matches.empty:
        return matches.iloc[0, 1]
    return None


total_items = find_value("total items")
total_price = find_value("total price")
show_list_count = find_value("show list")

summary = {
    "total_items": total_items,
    "total_price": total_price,
    "show_list_count": show_list_count,
    "total_task_time_sec": total_task_time_sec
}

print("\n=== SUMMARY METRICS ===")
print(summary)

# =========================
# 6. SAVE OUTPUT
# =========================
booth_visits.to_csv(output_dir / "booth_visits.csv", index=False)
product_time.to_csv(output_dir / "product_time.csv", index=False)
shopping_table.to_csv(output_dir / "shopping_table.csv", index=False)

pd.DataFrame({"visit_sequence": visit_sequence}).to_csv(
    output_dir / "visit_order.csv", index=False
)

pd.DataFrame([summary]).to_csv(
    output_dir / "summary_metrics.csv", index=False
)

print("\n✅ DATA EXTRACTED SUCCESSFULLY")
print("Files saved in folder: /output")