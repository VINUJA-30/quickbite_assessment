
#(Task 1 + Task 2)
import pandas as pd
from pathlib import Path

# 1. File path

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = BASE_DIR / "data" / "QuickBite_Assessment.xlsx"
OUTPUT_DIR = BASE_DIR / "cleaned_data"

OUTPUT_DIR.mkdir(exist_ok=True)

CLEANED_FILE = OUTPUT_DIR / "Quick_Bite_Data_Cleaning.xlsx"
COMBINED_FILE = OUTPUT_DIR / "Quick_Bite_Combined_Data.xlsx"
VALIDATION_FILE = OUTPUT_DIR / "ID_Validation_Report.csv"

issues = []


def add_issue(table, issue, column="", value="", action=""):
    """Store data-quality issues for the validation report."""
    issues.append({
        "table": table,
        "issue": issue,
        "column": column,
        "value": value,
        "action": action
    })


# 2. Read the original workbook

areas = pd.read_excel(INPUT_FILE, sheet_name="areas")
restaurants = pd.read_excel(INPUT_FILE, sheet_name="restaurants")
customers = pd.read_excel(INPUT_FILE, sheet_name="customers")
orders = pd.read_excel(INPUT_FILE, sheet_name="orders")
partners = pd.read_excel(INPUT_FILE, sheet_name="delivery_partners")
feedback = pd.read_excel(INPUT_FILE, sheet_name="feedback")

print("Original data loaded successfully.")


# 3. Standaridize all foreign key and primary key

def normalize_id(data, columns):
    for column in columns:
        data[column] = (
            data[column]
            .astype("string")
            .str.strip()
            .str.upper()
        )

    return data


areas = normalize_id(areas, ["area_id"])

restaurants = normalize_id(
    restaurants,
    ["restaurant_id", "area_id"]
)

customers = normalize_id(customers, ["customer_id"])

orders = normalize_id(
    orders,
    [
        "order_id",
        "customer_id",
        "restaurant_id",
        "partner_id"
    ]
)

partners = normalize_id(partners, ["partner_id"])

feedback = normalize_id(
    feedback,
    ["feedback_id", "order_id"]
)

print("All related IDs standardized.")


# 4. Remove duplicate rows

tables = {
    "areas": areas,
    "restaurants": restaurants,
    "customers": customers,
    "orders": orders,
    "delivery_partners": partners,
    "feedback": feedback
}

for table_name, data in tables.items():

    duplicate_count = int(data.duplicated().sum())

    if duplicate_count > 0:
        add_issue(
            table_name,
            "Exact duplicate rows",
            action=f"Removed {duplicate_count} exact duplicate rows"
        )

    data.drop_duplicates(inplace=True)


# 5. Check primary keys

primary_keys = {
    "areas": (areas, "area_id"),
    "restaurants": (restaurants, "restaurant_id"),
    "customers": (customers, "customer_id"),
    "orders": (orders, "order_id"),
    "delivery_partners": (partners, "partner_id"),
    "feedback": (feedback, "feedback_id")
}


for table_name, (data, key) in primary_keys.items():

    # Report missing primary keys
    missing_count = int(data[key].isna().sum())

    if missing_count > 0:
        add_issue(
            table_name,
            "Missing primary key",
            column=key,
            value=missing_count,
            action="Removed rows because they cannot be uniquely identified"
        )

        data.dropna(subset=[key], inplace=True)

    # Report duplicate primary keys
    duplicate_mask = data.duplicated(subset=[key], keep="first")
    duplicate_count = int(duplicate_mask.sum())

    if duplicate_count > 0:

        duplicate_values = data.loc[duplicate_mask, key].tolist()

        for value in duplicate_values:
            add_issue(
                table_name,
                "Duplicate primary key",
                column=key,
                value=value,
                action="Kept first occurrence; later occurrences removed"
            )

        data.drop_duplicates(subset=[key], keep="first", inplace=True)


# 6. Remove rows with missing mandatory values

mandatory_columns = {
    "areas": (areas, ["area_id", "area_name", "city"]),
    "restaurants": (
        restaurants,
        ["restaurant_id", "restaurant_name", "area_id", "cuisine"]
    ),
    "customers": (customers, ["customer_id", "customer_name"]),
    "delivery_partners": (
        partners,
        ["partner_id", "partner_name", "vehicle_type"]
    ),
    "orders": (
        orders,
        [
            "order_id",
            "customer_id",
            "restaurant_id",
            "order_datetime",
            "order_amount",
            "payment_mode",
            "order_status"
        ]
    ),
    "feedback": (feedback, ["feedback_id", "order_id"])
}


for table_name, (data, columns) in mandatory_columns.items():

    missing_mask = data[columns].isna().any(axis=1)
    missing_count = int(missing_mask.sum())

    if missing_count > 0:
        add_issue(
            table_name,
            "Missing mandatory values",
            column=", ".join(columns),
            value=missing_count,
            action="Removed rows with missing mandatory values"
        )

        data.dropna(subset=columns, inplace=True)


# 7. Standardize area names and cities

areas["area_name"] = (
    areas["area_name"]
    .astype("string")
    .str.strip()
)

areas["city"] = (
    areas["city"]
    .astype("string")
    .str.strip()
    .str.lower()
)

areas["city"] = areas["city"].replace({
    "blr": "Bengaluru",
    "bengaluru": "Bengaluru",
    "hyd": "Hyderabad",
    "hyderabad": "Hyderabad",
    "madras": "Chennai",
    "chennai": "Chennai",
    "mumbai": "Mumbai"
})


# 8. Standardize cuisine types

restaurants["cuisine"] = (
    restaurants["cuisine"]
    .astype("string")
    .str.strip()
    .str.lower()
)

restaurants["cuisine"] = restaurants["cuisine"].replace({
    "biriyani": "biryani",
    "n. indian": "north indian",
    "north-indian": "north indian",
    "south-indian": "south indian",
    "dessert": "desserts",
    "italian - pizza": "pizza"
})


# --------------------------------------------------
# 9. STANDARDIZE VEHICLE TYPES
# --------------------------------------------------

def standardize_vehicle(vehicle):

    vehicle = str(vehicle).strip().lower()
    vehicle = vehicle.replace(" ", "")
    vehicle = vehicle.replace("-", "")

    if vehicle in ["bike", "motorbike"]:
        return "Motorbike"

    elif vehicle in ["cycle", "bicycle"]:
        return "Bicycle"

    elif vehicle in ["activa", "scooty", "scooter"]:
        return "Scooter"

    elif vehicle in [
        "ev",
        "evscooter",
        "escooter",
        "electricscooter"
    ]:
        return "Electric Scooter"

    return vehicle.title()


partners["vehicle_type"] = partners["vehicle_type"].apply(
    standardize_vehicle
)


# 10. STANDARDIZE PAYMENT MODES

orders["payment_mode"] = (
    orders["payment_mode"]
    .astype("string")
    .str.strip()
    .str.lower()
)

orders["payment_mode"] = orders["payment_mode"].replace({
    "gpay": "UPI",
    "phonepe": "UPI",
    "upi": "UPI",
    "credit card": "Card",
    "debit card": "Card",
    "card": "Card",
    "cod": "Cash",
    "cash on delivery": "Cash",
    "cash": "Cash",
    "amazon pay": "Wallet",
    "paytm wallet": "Wallet",
    "wallet": "Wallet"
})


# 11. Standardize order status

orders["order_status"] = (
    orders["order_status"]
    .astype("string")
    .str.strip()
    .str.lower()
)

orders["order_status"] = orders["order_status"].replace({
    "completed": "Delivered",
    "delivered": "Delivered",
    "cancelled": "Cancelled",
    "canceled": "Cancelled"
})

# Format other statuses consistently
orders["order_status"] = (
    orders["order_status"]
    .str.title()
)


# 12. Clean coupon codes

orders["coupon_code"] = (
    orders["coupon_code"]
    .fillna("")
    .astype("string")
    .str.strip()
    .str.upper()
)

# 13. Convert numeric and datetime columns

orders["order_amount"] = (
    orders["order_amount"]
    .astype("string")
    .str.replace(",", "", regex=False)
    .str.replace(r"[^\d.-]", "", regex=True)
)

orders["order_amount"] = pd.to_numeric(
    orders["order_amount"],
    errors="coerce"
)

invalid_amounts = int(orders["order_amount"].isna().sum())

if invalid_amounts > 0:
    add_issue(
        "orders",
        "Invalid order amounts",
        column="order_amount",
        value=invalid_amounts,
        action="Removed rows with invalid order amounts"
    )

    orders.dropna(subset=["order_amount"], inplace=True)


orders["delivery_minutes"] = pd.to_numeric(
    orders["delivery_minutes"],
    errors="coerce"
)

orders["order_datetime"] = pd.to_datetime(
    orders["order_datetime"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)

invalid_dates = int(orders["order_datetime"].isna().sum())

if invalid_dates > 0:
    add_issue(
        "orders",
        "Invalid order dates",
        column="order_datetime",
        value=invalid_dates,
        action="Removed rows with invalid order dates"
    )

    orders.dropna(subset=["order_datetime"], inplace=True)


# 14. Validate foreign key relationships
def check_foreign_key(child, child_name, child_column,
                      parent, parent_name, parent_column):

    parent_ids = set(parent[parent_column].dropna())

    unmatched_mask = (
        child[child_column].isna()
        | ~child[child_column].isin(parent_ids)
    )

    unmatched_rows = child.loc[unmatched_mask]

    for value in unmatched_rows[child_column].tolist():

        add_issue(
            child_name,
            f"Unmatched reference to {parent_name}",
            column=child_column,
            value=value,
            action="Retained child row; unmatched parent details will be missing after left join"
        )

    return len(unmatched_rows)


relationships = [
    (orders, "orders", "customer_id", customers, "customers", "customer_id"),
    (orders, "orders", "restaurant_id", restaurants, "restaurants", "restaurant_id"),
    (orders, "orders", "partner_id", partners, "delivery_partners", "partner_id"),
    (restaurants, "restaurants", "area_id", areas, "areas", "area_id"),
    (feedback, "feedback", "order_id", orders, "orders", "order_id")
]

print("\nFOREIGN KEY VALIDATION")

for relationship in relationships:

    unmatched_count = check_foreign_key(*relationship)

    print(
        relationship[1],
        relationship[2],
        "->",
        relationship[4],
        ":",
        unmatched_count,
        "unmatched records"
    )

# 15. Save the cleaned data

with pd.ExcelWriter(CLEANED_FILE, engine="openpyxl") as writer:

    areas.to_excel(writer, sheet_name="areas", index=False)
    restaurants.to_excel(writer, sheet_name="restaurants", index=False)
    customers.to_excel(writer, sheet_name="customers", index=False)
    orders.to_excel(writer, sheet_name="orders", index=False)
    partners.to_excel(writer, sheet_name="delivery_partners", index=False)
    feedback.to_excel(writer, sheet_name="feedback", index=False)

print("\nSix cleaned worksheets saved.")


# 16. PREPARE FEEDBACK FOR ONE ROW PER ORDER

feedback["rating"] = pd.to_numeric(
    feedback["rating"],
    errors="coerce"
)

feedback["comment"] = (
    feedback["comment"]
    .fillna("")
    .astype("string")
    .str.strip()
)


def combine_comments(values):

    comments = [
        value for value in values
        if pd.notna(value) and str(value).strip() != ""
    ]

    return " | ".join(dict.fromkeys(comments))


feedback_summary = (
    feedback.groupby("order_id", as_index=False)
    .agg(
        rating=("rating", "mean"),
        comment=("comment", combine_comments)
    )
)

# 17. MERGE THE CLEANED TABLES

combined = orders.merge(
    customers[["customer_id", "customer_name"]],
    on="customer_id",
    how="left",
    validate="many_to_one",
    indicator="_customer_merge"
)

print(
    "Unmatched customer merges:",
    int((combined["_customer_merge"] == "left_only").sum())
)

combined.drop(columns="_customer_merge", inplace=True)


combined = combined.merge(
    restaurants[
        ["restaurant_id", "restaurant_name", "cuisine", "area_id"]
    ],
    on="restaurant_id",
    how="left",
    validate="many_to_one",
    indicator="_restaurant_merge"
)

print(
    "Unmatched restaurant merges:",
    int((combined["_restaurant_merge"] == "left_only").sum())
)

combined.drop(columns="_restaurant_merge", inplace=True)


combined = combined.merge(
    areas[["area_id", "area_name", "city"]],
    on="area_id",
    how="left",
    validate="many_to_one",
    indicator="_area_merge"
)

print(
    "Unmatched area merges:",
    int((combined["_area_merge"] == "left_only").sum())
)

combined.drop(columns="_area_merge", inplace=True)


combined = combined.merge(
    partners[["partner_id", "partner_name", "vehicle_type"]],
    on="partner_id",
    how="left",
    validate="many_to_one",
    indicator="_partner_merge"
)

print(
    "Unmatched delivery partner merges:",
    int((combined["_partner_merge"] == "left_only").sum())
)

combined.drop(columns="_partner_merge", inplace=True)


combined = combined.merge(
    feedback_summary,
    on="order_id",
    how="left",
    validate="many_to_one",
    indicator="_feedback_merge"
)

print(
    "Orders without feedback:",
    int((combined["_feedback_merge"] == "left_only").sum())
)

combined.drop(columns="_feedback_merge", inplace=True)


# 18. CREATE DATE, TIME, DAY AND LATE COLUMNS

combined["order_date"] = combined["order_datetime"].dt.date
combined["order_time"] = combined["order_datetime"].dt.time
combined["order_day"] = combined["order_datetime"].dt.day_name()

combined["is_late"] = (
    combined["delivery_minutes"] > 45
).astype("boolean")


# 19. SELECT THE REQUIRED 19 COLUMNS

required_columns = [
    "order_id",
    "order_date",
    "order_time",
    "order_day",
    "customer_name",
    "restaurant_name",
    "cuisine",
    "area_name",
    "city",
    "partner_name",
    "vehicle_type",
    "order_amount",
    "payment_mode",
    "coupon_code",
    "order_status",
    "delivery_minutes",
    "is_late",
    "rating",
    "comment"
]

combined = combined[required_columns]


# 20. VALIDATE THE COMBINED DATASET

print("\nCOMBINED DATA VALIDATION")

print("Combined rows:", len(combined))
print("Distinct order IDs:", combined["order_id"].nunique())
print("Required columns:", len(combined.columns))

if combined["order_id"].duplicated().any():
    raise ValueError("Duplicate order IDs found in combined dataset.")

if len(combined) != combined["order_id"].nunique():
    raise ValueError("One-row-per-order validation failed.")

if list(combined.columns) != required_columns:
    raise ValueError("The required 19 columns do not match.")

# Reconcile delivered revenue with the source orders
source_revenue = orders.loc[
    orders["order_status"].eq("Delivered"),
    "order_amount"
].sum()

combined_revenue = combined.loc[
    combined["order_status"].eq("Delivered"),
    "order_amount"
].sum()

print("Delivered source revenue:", round(source_revenue, 2))
print("Delivered combined revenue:", round(combined_revenue, 2))

if not abs(source_revenue - combined_revenue) < 0.01:
    raise ValueError("Revenue reconciliation failed.")

print("One row per order: PASSED")
print("Required 19 columns: PASSED")
print("Revenue reconciliation: PASSED")


# 21. SAVE THE COMBINED DATASET

combined.to_excel(
    COMBINED_FILE,
    index=False,
    engine="openpyxl"
)


# 22. save the validation report

validation_report = pd.DataFrame(issues)

if validation_report.empty:
    validation_report = pd.DataFrame(
        columns=["table", "issue", "column", "value", "action"]
    )

validation_report.to_csv(
    VALIDATION_FILE,
    index=False
)


# 23. DISPLAY FINAL RESULTS

print("\nDATA CLEANING AND MERGING COMPLETED!")

print("Cleaned workbook:", CLEANED_FILE)
print("Combined workbook:", COMBINED_FILE)
print("Validation report:", VALIDATION_FILE)

print("\nFinal combined dataset shape:", combined.shape)
print("\nVehicle types:")
print(partners["vehicle_type"].value_counts())

print("\nValidation issues recorded:", len(validation_report))
