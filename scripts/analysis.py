#(Task 3)

import pandas as pd
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = (
    BASE_DIR
    / "cleaned_data"
    / "Quick_Bite_Combined_Data.xlsx"
)

OUTPUT_DIR = BASE_DIR / "analysis"
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "Quick_Bite_Analysis.xlsx"

# 2. Read the combined data
data = pd.read_excel(INPUT_FILE)

print("Combined dataset loaded successfully.")
print("Dataset shape:", data.shape)

# 3. Validate the combined data

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

if list(data.columns) != required_columns:
    raise ValueError("Combined dataset does not contain the required 19 columns.")

if data["order_id"].duplicated().any():
    raise ValueError("Duplicate order IDs found. Expected one row per order.")

print("Required columns validation: PASSED")
print("One row per order validation: PASSED")


# 4. Standaritize the data

data["order_amount"] = pd.to_numeric(
    data["order_amount"],
    errors="coerce"
)

data["delivery_minutes"] = pd.to_numeric(
    data["delivery_minutes"],
    errors="coerce"
)

data["order_status"] = (
    data["order_status"]
    .astype("string")
    .str.strip()
    .str.lower()
)

data["order_day"] = (
    data["order_day"]
    .astype("string")
    .str.strip()
    .str.title()
)

# Unknown cuisine is kept as a category so that
# daily cuisine percentages account for every order.
data["cuisine"] = (
    data["cuisine"]
    .fillna("Unknown")
    .astype("string")
    .str.strip()
    .str.lower()
)

data["area_name"] = data["area_name"].fillna("Unknown")
data["city"] = data["city"].fillna("Unknown")
data["vehicle_type"] = data["vehicle_type"].fillna("Unknown")


# 5. Define the buisness population


delivered = data[
    data["order_status"] == "delivered"
].copy()


non_cancelled = data[
    data["order_status"] != "cancelled"
].copy()


vehicle_data = delivered.copy()


print("\nBUSINESS POPULATIONS")
print("Delivered orders:", len(delivered))
print("Non-cancelled orders:", len(non_cancelled))


# 6. Area performance

# A delivery time is valid only when it is available.
area_data = delivered.copy()

area_data["valid_delivery_time"] = (
    area_data["delivery_minutes"].notna()
)

area_data["late_delivery"] = (
    area_data["delivery_minutes"] > 45
).fillna(False)


area_performance = (
    area_data
    .groupby(
        ["area_name", "city"],
        dropna=False
    )
    .agg(
        total_orders=("order_id", "nunique"),
        total_revenue=("order_amount", "sum"),
        average_delivery_minutes=("delivery_minutes", "mean"),
        valid_delivery_time_orders=("valid_delivery_time", "sum"),
        late_orders=("late_delivery", "sum")
    )
    .reset_index()
)


# Late percentage denominator:
# Delivered orders with valid delivery times only.

area_performance["late_percentage"] = (
    area_performance["late_orders"]
    / area_performance["valid_delivery_time_orders"]
    .replace(0, pd.NA)
    * 100
)


# Rank areas by total revenue, highest first.

area_performance = area_performance.sort_values(
    by="total_revenue",
    ascending=False
).reset_index(drop=True)


area_performance["total_revenue"] = (
    area_performance["total_revenue"].round(2)
)

area_performance["average_delivery_minutes"] = (
    area_performance["average_delivery_minutes"].round(2)
)

area_performance["late_percentage"] = (
    area_performance["late_percentage"].round(2)
)


print("\nAREA PERFORMANCE")
print(area_performance.to_string(index=False))


# --------------------------------------------------
# 7. TOP CUISINE BY DAY
# --------------------------------------------------

# Count orders for each cuisine on each weekday.

cuisine_counts = (
    non_cancelled
    .groupby(
        ["order_day", "cuisine"],
        dropna=False
    )
    .agg(
        total_orders=("order_id", "nunique")
    )
    .reset_index()
)


# Calculate the total orders on each day.

cuisine_counts["daily_total_orders"] = (
    cuisine_counts
    .groupby("order_day")["total_orders"]
    .transform("sum")
)


# Calculate cuisine share as a percentage of that day's orders.

cuisine_counts["cuisine_share_percentage"] = (
    cuisine_counts["total_orders"]
    / cuisine_counts["daily_total_orders"]
    * 100
)



cuisine_counts["rank"] = (
    cuisine_counts
    .groupby("order_day")["total_orders"]
    .rank(
        method="dense",
        ascending=False
    )
    .astype(int)
)


# Mark the top cuisine or cuisines for each day.

cuisine_counts["is_top_cuisine"] = (
    cuisine_counts["rank"] == 1
)


# Keep all cuisines in the report so that
# percentages across each day can be checked.

weekday_order = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
]

cuisine_counts["order_day"] = pd.Categorical(
    cuisine_counts["order_day"],
    categories=weekday_order,
    ordered=True
)

cuisine_counts = cuisine_counts.sort_values(
    ["order_day", "rank", "cuisine"]
).reset_index(drop=True)


cuisine_counts["cuisine_share_percentage"] = (
    cuisine_counts["cuisine_share_percentage"].round(2)
)


print("\nTOP CUISINE BY DAY")
print(cuisine_counts.to_string(index=False))


# Validate daily cuisine shares before rounding.

daily_shares = (
    cuisine_counts
    .groupby("order_day", observed=True)["cuisine_share_percentage"]
    .sum()
)

if not daily_shares.between(99.95, 100.05).all():
    raise ValueError(
        "Daily cuisine percentages do not add up to approximately 100%."
    )

print("\nDaily cuisine percentage validation: PASSED")


# 8. VEHICLE PERFORMANCE

vehicle_data = delivered.copy()

vehicle_data["valid_delivery_time"] = (
    vehicle_data["delivery_minutes"].notna()
)

vehicle_data["late_delivery"] = (
    vehicle_data["delivery_minutes"] > 45
).fillna(False)


vehicle_performance = (
    vehicle_data
    .groupby("vehicle_type", dropna=False)
    .agg(
        total_orders=("order_id", "nunique"),
        valid_delivery_time_orders=("valid_delivery_time", "sum"),
        average_delivery_minutes=("delivery_minutes", "mean"),
        late_orders=("late_delivery", "sum")
    )
    .reset_index()
)


# Unknown delivery times are excluded from the denominator.

vehicle_performance["late_percentage"] = (
    vehicle_performance["late_orders"]
    / vehicle_performance["valid_delivery_time_orders"]
    .replace(0, pd.NA)
    * 100
)


vehicle_performance = vehicle_performance.sort_values(
    by="total_orders",
    ascending=False
).reset_index(drop=True)


vehicle_performance["average_delivery_minutes"] = (
    vehicle_performance["average_delivery_minutes"].round(2)
)

vehicle_performance["late_percentage"] = (
    vehicle_performance["late_percentage"].round(2)
)


print("\nVEHICLE PERFORMANCE")
print(vehicle_performance.to_string(index=False))


# 9. VALIDATE VEHICLE PERFORMANCE


if not (vehicle_data["order_status"] == "delivered").all():
    raise ValueError("Vehicle report contains non-delivered orders.")

if (vehicle_performance["late_orders"] >
        vehicle_performance["valid_delivery_time_orders"]).any():
    raise ValueError("Late orders exceed valid delivery-time orders.")

print("\nVehicle performance validation: PASSED")


# 10. VALIDATE AREA REVENUE RANKING

if not area_performance["total_revenue"].is_monotonic_decreasing:
    raise ValueError("Area revenue ranking is not descending.")

print("Area revenue ranking validation: PASSED")


# 11. SAVE THE THREE ANALYSIS WORKSHEETS

try:

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        area_performance.to_excel(
            writer,
            sheet_name="Area Performance",
            index=False
        )

        cuisine_counts.to_excel(
            writer,
            sheet_name="Top Cuisine by Day",
            index=False
        )

        vehicle_performance.to_excel(
            writer,
            sheet_name="Vehicle Performance",
            index=False
        )

except PermissionError:

    raise PermissionError(
        "Cannot save the analysis workbook. "
        "Close Quick_Bite_Analysis.xlsx in Excel and run the script again."
    )


# 12. Final result

print("DATA ANALYSIS COMPLETED SUCCESSFULLY")

print("\nOutput file:")
print(OUTPUT_FILE)

print("\nWorksheets created:")
print("1. Area Performance")
print("2. Top Cuisine by Day")
print("3. Vehicle Performance")
