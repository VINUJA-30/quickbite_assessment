# QuickBite Data Cleaning and Analysis

# 1. About the Project

This project is about cleaning and analysing QuickBite food delivery data using Python. The dataset contains information about orders, customers, restaurants, delivery partners, areas and customer feedback.

The main aim is to clean the data, combine the tables and find useful information about food delivery performance.

# 2. Understanding the Data

The Excel file contains 6 sheets:

- Areas – area and city details
- Restaurants – restaurant names and cuisine details
- Customers – customer information
- Orders – order details, payment and delivery information
- Delivery Partners – partner and vehicle details
- Feedback – customer ratings and comments

I understood that these tables are connected using IDs like customer_id, restaurant_id, partner_id and order_id.

# 3. Data Problems Found and Fixes

During data cleaning, I found some problems in the datasets:

- Duplicate records – removed duplicate rows.
- Duplicate IDs – checked duplicate primary IDs and handled them to avoid repeated records.
- Missing values – handled missing values in important columns.
- Different text formats – standardized payment modes, order status, cuisine, vehicle type and city names.
- ID format issues – removed extra spaces and converted IDs to uppercase.
- Date and amount formats – converted date columns into datetime format and amount columns into numeric format.
- Invalid ID relationships – checked whether the IDs in one table were available in the related tables.

After cleaning, I saved the cleaned datasets in an Excel file.

# 4. Combining the Data

I used left joins to combine the orders table with customers, restaurants, areas, delivery partners and feedback.

I also created these new columns:
- order_date – extracted from order_datetime.
- order_time – extracted from order_datetime.
- order_day – extracted from the date.
- is_late – marked as 1 when delivery time is more than 45 minutes, otherwise 0.

The final combined dataset contains 19 columns and one row for each order.

# 5. Analysis and How Results Were Calculated

# Area Performance
- Calculated total orders and total revenue for each area.
- Calculated average delivery time using delivered orders.
- Calculated late orders when delivery time is above 45 minutes.
- Calculated late percentage using valid delivery times.
- Sorted areas based on total revenue.

# Top Cuisine by Day
- Grouped orders by weekday and cuisine.
- Excluded cancelled orders.
- Counted orders for each cuisine on each day.
- Calculated the percentage of orders for each cuisine.
- Used dense ranking to identify the most popular cuisine for each day.

# Vehicle Performance
- Considered only delivered orders.
- Grouped orders based on vehicle type.
- Calculated total orders and average delivery time.
- Counted late deliveries and calculated late percentage.

# 6. Tools Used

- Python
- Pandas
- Excel
- OpenPyXL
- VS Code

# 7. Project Structure


quickbite-assessment/
│
├── data/
│   └── QuickBite_Assessment.xlsx
│
├── scripts/
│   ├── cleaning.py
│   └── analysis.py
│
├── cleaned_data/
│   ├── Quick_Bite_Data_Cleaning.xlsx
│   ├── Quick_Bite_Combined_Data.xlsx
│   └── ID_Validation_Report.csv
│
├── analysis/
│   └── Quick_Bite_Analysis.xlsx
│
└── README.md


# 8. How to Run the Project

Install the required libraries:

```bash
pip install pandas openpyxl
```

Run the cleaning script:

```bash
python scripts/cleaning.py
```

Run the analysis script:

```bash
python scripts/analysis.py
```

First run the cleaning script, then run the analysis script.

# 9. Output Files

- Quick_Bite_Data_Cleaning.xlsx – contains cleaned data from all six sheets.
- Quick_Bite_Combined_Data.xlsx – contains the final combined dataset.
- ID_Validation_Report.csv – contains the ID validation details.
- Quick_Bite_Analysis.xlsx – contains the three analysis results.

