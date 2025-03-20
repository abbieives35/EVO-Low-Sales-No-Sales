# EVO-Low-Sales-No-Sales

# Customer Sales Order Analysis

## 📌 Overview
This Python project analyzes customer sales data to identify purchasing patterns, predict future orders, and flag potential sales opportunities. It processes an Excel file containing sales transactions and provides insights on customer-item relationships, typical order cycles, and late reorders.

## 📊 Features
- **Data Cleaning & Preparation**  
  - Strips whitespace from column names and values.  
  - Converts date fields to datetime format for accurate calculations.  

- **Order Frequency Analysis**  
  - Calculates **most recent order date** for each customer-item pair.  
  - Determines **typical item and class order cycles** by analyzing historical order intervals.  
  - Estimates the **next expected order date** for each item and class.

- **Sales Insights & Alerts**  
  - Computes total **order count, quantity, and revenue** per customer-item.  
  - Flags **"late" orders** where customers have missed their predicted order cycle.  
  - Extracts relevant metadata such as **sales reps and specialists assigned**.  

# How It Works
1. **Load & Clean Data**  
   - Reads an Excel file with sales data.  
   - Removes duplicates and standardizes column formats.  

2. **Summarize Customer-Item & Class Relationships**  
   - Groups sales records by customer, item, and class.  
   - Computes key order statistics, including total quantity, total sales, and order counts.  

3. **Predict Order Cycles**  
   - Calculates **typical order intervals** based on past purchases.  
   - Determines **expected future order dates** based on historical trends.  

4. **Identify Late Reorders**  
   - Compares actual order timing with expected patterns.  
   - Flags customers whose purchasing patterns have deviated.  



