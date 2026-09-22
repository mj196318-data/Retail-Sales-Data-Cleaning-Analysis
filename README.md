# 🧹 Retail Sales Data Cleaning & Analysis — Excel

## 📌 Project Overview

This project focuses on cleaning, validating, and analyzing a retail sales dataset using Microsoft Excel.

The dataset contains **1,250 retail transactions** with issues such as missing values, duplicate transaction IDs, inconsistent text formatting, mixed date formats, currency symbols, and invalid quantity/sales values.

The goal of this project was to transform the raw dataset into a clean and analysis-ready dataset and generate meaningful business insights using Excel.

---

## 🎯 Project Objectives

* Clean and standardize raw retail sales data
* Identify and handle missing values
* Detect and remove duplicate transactions
* Standardize customer names, cities, categories, products, and payment methods
* Convert sales values into usable numeric data
* Identify negative, zero, and blank quantities
* Validate sales and quantity fields
* Create KPI summaries and category-level analysis
* Prepare the dataset for further analysis and visualization

---

## 🛠️ Tools & Techniques

**Tool:** Microsoft Excel

**Techniques Used:**

* TRIM
* CLEAN
* PROPER
* IF
* XLOOKUP / VLOOKUP / HLOOKUP
* Pivot Tables
* Data Cleaning & Validation
* Duplicate Detection
* Date Standardization
* Currency-to-Number Conversion
* KPI Analysis
* Percentage of Total

---

## 🧹 Data Cleaning Performed

The raw dataset contained several data-quality issues.

### Text Cleaning

Standardized:

* Customer Names
* Cities
* Categories
* Product Names
* Payment Methods

Used Excel functions such as `TRIM`, `CLEAN`, and `PROPER`.

### Duplicate Handling

Duplicate transaction records were identified using **Transaction ID** and duplicate records were removed after validation.

### Sales Cleaning

Sales values containing currency symbols and commas were converted into proper numeric values.

### Quantity Validation

Quantity values were classified into:

* Positive / OK
* Zero
* Negative
* Blank

### Missing Value Analysis

| Data Quality Issue     | Count |
| ---------------------- | ----: |
| Missing Customer Names |    18 |
| Missing Cities         |    15 |
| Missing Sales          |    10 |
| Negative Sales         |     6 |
| Negative Quantity      |     5 |
| Zero Quantity          |     8 |
| Blank Quantity         |    12 |
| Missing Payment Method |    16 |

---

## 📊 Key KPIs

| KPI                            |         Value |
| ------------------------------ | ------------: |
| Total Transactions             |         1,250 |
| Total Sales                    | ₹4,959,963.80 |
| Total Cost                     | ₹3,621,121.00 |
| Total Profit                   | ₹1,338,843.00 |
| Total Quantity Sold            |         5,524 |
| Average Order Value            |     ₹3,967.97 |
| Profit Margin                  |        26.99% |
| Average Quantity / Transaction |          4.42 |

---

## 📈 Category Analysis

The cleaned dataset was analyzed using Pivot Tables to understand sales contribution by category.

| Category       |         Sales |
| -------------- | ------------: |
| Electronics    | ₹2,447,357.19 |
| Home & Kitchen |   ₹915,594.72 |
| Clothing       |   ₹781,285.25 |
| Beauty         |   ₹477,220.67 |
| Books          |   ₹338,505.95 |

**Electronics** generated the highest sales contribution in the dataset.

---

## 🏆 Top Products by Sales

| Product        |       Sales |
| -------------- | ----------: |
| Wireless Mouse | ₹493,383.41 |
| Headphones     | ₹455,804.35 |
| Power Bank     | ₹454,482.02 |

---

## 💡 Key Learnings

Through this project, I practiced:

* Real-world data cleaning in Excel
* Handling missing and inconsistent data
* Duplicate identification and removal
* Data validation
* Working with mixed date and currency formats
* Creating business KPIs
* Using Pivot Tables for analysis
* Calculating percentage contribution
* Preparing raw data for reporting and visualization

---

## 📁 Project Files

* `Retail_Sales_Data_Cleaning project.xlsx` — Excel workbook containing the raw and cleaned sales data and analysis.

---

## 👤 Author

**Mohd Jamal**

Aspiring Data Analyst | Excel | SQL | Power BI
