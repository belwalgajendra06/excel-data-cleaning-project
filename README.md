# Excel Data Cleaning Project - E-commerce Sales Data

## 📌 Project Overview
This project focuses on cleaning a raw E-commerce sales dataset containing 1000+ rows. The data had duplicate entries, inconsistent text formats, blank values, and incorrect data types. I cleaned it using advanced Excel formulas to make it analysis-ready.

## 🔧 Dataset Details
- Total Rows: 1000+
- Key Columns: Customer Name, Email, Product, City, Region, Quantity, Discount, Payment Method, Order Status, Date

## 🧹 Data Cleaning Steps I Performed

1.  **Duplicate Handling:** Identified duplicate orders and removed them using Remove Duplicates.
2.  **Date Format Fix:** Used `DATEVALUE()` function and changed format to `DD-MM-YYYY` to fix text dates.
3.  **Customer Name Cleaning:** Used `PROPER()` function to convert names to proper case (e.g., ram kumar -> Ram Kumar).
4.  **Email Cleaning:** Used `TRIM()` function to remove extra spaces from Email column.
5.  **City & Region Cleaning:** Used `PROPER()` + `TRIM()` combination to standardize City (KANPUR -> Kanpur) and Region.
6.  **Quantity Cleaning:** Used `VALUE()` function to convert Quantity from text to number format.
7.  **Discount Cleaning:** Used `IF()` formula to check blank cells and filled empty discounts with `0`.
8.  **Payment Method Cleaning:** First used `PROPER()`, then used Find & Replace to correct `Upi` to `UPI` for consistency.
9.  **Order Status Cleaning:** Used `PROPER()` to standardize status (e.g., delivered -> Delivered).

## 🛠️ Tools & Functions Used
- Microsoft Excel
- Functions: `TRIM(), PROPER(), DATEVALUE(), VALUE(), IF()`
- Features: Remove Duplicates, Find & Replace, Format Cells

## 📁 Files in Repository
- `Excel_DataCleaning_1000Rows.xlsx` - Final Cleaned & Standardized Dataset

## ✅ Final Result
The final dataset is now 100% clean with no duplicates, proper date formats, consistent text casing, correct number formats, and no blank values in critical columns. It is now ready for Dashboard and Analysis.

## 👤 Author
Gajendra Belwal | MA Post Graduate(fresher)| Aspiring Data Analyst & MIS executive
