# Unit 7 – IF Logic & Conditional Functions (IFS, SUMIF, COUNTIF, AVERAGEIF)

## Overview

This project demonstrates the use of Excel's conditional functions to analyze a retail supply chain sales dataset. A summary report was created using SUMIF, COUNTIF, AVERAGEIF, and IFS functions to extract meaningful business insights from the data.

## Dataset

- Dataset: Retail Supply Chain Sales Dataset
- File Format: Microsoft Excel (.xlsx)
- Sheets:
  - Retails Order Full Dataset
  - Calendar Table
  - Summary

## Objectives

The objective of this assignment was to:

- Calculate total sales by product category using SUMIF.
- Count the number of transactions by region using COUNTIF.
- Calculate the average sales for each customer segment using AVERAGEIF.
- Classify each order into Small, Medium, or Large categories using the IFS function.
- Organize all summary calculations in a dedicated Summary worksheet.

## Functions Used

### 1. SUMIF

Purpose:
Calculate the total sales for each product category.

Categories included:
- Furniture
- Office Supplies
- Technology

Example Formula:

=SUMIF(Category_Range, Category, Sales_Range)

---

### 2. COUNTIF

Purpose:
Count the total number of orders for each region.

Regions included:
- Central
- East
- South
- West

Example Formula:

=COUNTIF(Region_Range, Region)

---

### 3. AVERAGEIF

Purpose:
Calculate the average sales for each customer segment.

Segments included:
- Consumer
- Corporate
- Home Office

Example Formula:

=AVERAGEIF(Segment_Range, Segment, Sales_Range)

---

### 4. IFS

Purpose:
Classify every order based on its sales amount.

Classification:

- Small
- Medium
- Large

Example Formula:

=IFS(
Sales<100,"Small",
Sales<500,"Medium",
Sales>=500,"Large"
)

## Summary Sheet

The Summary worksheet contains:

- Total Sales by Category
- Number of Orders by Region
- Average Sales by Customer Segment

The dataset sheet also contains an additional "Order Size" column generated using the IFS function.

## Skills Demonstrated

- Conditional calculations using SUMIF
- Counting records using COUNTIF
- Conditional averages using AVERAGEIF
- Multi-condition logical testing using IFS
- Worksheet organization
- Business data summarization
- Excel formula referencing
##File Source
Retail-Supply-Chain-Sales-Dataset.xlsx
## Learning Outcomes

After completing this assignment, I learned how to:

- Analyze retail sales data efficiently.
- Summarize business information using conditional functions.
- Apply logical conditions to classify records.
- Create organized summary reports in Excel.
- Use formula references to automate calculations.
