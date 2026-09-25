# SWYNEX - Data Clean Preparation

## 📌 Project Overview

**SWYNEX - Data Clean Preparation** is an Excel-based data cleaning and preparation project focused on identifying and correcting common data-quality issues in a raw dataset.

The project uses **Microsoft Excel and VBA macros** to automate data cleaning, validation, standardization, and basic analysis while preserving the original raw data for comparison.

## 🛠️ Tools Used

* Microsoft Excel
* VBA (Visual Basic for Applications)
* Pivot Tables
* Excel Charts

## 📂 Project Structure

```text
SWYNEX-Data-Clean-Preparation/
│
├── Raw_Data
├── Cleaned_Data
├── Data_Quality
├── Pivot_Table_Month
├── Pivot_Table_Sector
├── Summary
└── README.md
```

## 🧹 Data Cleaning & Preparation

### 1. Missing Values

Used VBA to identify and count missing values across columns and record the results for data-quality review.

### 2. Duplicate Records

Used a VBA-based duplicate detection approach to identify exact duplicate records and remove duplicate rows while maintaining a record of the cleaning process.

### 3. Text Standardization

Standardized inconsistent values in fields such as:

* City
* Gender
* Employment Type
* AI Adoption Level

VBA was used to trim unnecessary spaces and map inconsistent entries to standardized values.

### 4. Date Standardization

Converted inconsistent date/text entries into proper Excel dates and applied a consistent date format.

### 5. Numeric Data Cleaning

Cleaned fields such as:

* Monthly Income (INR)
* Working Hours per Week

Removed symbols, commas, and unnecessary text before converting the values into proper numeric format.

### 6. Outlier Detection

Used logical boundaries to identify unusual values in:

* Age
* Working Hours per Week
* Job Security Score

Outliers were **flagged for review instead of automatically deleted**.

### 7. Impossible Value Detection

Validated fields such as:

* CPI Inflation %
* UPI Transactions per Month

Values outside the defined logical range were flagged for further review.

### 8. Cleaned Dataset

The cleaned and validated records were stored separately from the raw dataset.

This preserves the original data and makes the cleaning process easier to audit and compare.

## 📊 Analysis

### Monthly Analysis

Created a Pivot Table to analyze patterns and aggregated results by month.

### Sector Analysis

Created a Pivot Table to compare data across different sectors.

### Summary

Prepared a summary of the major data-quality findings and cleaned dataset.

## 📈 Key Data Quality Checks

The project covers:

* Missing values
* Duplicate records
* Inconsistent text values
* Incorrect date formats
* Numeric values stored as text
* Outliers
* Impossible values
* Data validation

## 🎯 Project Objective

The main objective was to transform a **messy raw dataset into a structured, standardized, and analysis-ready dataset using Excel and VBA**, while maintaining data integrity and traceability throughout the cleaning process.
