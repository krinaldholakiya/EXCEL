<div align="center">

# -- ! Superstore Sales Intelligence Dashboard ! --
### *Excel-Based Sales & Customer Performance Analysis*

[![Excel](https://img.shields.io/badge/Microsoft%20Excel-2016%2B-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/en/microsoft-365/excel)
[![PivotTables](https://img.shields.io/badge/PivotTables-Data%20Summary-FF6F00?style=for-the-badge&logo=googlesheets&logoColor=white)](https://support.microsoft.com/en-us/office/pivottable)
[![Dashboard](https://img.shields.io/badge/Dashboard-Interactive-4CAF50?style=for-the-badge&logo=speedtest&logoColor=white)](https://www.microsoft.com/en/microsoft-365/excel)
[![Formulas](https://img.shields.io/badge/Formulas-VLOOKUP%20%7C%20IF%20%7C%20SUMIFS-9C27B0?style=for-the-badge&logo=googlesheets&logoColor=white)](https://www.microsoft.com/en/microsoft-365/excel)

<br/>

> *"Data doesn't speak until you give it a chart — turn rows into revenue insight."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🧹 Part A — Data Cleaning & Preparation](#-part-a--data-cleaning--preparation)
- [📊 Part B — Pivot Tables](#-part-b--pivot-tables)
- [🧮 Part C — Advanced Formulas](#-part-c--advanced-formulas)
- [📈 Part D — Interactive Dashboard](#-part-d--interactive-dashboard)
- [🖼️ Project Screenshots](#️-project-screenshots)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Superstore Sales Intelligence Dashboard** is an end-to-end Excel analytics project built on the classic **Superstore Sales Dataset**. It demonstrates core data analyst skills such as **data cleaning**, **PivotTables**, **advanced formulas (VLOOKUP, IF, SUMIFS, nested logic)**, **charting**, and a fully **interactive slicer-driven dashboard**.

This project is designed to:
- Strengthen understanding of Excel data cleaning and preparation workflows
- Practice building PivotTables and PivotCharts from raw transactional data
- Apply advanced formula logic (lookups, conditional classification, aggregation)
- Produce a visually structured, decision-ready sales dashboard

---

## 🎯 Problem Statement

> **Objective:** Analyze multi-year Superstore sales data and build an interactive Excel dashboard for sales & customer performance.

You are working as a data analyst tasked with turning raw, unstructured retail transaction records into actionable business insight. The workbook must clean the raw data, summarize it using PivotTables, layer in advanced formulas for lookups and classification, and present everything through a single interactive dashboard.

| 📂 Feature | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| Raw Data Sheet | Data Source | Cleaned Superstore transactional records |
| Pivot Tables | Summary | Category, Region, Sub-Category, Year & Customer summaries |
| Advanced Formulas | Logic | Lookup, classification & conditional sales analysis |
| Dashboard | Visualization | KPI cards, charts, and slicers in one interactive view |

The goal is to demonstrate **end-to-end Excel data analysis skills** through a clean, professional, business-ready dashboard.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🧹 **Data Cleaning** | Structured, de-duplicated, and formatted raw sales data |
| 📊 **5+ Pivot Tables** | Category, Sub-Category, Region, Order Year, and Customer summaries |
| 🧮 **Advanced Formulas** | VLOOKUP-based customer lookup, IF-based sales classification, SUMIFS regional totals |
| 📈 **KPI Cards** | Total Sales, Total Products, Total Orders, Total Customers, Average Order Value |
| 🥧 **Multiple Chart Types** | Bar, line, column, and pie charts for regional and category insight |
| 🎛️ **Interactive Slicers** | Filter dashboard live by Order Date, Region, Segment, and Category |
| 🖥️ **Single-Page Dashboard** | All KPIs, charts, and filters consolidated into one dashboard view |
| 📄 **Project Documentation** | Dedicated sheet describing objective, tools, and tasks completed |

---

## 🏗️ Project Structure

```
📦 superstore-sales-project/
│
├── 📄 Superstore_Sales_Project_practical.xlsx   ← Main Excel workbook
│   ├── 🧹 Raw Data                              ← Cleaned transactional dataset
│   ├── 📊 Pivot Tables                          ← Category/Region/Year/Customer summaries
│   ├── 🧮 Advanced Formula Analysis             ← Lookup & classification logic
│   ├── 📈 Dashboard                             ← Interactive KPI dashboard
│   └── 📝 Project Documentation                 ← Objective, tools & tasks
│
├── 🖼️ images/                                   ← Dashboard & output screenshots
│
└── 📄 README.md                                 ← Project documentation
```

---

## 🔄 Project Workflow

```
Project Start
      │
      ▼
┌─────────────────────────────┐
│   Dataset Selection & Import │  ← Superstore Sales Dataset
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Data Cleaning & Prep       │  ← Format, de-duplicate, structure columns
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Build Pivot Tables         │  ← Category, Region, Sub-Category, Year, Customer
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Apply Advanced Formulas    │  ← VLOOKUP, IF, SUMIFS
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Data Visualization         │  ← Bar, Line, Column & Pie charts
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│   Build Interactive Dashboard│  ← KPIs + Charts + Slicers
└────────────┬────────────────┘
             │
             ▼
      Documentation ✅
```

---

## 🧹 Part A — Data Cleaning & Preparation

### 📝 1. What Was Done

The raw Superstore dataset was cleaned and standardized before analysis — this included consistent date formats, verified numeric fields for Sales, and structured columns for Ship Mode, Customer, Segment, Location, Product, and Order Year/Month/Shipping Days.

### 🗺️ 2. Key Columns Prepared

| Column Group | Fields | Purpose |
|--------------|--------|---------|
| Customer Info | Customer ID, Customer Name, Segment | Identify and group customers |
| Location | Country, City, State, Postal Code, Region | Enable regional analysis |
| Product Info | Product ID, Category, Sub-Category, Product Name | Enable category analysis |
| Order Info | Order Year, Order Month, Shipping Days | Enable time-based trend analysis |
| Sales | Sales | Core metric for all aggregations |

**Sample of cleaned data:**

Ship Mode, Customer ID/Name, Segment, Location, Product hierarchy, Sales value, and derived Order Year/Month/Shipping Days columns — ready for pivoting and formula-based analysis.

---

## 📊 Part B — Pivot Tables

> Five PivotTables summarize the cleaned dataset from different business angles.

**Pivot Summaries Built:**
```
1. Category            → Sum of Sales
2. Sub-Category        → Sum of Sales
3. Region               → Sum of Sales
4. Order Year          → Sum of Sales
5. Customer Name        → Sum of Sales
```

**Sample Insight (Category):**
```
Furniture         : ₹7,28,658.58
Office Supplies   : ₹7,05,422.33
Technology        : ₹8,27,455.87
Grand Total       : ₹22,61,536.78
```

**Sample Insight (Order Year):**
```
2015 : ₹4,79,856.21
2016 : ₹4,59,436.01
2017 : ₹6,00,192.55
2018 : ₹7,22,052.02
```

---

## 🧮 Part C — Advanced Formulas

### 🔍 3. Formula Logic Used

> Combines lookup functions, conditional classification, and conditional aggregation to power the dashboard's dynamic insights.

**Logic:**
```
Total Sales        = SUM(Sales)
Average Sales       = AVERAGE(Sales)
Maximum / Minimum   = MAX(Sales) / MIN(Sales)
Total Unique Orders = COUNT of distinct Order IDs

Customer Lookup:
  Region  = VLOOKUP(Customer Name, DataTable, RegionCol, FALSE)
  Segment = VLOOKUP(Customer Name, DataTable, SegmentCol, FALSE)

Sales Classification:
  IF(Sales >= 5000, "Very High",
    IF(Sales >= 1000, "High Sales", "Low Sales"))

Region Sales Analysis:
  = SUMIFS(Sales, Region, "West")
```

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 🔎 `VLOOKUP` | Dynamic customer region/segment lookup |
| 🔀 `IF` (nested) | Sales-level & sales-category classification |
| ➕ `SUMIFS` | Region-wise conditional sales totals |
| 📐 `SUM / AVERAGE / MAX / MIN` | Core summary statistics |

**Sample Output:**
```
Total Sales     : ₹22,61,536.78
Average Sales    : ₹230.77
Maximum Sale     : ₹22,638.48
Minimum Sale     : ₹0.44
Customer "Claire Gute" → Region: South | Segment: Consumer
Sales 1000 → High Sales | Sales 10000 → Very High
Region "West" → Total Regional Sales: ₹7,10,219.68
```

---

## 📈 Part D — Interactive Dashboard

### 🎛️ 4. Dashboard Components

| Component | Detail |
|-----------|--------|
| KPI Cards | Total Sales, Total Products, Total Orders, Total Customers, Avg Order Value |
| Regional Sales Performance | Horizontal bar chart by Region |
| Year-wise Sales Trend | Line chart across 2015–2018 |
| Total Sales by Category | Column chart (Furniture, Office Supplies, Technology) |
| Sales Distribution by Category | Pie chart by top customers |
| Slicers | Order Date, Region, Segment, Category — all linked for live filtering |

---

## 🖼️ Project Screenshots

### 1️⃣ Interactive Dashboard Overview
![Dashboard Overview](images/01-dashboard-overview.png)

### 2️⃣ Pivot Table Summaries
![Pivot Tables](images/02-pivot-tables.png)

### 3️⃣ Cleaned Raw Sales Data
![Raw Data](images/03-raw-data.png)

### 4️⃣ Advanced Formula Analysis
![Advanced Formula Analysis](images/04-advanced-formulas.png)

### 5️⃣ Project Documentation Sheet
![Project Documentation](images/05-project-documentation.png)

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📊 **Microsoft Excel** | 2016+ / Microsoft 365 | Core analysis & dashboard platform |
| 📈 **PivotTables / PivotCharts** | Built-in | Data summarization & visualization |
| 🧮 **Formulas** | VLOOKUP, IF, SUMIFS, SUM, AVERAGE, MAX, MIN | Lookup, classification & aggregation |
| 🎛️ **Slicers** | Built-in | Interactive, linked dashboard filtering |
| 🎨 **Charts** | Bar, Line, Column, Pie | Visual storytelling of sales data |

---

## 📈 Results & Insights

After completing the analysis, the following outputs are produced:

- ✅ **Total Sales of ₹22,61,536.78** across 4,922 unique orders and 793 customers
- 🥇 **Technology** is the top-performing category at ₹8,27,455.87 in sales
- 🌍 **West Region** leads all regions with ₹7,10,219.68 in sales
- 📈 **Consistent year-over-year growth**, peaking in 2018 at ₹7,22,052.02
- 🎛️ **Fully interactive dashboard** — filter instantly by date, region, segment & category
- 🧮 **Reusable formula logic** for customer lookup and sales-level classification

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Business-Ready** | Mirrors a real-world retail analytics deliverable |
| 🔄 **Reusability** | Pivot & formula logic can be reapplied to any similar dataset |
| 📚 **Educational** | Demonstrates the full Excel analyst workflow end-to-end |
| 🖥️ **No Add-ins Required** | Built entirely with native Excel features |
| ⚡ **Lightweight** | Single workbook, instantly usable on any machine with Excel |
| 🧪 **Extensible** | Easy to add new KPIs, charts, or additional pivot breakdowns |
| 📖 **Well Documented** | Dedicated documentation sheet inside the workbook |
| 🛡️ **Data Integrity** | Cleaned and validated before any analysis is performed |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Your Name Here

[![GitHub](https://img.shields.io/badge/GitHub-yourhandle-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/yourhandle)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/yourhandle/)

> *"Every dashboard starts with a single clean row — just like every insight starts with a single question."*

**🎓 Role:** Data Analyst | Excel Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel · PivotTables · Data Cleaning · Dashboarding · Business Analysis

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support](https://support.microsoft.com/en-us/excel) — Official Excel documentation
- 📊 [Superstore Sales Dataset](https://www.kaggle.com/datasets) — Source dataset for analysis
- 🎛️ [Excel Slicers Guide](https://support.microsoft.com/en-us/office/use-slicers-to-filter-data) — Interactive filtering reference
- 🧮 [ExcelJet — Formulas](https://exceljet.net/) — Formula logic reference
- 🥧 [Chart.js / Excel Charts Guide](https://support.microsoft.com/en-us/office/available-chart-types-in-office) — Chart type reference
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support

---

<div align="center">

---

*Made with ❤️ and 📊 — Last updated: 17 September, 2026*

</div>
