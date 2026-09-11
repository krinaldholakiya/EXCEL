<div align="center">

# 📊 -- ! Data Intelligence Dashboard ! --
### *Excel-Based Sales Analytics, What-If Modeling & Interactive BI Dashboard*

[![Excel](https://img.shields.io/badge/Excel-365%2F2021-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![PivotTables](https://img.shields.io/badge/PivotTables-Dynamic%20Analysis-FF6F00?style=for-the-badge&logo=googlesheets&logoColor=white)](https://www.microsoft.com/excel)
[![Dashboard](https://img.shields.io/badge/Dashboard-Interactive%20BI-4CAF50?style=for-the-badge&logo=powerbi&logoColor=white)](https://www.microsoft.com/excel)
[![WhatIf](https://img.shields.io/badge/What--If-Scenario%20Analysis-9C27B0?style=for-the-badge&logo=googleanalytics&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"Raw data whispers, but a well-built dashboard makes it sing."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [📑 Sheet-by-Sheet Breakdown](#-sheet-by-sheet-breakdown)
- [🖼️ Screenshots](#️-screenshots)
- [🛠️ Tech Stack & Excel Features Used](#️-tech-stack--excel-features-used)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [🚀 How to Use](#-how-to-use)
- [📄 License](#-license)
- [👤 Author](#-author)

---

## 📌 Overview

The **Data Intelligence Dashboard** is a complete, multi-sheet Excel analytics project built on a 250-row transactional sales dataset. It transforms raw, flat sales records into a fully interactive **Business Intelligence dashboard** using **Excel Tables, PivotTables, PivotCharts, Slicers, formulas (SUMIF, XLOOKUP, INDEX/MATCH, array formulas), Scenario Manager, and What-If Analysis**.

This project is designed to:
- Convert raw transactional data into structured, analyzable Excel Tables
- Build KPI-driven analysis sheets using dynamic formulas
- Segment and score customers by value and purchase behavior
- Model "what-if" business scenarios (Base / High / Low sales cases)
- Visualize results through PivotCharts and a single-page executive dashboard

---

## 🎯 Problem Statement

> **Objective:** Turn a raw sales transaction log into an end-to-end, decision-ready Business Intelligence workbook.

A business collects raw transactional data — customers, products, regions, payment methods, and order values — but this data alone doesn't answer strategic questions. This project builds the analytical layer on top of that raw data to answer:

| 📂 Question | 📄 Sheet Used | 🔍 How |
|------------|---------|----------------|
| What are our core KPIs? | Analysis | SUM, AVERAGE, MAX, MIN on the Sales Table |
| Who are our most valuable customers? | Customer Analysis | XLOOKUP, SUMIF, COUNTIF, INDEX/MATCH |
| What happens if sales rise or fall? | Scenario Summary | Excel Scenario Manager |
| What if price or quantity changes? | What-If & Regression | Manual What-If formulas |
| How does revenue break down visually? | Pivot Analysis / Visualizations | PivotTables + PivotCharts |
| What's the one-page executive view? | Dashboard | Slicers + linked charts + KPI cards |

The goal is to demonstrate **end-to-end data analytics skills in Excel** — from raw data to a boardroom-ready dashboard.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 📥 **Structured Sales Table** | 250 transactions converted into a native Excel Table (`SalesData`) for dynamic formula references |
| 🧮 **KPI Engine** | Auto-calculated Total Revenue, Transactions, Quantity Sold, Avg/Max/Min Order Value |
| 🛍️ **Product Performance Ranking** | Each product tagged ▲ High / ► Medium / ▼ Low based on revenue thresholds |
| 👥 **Customer Value Segmentation** | Every customer classified as High Value / Medium Value / Regular using `XLOOKUP` + `SUMIF` |
| 🏆 **Top Customer Detection** | `INDEX` + `MATCH` automatically surfaces the highest-revenue customer |
| 🔮 **Scenario Manager Report** | Base Case, High Sales, and Low Sales scenarios modeled and compared side-by-side |
| ➗ **What-If Revenue Calculator** | Editable Quantity input instantly recalculates Expected Revenue |
| 📊 **PivotTables + PivotCharts** | Revenue by Region, Category, Product, and Month — fully interactive |
| 🖥️ **Executive Dashboard** | Single-page dashboard with KPI cards, charts, and Region/Category/Product/Month slicers |
| 📝 **Auto-Generated Final Report** | Executive summary sheet that recalculates live from the source data |

---

## 🏗️ Project Structure

```
📦 data-intelligence-dashboard/
│
├── 📊 FINAL_PROJECT.xlsx        ← Main Excel workbook (9 sheets)
│   ├── Raw Data                 ← 250-row transactional dataset
│   ├── Analysis                 ← KPIs + product performance formulas
│   ├── Customer Analysis        ← Customer segmentation & top customer
│   ├── Scenario Summary         ← Excel Scenario Manager report
│   ├── What-If & Regression     ← Editable what-if revenue model
│   ├── Pivot Analysis           ← Region / Product / Category / Year pivots
│   ├── Visualizations           ← Supporting PivotCharts
│   ├── Dashboard                ← Interactive executive dashboard
│   └── Final Report             ← Auto-generated executive summary
│
├── 🖼️ images/
│   ├── raw-data.png             ← Raw Data sheet preview
│   ├── visualizations.png       ← PivotChart visualizations
│   └── dashboard.png            ← Final interactive dashboard
│
└── 📄 README.md                 ← Project documentation
```

---

## 🔄 Project Workflow

```
Raw Transactional Data (250 rows)
              │
              ▼
   ┌─────────────────────────┐
   │   Convert to Excel Table │  ← "SalesData" (structured references)
   └────────────┬─────────────┘
                │
   ┌────────────┼─────────────────────┬───────────────────────┐
   ▼            ▼                     ▼                        ▼
┌─────────┐ ┌───────────────┐  ┌──────────────────┐   ┌──────────────────┐
│Analysis │ │Customer        │  │Scenario Summary /  │   │Pivot Analysis /   │
│(KPIs +  │ │Analysis        │  │What-If & Regression│   │Visualizations     │
│Products)│ │(Segmentation)  │  │(Forecast modeling)  │   │(PivotCharts)       │
└────┬────┘ └───────┬────────┘  └──────────┬──────────┘   └──────────┬─────────┘
     │              │                       │                          │
     └──────────────┴───────────┬───────────┴──────────────────────────┘
                                 ▼
                     ┌───────────────────────┐
                     │  Interactive Dashboard  │  ← Slicers + KPI cards + Charts
                     └───────────┬─────────────┘
                                 ▼
                     ┌───────────────────────┐
                     │     Final Report        │  ← Executive summary
                     └───────────────────────┘
```

---

## 📑 Sheet-by-Sheet Breakdown

### 1️⃣ Raw Data
The foundation of the entire workbook — **250 transactions** across 19 columns including `Transaction_ID`, `Customer_ID`, `Product_Name`, `Category`, `Quantity`, `Unit_Price`, `Payment_Method`, `Region`, `Customer_Segment`, `Total_Amount`, and derived date fields. Stored as a native Excel Table named **`SalesData`** so every downstream formula updates automatically when data changes.

### 2️⃣ Analysis
Core KPI engine built with structured-reference formulas:
```excel
Total Revenue            =SUM(SalesData[Total_Amount])
Total Transactions       =COUNTA(SalesData[Transaction_ID])
Total Quantity Sold      =SUM(SalesData[Quantity])
Average Transaction      =AVERAGE(SalesData[Total_Amount])
```
Also includes a **Product Analysis** table where each product's revenue is calculated with `SUMIF` and tagged automatically:
```excel
=IF(I6>=10000,"▲ High",IF(I6>=5000,"► Medium","▼ Low"))
```

### 3️⃣ Customer Analysis
Segments all customers by revenue contribution using `XLOOKUP`, `COUNTIF`, and `SUMIF`:
```excel
Customer Name    =XLOOKUP(A4,SalesData[Customer_ID],SalesData[Customer_Name],"")
Total Revenue    =SUMIF(SalesData[Customer_ID],A4,SalesData[Total_Amount])
Status           =IF(D4>=5000,"High Value",IF(D4>=2500,"Medium Value","Regular"))
```
The **Top Customer** is pulled out automatically using `INDEX` + `MATCH` against the highest revenue value.

### 4️⃣ Scenario Summary
Generated with Excel's built-in **Scenario Manager**, comparing three business cases side-by-side:

| Scenario | Quantity (Changing Cell) | Resulting Revenue |
|----------|---------------------------|--------------------|
| Base Case | 100 | 29,667 |
| High Sales | 150 | 44,500.5 |
| Low Sales | 70 | 20,766.9 |

### 5️⃣ What-If & Regression
A live, editable revenue model:
```excel
Average Unit Price   =AVERAGE(SalesData[Unit_Price])
Quantity             100  (user-editable)
Expected Revenue     =B3*B4
```

### 6️⃣ Pivot Analysis
Multiple **PivotTables** slicing the same dataset from different angles: Revenue by Region, Sum of Quantity/Revenue by Product, Revenue by Category, and Revenue by Year — all refreshable with one click.

### 7️⃣ Visualizations
A set of **PivotCharts** (3D column, line, bar, and pie) built directly on top of the Pivot Analysis tables, giving a visual read on regional, monthly, product, and category performance.

### 8️⃣ Dashboard
The centerpiece: a single-page **interactive dashboard** combining KPI cards (Total Revenue, Total Transactions, Total Quantity, Avg Order Value), four linked charts, and **Slicers** for Date, Region, Category, Product, and Month — letting anyone filter the entire dashboard without touching a formula.

### 9️⃣ Final Report
An auto-generated **executive summary** that recalculates Total Revenue, Transactions, Quantity, and Average Order Value directly from `SalesData`, plus a written revenue insight — ready to paste into a report or email.

---

## 🖼️ Screenshots

### 📥 Raw Data — Structured Sales Table
> The 250-row `SalesData` table, the single source of truth every other sheet references.

![Raw Data Table](images/raw-data.png)

---

### 📊 Visualizations — PivotCharts
> Revenue by Region, Monthly Revenue Trend, Revenue by Product, and Revenue by Category — all built on live PivotTables.

![Visualizations](images/visualizations.png)

---

### 🖥️ Interactive Dashboard
> The final one-page executive dashboard with KPI cards, charts, and Region / Category / Product / Month slicers for instant filtering.

![Dashboard](images/dashboard.png)

---

## 🛠️ Tech Stack & Excel Features Used

| Tool / Feature | Purpose |
|------|---------|
| 📊 **Excel Tables** | Structured references (`SalesData[...]`) so formulas auto-expand with new data |
| 🧮 **SUM / SUMIF / COUNTIF / COUNTA** | KPI and segment calculations |
| 🔎 **XLOOKUP** | Dynamic customer name lookups |
| 🎯 **INDEX + MATCH** | Top-customer / top-value detection |
| 🔀 **IF / Nested IF** | Performance tagging (High/Medium/Low, Value segmentation) |
| 🔮 **Scenario Manager** | Base / High / Low sales scenario modeling |
| ➗ **What-If Analysis** | Editable revenue projection model |
| 📈 **PivotTables & PivotCharts** | Multi-angle revenue analysis and visualization |
| 🎛️ **Slicers** | Interactive, click-to-filter dashboard controls |
| 🖨️ **Dynamic Formulas (TODAY/NOW)** | Live date & time tracking on the Analysis sheet |

---

## 📈 Results & Insights

- ✅ **Total Revenue, Transactions & Quantity** auto-calculated and always in sync with raw data
- 👥 **Customer segmentation** instantly flags High Value vs. Regular customers
- 🛍️ **Product performance tagging** highlights which SKUs are driving revenue
- 🔮 **Scenario comparison** shows revenue could swing from ~20.7K (Low) to ~44.5K (High) against a 29.7K base case
- 🖥️ **One filter, four views** — Dashboard slicers update KPI cards and all charts simultaneously
- 📝 **Executive Final Report** ready to share without opening any other sheet

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **End-to-End Analytics** | Covers raw data → formulas → pivots → dashboard → report in one file |
| 🔄 **Fully Dynamic** | Structured Tables mean new rows flow automatically into every KPI and chart |
| 📚 **Educational** | Demonstrates real-world use of XLOOKUP, INDEX/MATCH, Scenario Manager & Slicers |
| 🖥️ **No Add-ins Required** | Built entirely with native Excel features |
| ⚡ **Decision-Ready** | Dashboard and Final Report are shareable as-is with stakeholders |
| 🧪 **Extensible** | Easy to add more scenarios, KPIs, or pivot breakdowns |
| 🎛️ **Interactive** | Slicers let non-technical users explore data without writing formulas |
| 🛡️ **Single Source of Truth** | Every sheet references the same `SalesData` table, avoiding inconsistent numbers |

---

## 🚀 How to Use

1. Open **`FINAL_PROJECT.xlsx`** in Microsoft Excel (2016+ / 365 recommended for `XLOOKUP` support).
2. Start on the **Raw Data** sheet to review the source `SalesData` table.
3. Explore **Analysis** and **Customer Analysis** for KPI and segmentation formulas.
4. Open **Scenario Summary** (or re-run via *Data → What-If Analysis → Scenario Manager*) to compare Base/High/Low cases.
5. Edit the **Quantity** cell on **What-If & Regression** to see Expected Revenue update live.
6. Go to the **Dashboard** sheet and click any Region, Category, Product, or Month slicer to filter everything at once.
7. Check **Final Report** for a ready-to-share executive summary.

---

## 📄 License

This project is shared for **educational and portfolio purposes**. Feel free to use, adapt, and build on it with attribution.

```
Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### Krinal

> *"Numbers tell the story — a good dashboard just makes it easy to read."*

**🎓 Role:** Data Analyst | Excel & BI Enthusiast \
**📍 Location:** India \
**🛠️ Skills:** Excel · PivotTables · Dashboarding · What-If Analysis · Data Storytelling

</div>

---

<div align="center">

---

*Made with 📊 and ❤️ — Last updated: 11 September, 2026*

</div>
