<div align="center">

# -- ! Excel Formula Mastery Workbook ! --
### *Multi-Sheet Data Analysis using Advanced Excel Formulas & Functions*

[![Excel](https://img.shields.io/badge/Excel-2016%2B-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Formulas](https://img.shields.io/badge/Formulas-VLOOKUP%20%7C%20XLOOKUP%20%7C%20INDEX--MATCH-FF6F00?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)
[![Data](https://img.shields.io/badge/Data-3%20Sheets%20%7C%2060%2B%20Rows-4CAF50?style=for-the-badge&logo=googlesheets&logoColor=white)](https://www.microsoft.com/excel)
[![Logic](https://img.shields.io/badge/Logic-Conditional%20%26%20Lookup-9C27B0?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://www.microsoft.com/excel)

<br/>

> *"A spreadsheet is only as smart as the formulas behind it — master them, and data becomes insight."*

</div>

---

## 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Problem Statement](#-problem-statement)
- [✨ Key Features](#-key-features)
- [🏗️ Project Structure](#️-project-structure)
- [🔄 Project Workflow](#-project-workflow)
- [🎓 Part A — Students Grade Sheet](#-part-a--students-grade-sheet)
- [💰 Part B — Sales Data Sheet](#-part-b--sales-data-sheet)
- [🧑‍💼 Part C — Employee Data Sheet](#-part-c--employee-data-sheet)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Results & Insights](#-results--insights)
- [🏆 Advantages](#-advantages)
- [📄 License](#-license)
- [👤 Author](#-author)
- [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Excel Formula Mastery Workbook** is a hands-on, multi-sheet Microsoft Excel project that demonstrates core spreadsheet skills such as **lookup functions**, **conditional logic**, **text manipulation**, **date arithmetic**, and **aggregate calculations**. The workbook is organized into three independent data sheets, each modeling a real-world business scenario.

This project is designed to:
- Strengthen understanding of `IF`, `AND`, `OR` nested conditional logic
- Practice lookup formulas — `VLOOKUP`, `XLOOKUP`, `INDEX` + `MATCH`
- Apply aggregation formulas — `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`
- Work with text functions, date functions, and dynamic ranges (`INDIRECT`, `OFFSET`)

---

## 🎯 Problem Statement

> **Objective:** Build a multi-sheet Excel workbook that analyzes student performance, sales transactions, and employee records using formula-driven logic — no manual calculation.

You are building a practice workbook for learners studying Excel formulas. Each sheet must accept raw tabular data and compute derived columns automatically — grading, discounts, bonuses, tenure, and search-based lookups — using only native Excel functions.

| 📂 Sheet | 📄 Type | 🔍 Description |
|------------|---------|----------------|
| STUDENTS_GRADE | Academic Data | Computes totals, averages, letter grades, and lookups |
| SALES_DATA | Transactional Data | Computes discounts, net amount, bonus eligibility |
| EMPLOYEE_DATA | HR Data | Computes bonus, tenure, service days, and lookups |

The goal is to demonstrate **practical spreadsheet formula skills** through clean, formula-driven, reusable worksheets.

---

## ✨ Key Features

| Feature | Description |
|--------|-------------|
| 🔁 **3 Independent Sheets** | Students, Sales, and Employee datasets in one workbook |
| 🔎 **Multiple Lookup Styles** | `VLOOKUP`, `XLOOKUP`, and `INDEX` + `MATCH` side by side |
| 🧮 **Conditional Grading** | Nested `IF` ladder converts averages into A–F letter grades |
| 💸 **Dynamic Discount Engine** | Tiered discount % based on sale amount thresholds |
| 📆 **Date-Aware Calculations** | Tenure and days-since-joining computed from `TODAY()` |
| 🧩 **Dynamic Ranges** | `INDIRECT` and `OFFSET` used for flexible range summation |
| ✅ **Error-Safe Lookups** | `IFERROR` wraps lookups to avoid `#N/A` breakage |
| 🔤 **Text Manipulation** | `LEFT`, `FIND`, `UPPER`, `LOWER` used for name parsing |

---

## 🏗️ Project Structure

```
📦 excel-formula-mastery-workbook/
│
├── 📊 pr-1.xlsx              ← Main Excel workbook (entry point)
│   ├── 📄 STUDENTS_GRADE     ← Sheet 1: Academic performance analysis
│   ├── 📄 SALES_DATA         ← Sheet 2: Sales & discount analysis
│   └── 📄 EMPLOYEE_DATA      ← Sheet 3: HR & tenure analysis
│
└── 📄 README.md              ← Project documentation
```

---

## 🔄 Project Workflow

```
Workbook Opened
      │
      ▼
┌─────────────────────────────┐
│   Choose a Data Sheet       │  ← STUDENTS_GRADE / SALES_DATA / EMPLOYEE_DATA
└────────────┬────────────────┘
             │
   ┌─────────┼─────────────┐
   ▼         ▼              ▼
┌────────┐ ┌────────┐ ┌───────────┐
│Students│ │ Sales  │ │ Employee  │
│ Sheet  │ │ Sheet  │ │  Sheet    │
└───┬────┘ └───┬────┘ └─────┬─────┘
    │          │             │
    ▼          ▼             ▼
┌─────────────────────────────┐
│  Formulas Auto-Calculate    │  ← Grades / Discounts / Tenure
│  on every row (drag-fill)   │
└────────────┬────────────────┘
             │
             ▼
┌─────────────────────────────┐
│  Lookup Panel Returns Result│  ← VLOOKUP / XLOOKUP / INDEX-MATCH
└────────────┬────────────────┘
             │
             ▼
     Insights Ready ✅
```

---

## 🎓 Part A — Students Grade Sheet

### 📝 1. What does this sheet do?

The `STUDENTS_GRADE` sheet takes raw marks (Math, Science, English) for 20 students and auto-computes totals, averages, letter grades, eligibility flags, and name-based text transformations.

---

### 🗺️ 2. Formula Map — Overview

| Column | Formula Used | Purpose |
|---------|-------|------------|
| Total | `SUM` | Adds marks across all three subjects |
| Average | `AVERAGE` | Mean score across subjects |
| Grade | Nested `IF` | Converts average into A / B / C / D / F |
| 80+ In Math & Science | `IF` + `AND` | Flags students strong in both subjects |
| First Name | `LEFT` + `FIND` | Extracts first name from full name |
| Name UPPER / LOWER | `UPPER` / `LOWER` | Case conversion of the name |
| Avg (Score > 60) | `AVERAGEIFS` | Average of students scoring above 60 |
| Above 50 in Maths | `COUNTIFS` | Counts students scoring above 50 in Math |
| Student Lookup | `VLOOKUP` + `IFERROR` | Finds a student's name by ID, error-safe |
| Score Lookup | `XLOOKUP` | Finds a student's Math score by ID |

---

### 🔺 3. Grade Classification

> Converts the numeric average into a letter grade using a nested `IF` ladder.

**Logic:**
```
=IF(H2>=90,"A",IF(H2>=80,"B",IF(H2>=70,"C",IF(H2>=60,"D","F"))))
```

**Sample Output:**
```
Average 95.0 → Grade A
Average 84.3 → Grade B
Average 74.6 → Grade C
Average 61.6 → Grade D
Average 58.3 → Grade F
```

---

### 🏔️ 4. Name Parsing

> Splits the full name to extract the first name and produces upper/lower case versions.

**Logic:**
```
=LEFT(B2,FIND(" ",B2)-1)     → First Name
=UPPER(B2)                    → NAME IN CAPS
=LOWER(B2)                    → name in lowercase
```

---

### 🔻 5. Safe Student Lookup

> Looks up a student by ID and returns "Not Found" instead of an error if the ID doesn't exist.

**Logic:**
```
=IFERROR(VLOOKUP(A24,A2:B21,2,FALSE),"Not Found")
=XLOOKUP(E23,A2:A21,C2:C21,"Not Found")
```

---

## 💰 Part B — Sales Data Sheet

### 🔍 6. Discount & Bonus Engine

> Iterates over 20 sales transactions, applies a tiered discount, and flags bonus-eligible salespeople.

**Logic:**
```
=IF(E2>=40000,0.1,IF(E2>=25000,0.05,0))     → Discount %
=E2*G2                                       → Discount Amount
=E2-H2                                       → Net Amount
=IF(AND(E2>30000,C2="North"),"Yes","No")    → High-value North sale
=IF(OR(E2>40000,C2="North"),"Yes","No")     → Bonus Eligible
```

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 🔁 Tiered `IF` | Discount rate changes at 25,000 and 40,000 thresholds |
| ➗ `AND` / `OR` Logic | Combines amount and region conditions |
| ➕ `SUMIFS` | Sums net sales filtered by region and product |
| 🧩 `INDIRECT` + `OFFSET` | Sums a dynamically named or sized range |
| 🖨️ `XLOOKUP` | Retrieves a salesperson's sale amount by name |

**Sample Output (row 5):**
```
Amount: 43,388 → Discount: 10% → Discount Amt: 4,338.80
Net Amount: 39,049.20 | High-Value North Sale: Yes | Bonus Eligible: Yes
```

---

## 🧑‍💼 Part C — Employee Data Sheet

### 🔍 7. Tenure & Bonus Analysis

> Iterates over 20 employee records and computes bonus, years of service, and days since joining using date arithmetic.

**Logic:**
```
=ROUND(D2*0.1,-2)             → Bonus (10% of salary, rounded)
=ROUND((TODAY()-E2)/365,1)    → Years of Service
=TODAY()-E2                   → Days Since Joining
=INDEX(C2:C21,MATCH(N1,A2:A21,0))   → Department lookup by ID
=ABS(VLOOKUP(K2,A2:E21,5,FALSE)-VLOOKUP(K1,A2:E21,5,FALSE))  → Salary gap between two employees
```

**Key Concepts Used:**

| Concept | Detail |
|---------|--------|
| 📆 `TODAY()` | Dynamic current-date reference for tenure math |
| 🔁 `ROUND` | Cleans bonus and tenure figures to readable precision |
| 🔍 `INDEX` + `MATCH` | Two-way lookup independent of column order |
| ➖ `ABS` | Ensures a positive salary difference regardless of order |

**Sample Output (Employee 3005):**
```
Salary: 54,149 | Bonus: 5,400 | Years of Service: 11.7 | Department: IT
```

---

## 🛠️ Tech Stack

| Tool | Version | Purpose |
|------|---------|---------|
| 📊 **Microsoft Excel** | 2016+ | Core spreadsheet application |
| 🔍 **Lookup Functions** | Built-in | `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH` |
| 🧮 **Logical Functions** | Built-in | `IF`, `AND`, `OR`, `IFERROR` |
| ➕ **Aggregate Functions** | Built-in | `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` |
| 🔤 **Text Functions** | Built-in | `LEFT`, `FIND`, `UPPER`, `LOWER` |
| 📆 **Date Functions** | Built-in | `TODAY`, date subtraction |
| 🧩 **Dynamic Range Functions** | Built-in | `INDIRECT`, `OFFSET` |

---

## 📈 Results & Insights

After opening the workbook, the following outputs are produced:

- ✅ **3 Fully Formula-Driven Sheets** — Students, Sales, and Employee data
- 🎓 **Automatic Grading** — Every student is classified from A to F instantly
- 💸 **Live Discount Calculation** — Net sale amount recalculates the moment raw amount changes
- 📆 **Real-Time Tenure Tracking** — Years of service updates automatically with `TODAY()`
- 🔍 **Error-Safe Lookups** — ID/name-based searches never break the sheet with `#N/A`

---

## 🏆 Advantages

| Advantage | Detail |
|-----------|--------|
| 🎓 **Beginner to Intermediate Friendly** | Covers logic, lookup, and aggregate formulas in one file |
| 🔄 **Reusability** | Formula patterns can be copied into any similar dataset |
| 📚 **Educational** | Each sheet reinforces a different formula family |
| 🖥️ **No Add-ins Needed** | Works with native Excel — no macros or plugins |
| ⚡ **Lightweight** | Single workbook, instantly usable in Excel or Google Sheets |
| 🧪 **Extensible** | Easy to add new sheets (Inventory, Attendance, etc.) |
| 📖 **Readable Formulas** | Clear cell references make logic easy to trace |
| 🛡️ **Input Safety** | `IFERROR` wraps every lookup against bad input |

---

## 📄 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for full details.

```
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

<div align="center">

### KRINAL DHOLAKIYA


[![GitHub](https://img.shields.io/badge/GitHub-krinaldholakiya-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/krinaldholakiya)

> *"Every insight starts with a single formula — just like every workbook starts with a single cell."*

**🎓 Role:** Excel Learner | Data Enthusiast \
**📍 Location:** India\
**🛠️ Skills:** Excel · Formulas · Lookup Functions · Data Analysis · Logic Building

</div>

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

- 📚 [Microsoft Excel Support Docs](https://support.microsoft.com/en-us/excel) — Official Excel function reference
- 🔍 [ExcelJet](https://exceljet.net/) — In-depth formula tutorials
- 📐 [GeeksForGeeks — Excel Formulas](https://www.geeksforgeeks.org/) — Formula usage examples
- 🖥️ [W3Schools](https://www.w3schools.com/) — Beginner spreadsheet reference
- 🧮 [Chandoo.org](https://chandoo.org/) — Advanced Excel techniques
- 💬 [Stack Overflow Community](https://stackoverflow.com/) — Problem-solving support
- 📖 [Kaggle Learn](https://www.kaggle.com/learn) — Data analysis courses

---

<div align="center">

---

*Made with ❤️ — Last updated: 29 August, 2026*

</div>
