# 📊 PR. 2 — ANALYZER
### 🚀 Advanced Excel Analytics | Customer Insights | Interactive Dashboard

<div align="center">

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Data Analytics](https://img.shields.io/badge/Data_Analytics-0078D4?style=for-the-badge&logo=googleanalytics&logoColor=white)
![Dashboard](https://img.shields.io/badge/Dashboard-Interactive-8A2BE2?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed-00B894?style=for-the-badge)

**✨ FROM RAW DATA TO POWERFUL BUSINESS INSIGHTS ✨**

*Analyze • Visualize • Discover • Decide*

---

### 💎 Transforming 200 Sales Transactions into Meaningful Insights!

</div>

## 🌟 Project Overview

**PR. 2 — Analyzer** is an advanced Microsoft Excel practical project focused on transforming sales transaction data into meaningful business insights using analytical functions, statistical techniques, conditional formatting, and data visualization.

The project explores customer purchasing behavior, sales performance, profitability, discount scenarios, regional trends, and product category performance.

> 💡 **"Data tells a story. Analytics helps us understand it."**

---

## 🎯 Project Objectives

- 📊 Analyze customer sales and profitability.
- 🏆 Identify the top 10 highest-value customers.
- 🎨 Highlight important values using Conditional Formatting.
- 🔮 Explore discount scenarios using What-If Analysis.
- 📈 Examine the relationship between Sales and Profit using Regression.
- 🧮 Calculate Descriptive Statistics.
- 📅 Track monthly sales growth and trends.
- 🔎 Retrieve customer information using INDEX and MATCH.
- 🌍 Compare regional and product category performance.
- 📉 Visualize business KPIs using charts and a dashboard.

---

## 📁 Project Dataset

The dataset contains **200 sales transactions** with the following fields:

| Column | Field | Description |
|---|---|---|
| A | Customer_ID | Unique customer identifier |
| B | Customer_Name | Customer name |
| C | Region | Geographic sales region |
| D | Product_Category | Product classification |
| E | Sales | Transaction sales amount |
| F | Quantity | Units sold |
| G | Discount | Discount percentage |
| H | Order_Date | Transaction date |
| I | Profit | Profit earned |

**📌 Dataset Range:** `Dataset!A1:I201`

---

## 📊 Business Performance Snapshot

<div align="center">

| 📈 KPI | 💎 Value |
|---|---:|
| 🧾 Total Transactions | **200** |
| 💰 Total Sales | **195,217.76** |
| 💵 Total Profit | **68,287.01** |
| 📊 Overall Profit Margin | **34.98%** |
| 📉 Average Sales per Transaction | **976.09** |
| 🎯 Average Profit per Transaction | **341.44** |

</div>

> These values are calculated from the supplied dataset. Currency is not specified in the source data.

---

## 🧠 Advanced Excel Techniques

### 🎨 01. Conditional Formatting

Conditional Formatting makes important values stand out automatically.

**Applications:**
- Data Bars to visualize sales amounts.
- Color-based highlighting for important values.
- Top 10 customer identification.
- Performance comparisons.

**Example:**
```excel
=RANK.EQ(E2,$E$2:$E$201,0)<=10
```

---

### 🔮 02. What-If Analysis

What-If Analysis explores how changing discount assumptions may affect business profitability.

**Concept:**

`Discount Change → Profit Adjustment → Scenario Comparison`

**Example model:**
```excel
=I2-E2*(K2-G2)
```

Here, `K2` represents a proposed discount rate, while `G2` is the original discount rate.

*This is an illustrative scenario model that assumes the change in discount directly changes profit without changing demand or other costs.*

---

### 📈 03. Regression Analysis

Regression helps investigate the relationship between two numerical variables.

**Independent Variable (X):** Sales

**Dependent Variable (Y):** Profit

**Excel Functions:**
```excel
=SLOPE(I2:I201,E2:E201)
=INTERCEPT(I2:I201,E2:E201)
=RSQ(I2:I201,E2:E201)
```

**Regression Equation:**

`Predicted Profit = Intercept + (Slope × Sales)`

This analysis helps understand how profit tends to change as sales increase.

---

### 🧮 04. Descriptive Statistics

Descriptive Statistics summarizes the distribution and characteristics of numerical data.

| Measure | Sales | Profit |
|---|---:|---:|
| Count | 200 | 200 |
| Sum | 195,217.76 | 68,287.01 |
| Mean | 976.09 | 341.44 |
| Median | 976.72 | 332.26 |
| Minimum | 258.51 | 64.63 |
| Maximum | 1,923.66 | 769.46 |

**Functions Used:**

`COUNT()` · `SUM()` · `AVERAGE()` · `MEDIAN()` · `MIN()` · `MAX()` · `STDEV.S()`

---

### 📅 05. Monthly Sales Analysis

Monthly analysis helps track changes in sales performance over time.

**Month Extraction:**
```excel
=TEXT(H2,"mmm-yyyy")
```

**Month-over-Month Growth:**
```excel
=IFERROR((B3-B2)/B2,0)
```

**Growth Indicators:**

⬆️ Positive Growth — Sales Increased

⬇️ Negative Growth — Sales Decreased

➡️ Stable Performance — No Change

---

### ⏰ 06. Timestamp Generation

Excel can automatically display the current date and time.

```excel
=NOW()
```

**Note:** `NOW()` is dynamic and updates when Excel recalculates.

---

### 🏆 07. High-Value Customer Analysis

Customer ranking helps identify valuable customers based on total purchase amounts.

**Functions & Techniques:**

`SUMIF()` · `RANK.EQ()` · `INDEX()` · `MATCH()` · `FILTER()`

**Example:**
```excel
=SUMIF(Dataset!$A$2:$A$201,A2,Dataset!$E$2:$E$201)
```

This formula calculates the total sales for a particular Customer ID.

**Business Benefits:**
- Identify loyal customers.
- Understand purchasing behavior.
- Recognize high-revenue customer groups.
- Support customer retention strategies.

---

### 🌍 08. Region & Product Category Analysis

Sales and profitability are compared across different regions and product categories.

**Regions:**

North • South • East • West • Central

**Product Categories:**

Electronics • Furniture • Clothing • Books • Office Supplies

**Example:**
```excel
=SUMIFS(Dataset!$E$2:$E$201,Dataset!$C$2:$C$201,A2,Dataset!$D$2:$D$201,B2)
```

This calculates sales for a selected region and product category.

---

## 📉 Data Visualization

The project uses visual representations to make complex information easier to understand.

| Chart Type | Purpose |
|---|---|
| 📊 Bar Chart | Compare sales and profit across categories |
| 📈 Line Chart | Analyze monthly sales trends |
| 🥧 Pie Chart | Visualize regional or category contributions |
| 📉 Histogram | Understand sales distribution |
| 📋 KPI Dashboard | Summarize business performance |

---

## 💻 Dashboard Highlights

The Executive Dashboard brings important metrics together in one organized view.

**Key Dashboard Elements:**

- 💰 Total Sales
- 💵 Total Profit
- 📊 Profit Margin
- 🧾 Transaction Count
- 🏆 Customer Insights
- 🌍 Regional Performance
- 📈 Monthly Sales Trends
- 🛍️ Product Category Comparisons

> **Goal:** Make business performance easier to monitor, interpret, and communicate.

---

## 🛠️ Tools & Technologies

| Tool / Feature | Application |
|---|---|
| Microsoft Excel | Data analysis and spreadsheet management |
| Excel Formulas | Automated calculations |
| Conditional Formatting | Visual highlighting and Data Bars |
| What-If Analysis | Discount scenario modeling |
| Statistical Functions | Descriptive and regression calculations |
| INDEX & MATCH | Dynamic customer lookup |
| SUMIF & SUMIFS | Conditional aggregation |
| Charts | Data visualization |
| Dashboard | Business performance reporting |

---

## 🚀 How to Use

**Step 1:** Download the completed Excel workbook.

**Step 2:** Open the `.xlsx` file using Microsoft Excel.

**Step 3:** Explore the `Dataset` worksheet to understand the original records.

**Step 4:** Review the descriptive statistics and regression results.

**Step 5:** Change the proposed discount in the What-If Analysis worksheet to explore alternative scenarios.

**Step 6:** Examine customer rankings, monthly trends, and region/product summaries.

**Step 7:** Open the Executive Dashboard to view the visual business summary.


---

## 🎓 Learning Outcomes

Through this project, I gained practical experience in:

- Performing advanced spreadsheet analysis.
- Applying Conditional Formatting to business datasets.
- Understanding What-If Analysis and regression concepts.
- Calculating descriptive statistics.
- Ranking and analyzing high-value customers.
- Creating region-wise and category-wise performance summaries.
- Interpreting charts and business KPIs.
- Presenting insights through a structured dashboard.

---

## 📌 Implementation Notes

- The workbook uses the original 200 transaction records.
- What-If Analysis uses an assumption-based profit model.
- Regression is calculated using Excel statistical formulas.
- Regional and product summaries are formula-based.
- A formula-based summary is not the same as a native refreshable Excel PivotTable.
- The statistical calculations should not be confused with a separately generated Data Analysis ToolPak output.

These distinctions help make the project transparent and technically accurate.

---

## 👩‍💻 Developed By

**Minaxi Choudhari**

📚 **Practical:** PR. 2 — Analyzer

💻 **Technology:** Microsoft Excel

---

<div align="center">

## 💖 THANK YOU FOR VISITING!

### ✨ Every Number Has a Story — Let Data Reveal It! ✨

**⭐ If you found this project useful, consider starring the repository!**

**Made with , creativity, dedication & analytical thinking.**

</div>
# pr2excel
