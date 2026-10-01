# 💰 Loan Portfolio Analysis – Excel Dashboard

An end-to-end **Microsoft Excel analytics project** built to analyze loan portfolio data and transform raw financial information into meaningful business insights.

The project uses **Excel formulas, PivotTables, charts, slicers, timelines, and KPI calculations** to analyze loan amounts, repayment behavior, revolving balances, verification status, loan status, customer characteristics, and state-level performance.

---

## 🚀 Project Overview

Financial and loan datasets can contain thousands of records and dozens of attributes, making it difficult to identify useful trends directly from raw data.

This project transforms raw loan data into an interactive Excel-based analytical solution that helps answer questions such as:

* How has the total loan amount changed over time?
* Which loan grades and sub-grades have higher revolving balances?
* How does payment behavior differ between verified and non-verified customers?
* Which states have the highest number of loans?
* How does loan status vary by state?
* How does home ownership relate to last payment activity?
* What are the major characteristics of the loan portfolio?

The final **Dashboard** brings these analyses together into a single reporting view.

---

## 🎯 Business Objectives

The main objectives of this project are:

* Analyze the overall loan portfolio
* Track loan amount trends by year
* Analyze revolving balance by grade and sub-grade
* Compare total payments across verification statuses
* Analyze loan status geographically
* Understand payment behavior by home ownership
* Build reusable KPI calculations
* Create an interactive Excel dashboard for business analysis

---

## 📂 Dataset Overview

The workbook contains multiple data and analysis sheets.

### Main Data Sheets

#### `Finance`

Contains financial and repayment-related attributes such as:

* `id`
* `delinq_2yrs`
* `earliest_cr_line`
* `inq_last_6mths`
* `mths_since_last_delinq`
* `mths_since_last_record`
* `open_acc`
* `pub_rec`
* `revol_bal`
* `revol_util`
* `total_acc`
* `initial_list_status`
* `out_prncp`
* `out_prncp_inv`
* `total_pymnt`
* `total_pymnt_inv`
* `total_rec_prncp`
* `total_rec_int`
* `total_rec_late_fee`
* `recoveries`
* `collection_recovery_fee`
* `last_pymnt_d`
* `last_pymnt_amnt`
* `next_pymnt_d`
* `last_credit_pull_d`

#### `Combined Data`

Contains the combined loan-level dataset with borrower, loan, financial, and geographic attributes.

Important fields include:

* `id`
* `member_id`
* `loan_amnt`
* `funded_amnt`
* `funded_amnt_inv`
* `term`
* `int_rate`
* `installment`
* `grade`
* `sub_grade`
* `emp_title`
* `emp_length`
* `home_ownership`
* `annual_inc`
* `verification_status`
* `issue_d`
* `loan_status`
* `purpose`
* `zip_code`
* `addr_state`
* `dti`
* `delinq_2yrs`
* `inq_last_6mths`
* `revol_bal`
* and additional borrower/loan attributes

The `Finance` sheet contains approximately **39.7K loan-level records**, while the `Combined Data` sheet contains approximately **39.7K combined records**.

---

# 🧱 Workbook Structure

The workbook is organized into the following major sections:

| Sheet            | Purpose                                    |
| ---------------- | ------------------------------------------ |
| `Finance`        | Core financial and repayment data          |
| `Combined Data`  | Combined borrower and loan dataset         |
| `Data for kpi 3` | KPI calculation support                    |
| `Kpi -1`         | Year-wise loan amount analysis             |
| `Kpi - 2`        | Grade/sub-grade revolving balance analysis |
| `Kpi-3`          | Verification status payment analysis       |
| `Kpi - 4`        | State-wise loan status analysis            |
| `kpi 5`          | Home ownership and last payment analysis   |
| `DashBoard`      | Final interactive dashboard                |

---

# 📊 Dashboard Analysis

## 1️⃣ Year-wise Loan Amount

### Purpose

Analyze how the total loan portfolio has changed across different years.

### Key Analysis

* Year-wise total loan amount
* Historical loan growth
* Loan amount distribution over time

### Excel Techniques Used

* PivotTables
* Aggregation
* Year-based grouping
* Charts

### Business Value

This analysis provides visibility into how the loan portfolio has evolved over time and helps identify periods of higher or lower lending activity.

---

## 2️⃣ Grade & Sub-Grade Wise Revolving Balance

### Purpose

Analyze revolving balances across borrower credit grades and sub-grades.

### Key Analysis

* Revolving balance by sub-grade
* Revolving balance by grade
* Comparison across credit segments

### Example Segments

```text
A1
A2
A3
...
B1
B2
B3
...
G5
```

### Excel Techniques Used

* PivotTables
* Grouping
* Column/bar charts
* Credit-grade segmentation

### Business Value

This analysis helps understand how revolving credit balances are distributed across different borrower credit segments.

---

## 3️⃣ Verification Status vs Total Payment

### Purpose

Compare payment amounts for verified and non-verified borrowers.

### Key Analysis

* Total payment for verified customers
* Total payment for non-verified customers
* Verification status comparison

### Excel Techniques Used

* `SUMIF`
* PivotTables
* KPI calculations
* Charts

### Example Formula

```excel
=SUMIF('Combined Data'!$O:$O,F11,'Combined Data'!$AL:$AL)
```

### Business Value

This analysis helps identify differences in payment values across verification categories.

---

## 4️⃣ State-wise Loan Status

### Purpose

Analyze the distribution of loan statuses across U.S. states.

### Loan Status Categories

* Charged Off
* Current
* Fully Paid

### Key Analysis

* Loan count by state
* Loan status by state
* State-level portfolio distribution

### Excel Techniques Used

* PivotTables
* Slicers
* Timeline filtering
* State-wise charts

### Business Value

This view makes it easier to identify regional differences in loan activity and repayment status.

---

## 5️⃣ Home Ownership vs Last Payment Date

### Purpose

Analyze last payment activity across different home ownership categories.

### Home Ownership Categories

* RENT
* MORTGAGE
* OWN
* OTHER
* NONE

### Key Analysis

* Last payment date
* Last payment amount
* Maximum last payment date by ownership category
* Payment statistics by home ownership

### Excel Techniques Used

* `MAX`
* `SUMIFS`
* Array-style calculations
* PivotTables
* Date analysis

### Example Formula

```excel
=MAX(IF($A$2:$A$39718=A2,$C$2:$C$39718))
```

Another calculation used in the analysis:

```excel
=SUMIFS($B$2:$B$39718,$C$2:$C$39718,D2,$A$2:$A$39718,A2)
```

### Business Value

This analysis provides additional borrower-level context around payment behavior.

---

# 📊 KPI Analysis

The project includes several KPI calculations designed to summarize the loan portfolio.

### Key Metrics

* Total Loan Amount
* Total Payment
* Total Recoveries
* Last Payment Amount
* Verified Customer Payment
* Non-Verified Customer Payment
* Loan Amount by Year
* Revolving Balance by Grade
* Revolving Balance by Sub-Grade
* Loan Count by State
* Loan Status Distribution
* Last Payment Statistics

---

# 🎛️ Interactive Excel Features

The workbook uses interactive Excel functionality to improve analysis.

### Slicers

Users can filter analysis based on fields such as:

* Grade
* Verification Status
* State
* Loan Status

### Timelines

Date-based analysis can be filtered using Excel timeline controls.

### PivotTables

PivotTables are used extensively to summarize:

* Loan amounts
* Loan status
* Revolving balances
* Total payments
* Geographic distributions

### Charts

The workbook converts PivotTable results into visualizations for easier interpretation.

---

# 🛠 Tools & Excel Features Used

* **Microsoft Excel**
* PivotTables
* PivotCharts
* Slicers
* Timelines
* Excel formulas
* `SUMIF`
* `SUMIFS`
* `MAX`
* `IF`
* Data aggregation
* Date analysis
* KPI calculations
* Dashboard design

---

# 📈 Business Questions Answered

This project addresses several practical analytical questions:

### Loan Portfolio

* How much money has been issued through loans?
* How has loan volume changed over time?
* Which years recorded higher loan amounts?

### Credit Analysis

* Which grades have higher revolving balances?
* How does revolving balance differ by sub-grade?

### Payment Analysis

* How much has been paid by verified customers?
* How much has been paid by non-verified customers?
* How does total payment vary across loan statuses?

### Geographic Analysis

* Which states have the highest loan activity?
* How does loan status vary by state?

### Borrower Analysis

* How does home ownership relate to payment behavior?
* What are the latest payment patterns across ownership groups?

---

# 📌 Key Excel Skills Demonstrated

This project demonstrates practical experience with:

* Large dataset handling
* Data preparation
* Data aggregation
* PivotTable development
* PivotChart creation
* KPI design
* Conditional analysis
* Date-based analysis
* Slicers and timelines
* Formula-based calculations
* Dashboard development
* Business-oriented data storytelling

---

# 💡 Business Value

The Excel dashboard provides a consolidated view of the loan portfolio and allows users to:

* Monitor portfolio performance
* Analyze historical loan trends
* Compare borrower credit segments
* Understand payment behavior
* Identify state-level differences
* Explore borrower characteristics
* Interactively filter and analyze the data

The project demonstrates how **raw transactional data can be converted into a structured analytical dashboard using Excel**.

---

# 📚 Key Learnings

Through this project, I gained practical experience in:

* Working with large Excel datasets
* Combining data from multiple sources/tables
* Creating KPI calculation sheets
* Building PivotTables and PivotCharts
* Using advanced Excel formulas
* Working with dates and financial metrics
* Creating interactive dashboards
* Converting business questions into analytical views
* Presenting data in a clear and structured format

---

# 🚀 Future Enhancements

Possible improvements include:

* Adding automated Power Query data refresh
* Replacing manual KPI sheets with Power Query transformations
* Adding dynamic Year-over-Year analysis
* Adding Month-over-Month payment analysis
* Adding borrower risk segmentation
* Adding advanced conditional formatting
* Adding an executive summary page
* Migrating the dashboard to Power BI for scalable reporting

---

# 👤 Author

**Humaira**

Data Analyst | Excel | SQL | Power BI | Data Analytics

---

## 📌 Project Note

This project was created for **learning, portfolio, and demonstration purposes** to showcase practical Excel data analysis and dashboard development skills.

The workbook demonstrates the complete analytical workflow:

```text
Raw Data
   ↓
Data Preparation
   ↓
KPI Calculations
   ↓
PivotTables
   ↓
PivotCharts
   ↓
Slicers / Timelines
   ↓
Interactive Dashboard
```

The project highlights how Excel can be used not only for spreadsheet calculations, but also as a powerful tool for **business analysis and data storytelling**.

