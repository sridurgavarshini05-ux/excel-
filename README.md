# excel-
**Sales Operations Analytics: Excel Minor Project**

An Excel-based analysis of retail sales data, completed as the minor project of the Data Analytics / Data Science track at Unlox Academy.

**Overview**

The project simulates a real analyst task. A retailer operating across six Indian cities (Mumbai, Delhi, Chennai, Pune, Hyderabad, Bengaluru) shares five raw data tables. The head of sales needs answers to business questions before a leadership review.

The data is deliberately messy: duplicate order IDs, extra spaces in names, inconsistent city casing, orders linked to customers who do not exist, and sales of discontinued products. Finding and handling these issues is part of the task.

**Dataset**
Sheet	Description
Orders	About 2,000 orders (order id, date, customer, product, sales rep, quantity, price, discount, final amount, payment status, ship city)
Customers	300 customers (name, age, city, segment)
Products	60 products (category, cost, price, active flag)
Sales_Reps	30 sales reps (region, hire date, team lead)
Targets	Monthly sales targets per rep, Jan to Dec 2024
Project Structure
Section	Topic	Skills used
Q1	Lookups	VLOOKUP, INDEX + MATCH, two-way lookup, chained lookups, approximate match for discount slabs, self-join for team lead
Q2	Data cleaning	Duplicate detection, TRIM, COUNTIF, handling case differences and broken references
Q3	Conditional aggregation	SUMIFS, COUNTIFS, AVERAGEIFS, helper columns
Q4	Pivot tables and analysis	Region x category pivots, % of row total, quarterly grouping, profit margin
Q5	Advanced formulas	RANK, LARGE, SUMPRODUCT, TEXTJOIN, MAXIFS, cumulative totals
**Selected Results**
Question	Result
Order IDs that appear more than once	20
Revenue from South region, Electronics category	1,314,107.39
Highest single month x rep revenue value	216,421.52
Share of North region revenue from Apparel	14.02%
Total revenue for Q4 2024 (Oct to Dec)	64,59,442.98
Average profit margin on orders shipped to Delhi	44.90%
Rep ranked 7th by total revenue	Amit Singh
Cumulative revenue up to 31 Aug 2024	18,050,689.60
Second-best product in the West region	Smart Watch
Top 5 reps by revenue	Neha Desai, Shalini Patel, Deepak Shah, Nisha Menon, Arun Iyer
Highest single Grocery order shipped to Bengaluru	22,134
Example Formulas

Two-way lookup of a rep's target for a given month:

excel
=INDEX(Targets!$B$2:$M$31, MATCH("R012", Targets!$A$2:$A$31, 0), MATCH("Mar-24", Targets!$B$1:$M$1, 0))

Approximate match for a discount slab:

excel
=VLOOKUP(discount, $B$18:$C$21, 2, TRUE)

Quarterly revenue with a date range:

excel
=SUMIFS(Orders!$I$2:$I$2001, Orders!$B$2:$B$2001, ">="&DATE(2024,10,1), Orders!$B$2:$B$2001, "<="&DATE(2024,12,31))

Average margin for Delhi, with TRIM to handle stray spaces in city names:

excel
=SUMPRODUCT((TRIM(Orders!$K$2:$K$2001)="Delhi")*(Orders!$G$2:$G$2001-SUMIF(Products!$A$2:$A$61,Orders!$D$2:$D$2001,Products!$D$2:$D$61))/Orders!$G$2:$G$2001)/SUMPRODUCT(--(TRIM(Orders!$K$2:$K$2001)="Delhi"))

Concatenate the top 5 reps by revenue:

excel
=TEXTJOIN(" | ", TRUE, INDEX(Sales_Reps!$B$2:$B$31, MATCH(LARGE(SUMIF(Orders!$E$2:$E$2001, Sales_Reps!$A$2:$A$31, Orders!$I$2:$I$2001), {1,2,3,4,5}), SUMIF(Orders!$E$2:$E$2001, Sales_Reps!$A$2:$A$31, Orders!$I$2:$I$2001), 0)))
Key Learnings
Combining several tables with lookups before analysing.
Spotting data quality problems (duplicates, orphan records, inconsistent text) and handling them with formulas instead of editing the raw data.
Writing formulas that recalculate automatically, with no hardcoded values.
Choosing between pivot tables and formulas for the same question and cross-checking them.
How to Use
Download the workbook from this repository.
Open it in Microsoft Excel (2019 or later recommended; Microsoft 365 for the dynamic array formulas).
Go to the question sheets (Q1 to Q5). Answers are in the highlighted cells and are formula-based.
Tools

Microsoft Excel

Author

SRI DURGA VARSINI P, Data Science Intern, Unlox Academy GitHub

Acknowledgements

Project brief and dataset provided by Unlox Academy for learning purposes.
