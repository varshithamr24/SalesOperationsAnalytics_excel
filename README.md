📊 Excel Sales Data Analysis & Business Intelligence

📌 Project Overview

This project demonstrates practical Microsoft Excel data analysis and business intelligence techniques using sales order data.

The workbook combines information from multiple sheets such as Orders, Products, Customers, Sales Representatives, and Targets to answer real-world business questions related to revenue, customers, products, regions, sales representatives, and profitability.

The project is organized into three major analytical sections:

Q3 – Conditional Aggregation

Q4 – Pivot Tables

Q5 – Advanced Excel Analysis

🎯 Project Objectives

1.Analyze sales revenue using multiple business conditions.

2.Calculate regional, category, customer-segment, and sales-representative performance.

3.Use Excel formulas for conditional aggregation.

4.Build PivotTables for multidimensional sales analysis.

5.Rank sales representatives based on total revenue.

6.Identify top-performing products and representatives.

7.Analyze cumulative and quarterly revenue.

8.Calculate and compare profit margins.

9.Convert raw business data into meaningful insights.

🛠️ Tools & Technologies

Microsoft Excel

Excel Formulas

PivotTables

Data Aggregation

Data Analysis

Business Intelligence

Excel Functions Used

SUMIFS · COUNTIFS · AVERAGEIFS · IF · MAX · MAXIFS · RANK.EQ · LARGE · INDEX · MATCH · TEXTJOIN · VLOOKUP

📈 Q3 – Conditional Aggregation

This section analyzes sales data using multiple conditions.

Total revenue for a specific quarter and city.

Number of completed orders satisfying discount conditions.

Average order value based on product category and customer segment.

Combined revenue of top sales representatives.

Sales representative performance against targets.

Revenue from customers within a specific age range and segment.

📊 Q4 – Pivot Tables

PivotTables are used to summarize and analyze large amounts of sales data.

Region vs Product Category

Revenue is summarized using Sales Region as rows, Product Category as columns, and Sum of Final Amount as values.

Monthly Sales Representative Revenue

A PivotTable uses Order Month as rows, Sales Representative ID as columns, and Sum of Final Amount as values.

Percentage of Row Total

The PivotTable uses Show Values As → % of Row Total to identify category contribution by region.

Quarterly Revenue

Order dates are grouped into quarters to analyze quarterly revenue trends.

Profit Margin

Profit margin is calculated as:

Profit Margin = (Unit Price - Unit Cost) / Unit Price

🚀 Q5 – Advanced Excel Analysis

Q5.1 – Sales Representative Ranking

Sales representatives are ranked according to total revenue, where the highest revenue receives Rank 1.

Example:

=RANK.EQ(I2,$I$2:$I$31,0)

Q5.2 – Cumulative Revenue

Revenue from January 2024 through August 2024 is calculated using a date-based SUMIFS condition.

Q5.3 – Top Product in the West Region

Orders are filtered to the West region, revenue is summarized by product, and products are ranked to identify the 2nd highest revenue-generating product.

Q5.4 – Top 5 Sales Representatives

The top five representatives are combined into a single cell using TEXTJOIN.

Example:

=TEXTJOIN(" | ",TRUE,Top_5_Rep_Names)

Q5.5 – Highest Order Value

The maximum order value is identified where:

Ship City = Bengaluru

Product Category = Grocery

Example:

=MAXIFS(Orders!$I$2:$I$2001,
Orders!$K$2:$K$2001,"Bengaluru",
Orders!$L$2:$L$2001,"Grocery")

📌 Key Results

Analysis

Result

Q3.1 – Mumbai Q3 Revenue

1,217,239.01

Q3.2 – Discount >15% & Completed Orders

723

Q3.3 – Electronics & Corporate Average

13,981.40156

Q4.1 – South + Electronics Revenue

1,314,107.39

Q4.2 – Highest Pivot Value

216,421.52

Q4.3 – North Apparel Contribution

14.02%

Q4.4 – Q4 2024 Revenue

6,459,442.98

Q5.1 – Rank #7 Representative

Amit Singh

Q5.2 – Revenue through Aug 2024

18,050,689.60

Q5.3 – 2nd Highest West Product

Smart Watch

Note: Results may change if the underlying workbook data or formulas are modified.

💡 Business Insights

This project shows how Excel can help businesses:

Identify high-performing sales representatives.

Compare regional revenue.

Determine strong product categories.

Track monthly and quarterly sales.

Analyze customer segments.

Identify top-performing products.

Evaluate profitability.

Support data-driven decisions.

📚 Skills Demonstrated

Data Cleaning

Data Transformation

Excel Formula Development

Conditional Aggregation

Multi-Sheet Data Analysis

PivotTable Creation

Revenue Analysis

Sales Performance Analysis

Ranking

Top-N Analysis

Profit Margin Calculation

Business Intelligence

▶️ How to Use

Download or clone this repository.

Open the Excel workbook.

Review the Orders, Products, Customers, Sales Representatives, and Targets sheets.

Open the Q3, Q4, and Q5 analysis sections.

Review the formulas and PivotTables.

Modify filters or formulas to perform additional analysis.

🌟 Project Outcome

The project transforms raw sales records into structured business insights using Excel. It demonstrates the ability to combine formulas, helper columns, PivotTables, ranking techniques, and multi-sheet data to solve practical analytical problems.
