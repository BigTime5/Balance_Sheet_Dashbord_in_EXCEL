# Balance Sheet Dashboard in Excel: A Game-Changing Financial Solution

🚀 **Problem Solved!** Say goodbye to static, error-prone financial reporting! This project delivers a **dynamic, interactive balance sheet dashboard** built in Excel, empowering businesses and financial professionals to effortlessly transform complex trial balance data into actionable insights. By leveraging **Power Query**, **Power Pivot**, **DAX**, and **Pivot Tables**, this reusable dashboard streamlines financial analysis, saving time and ensuring accuracy. Whether you're an accountant or a business owner, this tool unlocks the power to track key metrics like **working capital** and **debt-to-equity ratio** with ease, making financial decision-making faster and smarter.

## Why This Project Was Necessary

Financial reporting can be a headache—manually crunching numbers, reconciling trial balances, and calculating retained earnings often leads to errors and wasted hours. Traditional spreadsheets lack the flexibility to adapt to changing data or provide real-time insights. This project was born to **solve these pain points** by creating a **scalable, automated, and interactive dashboard** that:

- **Simplifies Data Management**: Transforms raw trial balance and chart of accounts data into a structured, analysis-ready format.
- **Saves Time**: Automates calculations for retained earnings, working capital, and financial ratios, eliminating repetitive manual work.
- **Enhances Decision-Making**: Provides dynamic visualizations and slicers for instant insights into financial health across months.
- **Ensures Reusability**: Allows annual updates with minimal effort, making it a long-term solution for financial reporting.
- **Empowers Non-Accountants**: Offers intuitive tools and visuals that make complex financial data accessible to all.

This dashboard is a **must-have** for anyone looking to elevate their financial analysis, reduce errors, and make data-driven decisions with confidence.

## Key Steps to Success

1. **Master Your Data Foundation**
   - Harness two core datasets: **Trial Balance** (monthly postings) and **Chart of Accounts** (account codes and classifications).
   - Accurately calculate month-to-month profits/losses to handle **retained earnings** like a pro.

2. **Transform with Power Query**
   - Seamlessly import and reshape trial balance data using **Power Query**.
   - Convert data to a long format (dates and amounts) for streamlined analysis.
   - Apply consistent date formatting (e.g., "1 January") for clarity.
   - Integrate the Chart of Accounts with case-sensitive M code for precision.
   - Preserve original data with reference queries while enabling adjustments.

3. **Streamline Trial Balance Adjustments**
   - Create copies of the trial balance to manage **closing accounts** and **retained earnings**.
   - Filter income statement lines (codes < 10,000) and reset values by multiplying by -1.
   - Group data by month to compute profits/losses for retained earnings.
   - Append adjusted data to the trial balance for a unified dataset.

4. **Calculate Cumulative Retained Earnings**
   - Build a custom column using **List.Sum** and **List.FirstN** in Power Query for running totals.
   - Ensure accurate cumulative retained earnings with indexed summations.
   - Format outputs as decimals to align with the trial balance.

5. **Create a Dynamic Calendar Table**
   - Generate a date range with **List.Min** and **List.Max** for a comprehensive timeline.
   - Add month values and abbreviated names for clear visualizations.
   - Format as a table for seamless integration into the data model.

6. **Build a Robust Data Model in Power Pivot**
   - Load **Chart of Accounts**, **adjusted Trial Balance**, and **Calendar** tables as connections.
   - Establish a **star schema**: Trial Balance as the fact table, with Chart of Accounts and Calendar as dimension tables.
   - Connect tables via foreign keys (account codes and dates).
   - Sort months chronologically for intuitive analysis.

7. **Power Up with Pivot Tables and DAX**
   - Create **Pivot Tables** to aggregate balance sheet data (assets, liabilities, equity).
   - Use **DAX** and the **CALCULATE** function to compute current and previous period values.
   - Rename measures (e.g., "Previous") and add subtotals for clarity.

8. **Unlock Insights with Cube Functions and Metrics**
   - Leverage **Cube functions** to extract measures for current assets, liabilities, and **working capital**.
   - Calculate critical ratios like **debt-to-equity** using long-term loans and equity.
   - Compute working capital with simple, accurate formulas.

9. **Visualize with Impactful Charts and Slicers**
   - Build a summary table for monthly **current assets** and **liabilities**.
   - Create **stacked bar charts** for working capital and **waterfall charts** for changes.
   - Add **slicers** for dynamic month filtering, styled for seamless dashboard integration.
   - Apply **conditional formatting** to highlight variances and debt-to-equity thresholds (e.g., waffle charts).
   - Include a dynamic text box for real-time **debt-to-equity percentage** display.

10. **Deliver a Polished Dashboard**
    - Format data labels (e.g., millions with two decimals) for crystal-clear readability.
    - Remove clutter (titles, legends) and align colors for a professional look.
    - Ensure the dashboard dynamically updates with slicer selections for real-time insights.

![Balance-sheet-dashboard](assets/images/dashboard_balance-sheet.gif)


## Value Packed: Why This Dashboard Shines

- **Interactive & User-Friendly**: Slicers enable effortless month-by-month analysis, making data exploration intuitive.
- **Time-Saving & Reusable**: Update trial balance data annually with a single refresh, minimizing rework.
- **Insightful Metrics**: Tracks **working capital**, **debt-to-equity**, and variances to fuel informed decisions.
- **Powerful Tools**: Combines **Power Query**, **Power Pivot**, **DAX**, and **Pivot Tables** for a robust solution.
- **Accessible to All**: Simplifies complex financial data for accountants and non-accountants alike.

This dashboard transforms raw financial data into a **strategic asset**, delivering clarity, efficiency, and actionable insights for businesses of all sizes.

## Prerequisites

- Microsoft Excel with **Power Query** and **Power Pivot** enabled.
- Basic knowledge of accounting (trial balance, chart of accounts, retained earnings).
- Familiarity with Excel functions and data modeling.

## How to Use

1. Import **trial balance** and **chart of accounts** datasets.
2. Follow the steps above to transform data, build the data model, and create visualizations.
3. Use **slicers** to filter by month and analyze financial metrics.
4. Update annually by refreshing trial balance data in **Power Query**.

## Value Delivered

🌟 This project tackles the chaos of financial reporting head-on, delivering a **scalable, automated, and visually stunning dashboard** that empowers users to monitor financial health with confidence. Whether you're managing a small business or analyzing corporate finances, this solution brings **clarity, efficiency, and strategic insight** to the table. Ready to revolutionize your financial analysis? Dive in and build your dashboard today!
