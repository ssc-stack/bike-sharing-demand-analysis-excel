# Bike Sharing Demand Analysis Using Excel

**Prepared by:** Shashank Singh Chauhan  
**Project Organization:** NextHikes IT Solutions  
**Submission Date:** November 20, 2026

## Overview
Analysis of bike-sharing rental demand by time, weather, season and user type using Excel dashboards, statistical forecasting, and anomaly indicators.

## Repository contents
- `datasets/` — three source Excel datasets
- `reports/` — final analytical workbook and summary report (Word and PDF)
- `presentation/` — editable project presentation
- `scripts/` — VBA module for Excel refresh/report automation

## Key findings
- 1,000 hourly records analyzed
- 58,304 total rentals
- Registered users: approximately 91.6% of rentals
- Peak demand hour: 5 PM

## How to view
1. Open `reports/Bike_Sharing_Final_Automated_Report_Updated.xlsx` in Microsoft Excel.
2. Explore the Executive Dashboard, demand segments, daily trends, data quality, and forecast worksheets.
3. Review the PDF summary and PowerPoint presentation for methodology and recommendations.

## Automation
The workbook includes Excel-based analysis and statistical forecasting. To use the separately provided VBA module, save a working copy of the workbook as `.xlsm`, import `scripts/Bike_Sharing_Automation.bas` through the VBA editor (Alt+F11), and run `RefreshBikeSharingReport`. Test macros in desktop Excel before relying on refreshed results. New data is not automatically ingested without running the refresh workflow.

## Author
Shashank Singh Chauhan  
[GitHub](https://github.com/ssc-stack) · [LinkedIn](https://www.linkedin.com/in/shashankchauhan-)
