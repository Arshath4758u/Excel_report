Arshath's Power BI & Excel Project

1. Project Overview

This project consists of two main files that work together:

ARSHATH_pbi.pbix: An interactive business intelligence dashboard created with Microsoft Power BI.

ARSHATH-1.xlsm: A macro-enabled Excel workbook that serves as the complete data source and pre-processing engine for the dashboard.

The Power BI dashboard is designed to provide insightful data visualizations and analysis based on the comprehensive data managed in the Excel workbook. The dashboard allows users to explore key metrics, identify trends, and filter data dynamically.

2. Features

Interactive Visualizations: The dashboard includes multiple charts, graphs, and tables that respond to user interactions.

Dynamic Filtering: Utilizes slicers and filters to allow users to drill down into specific data segments.

Multi-Page Report: The report is structured with several pages, each focusing on a different aspect of the analysis.

Custom Theme: The dashboard uses a custom theme (Innovate.json and CY24SU10.json) for a consistent and professional appearance.

Complex Data Model: The dashboard connects to the detailed data model managed and refreshed from the source Excel file.

3. Technology Stack

Visualization & Dashboarding: Microsoft Power BI

Data Source & Preprocessing: Microsoft Excel (.xlsm)

4. Data Source (ARSHATH-1.xlsm)

The single source of truth for this dashboard is the Excel file ARSHATH-1.xlsm. This file is not just a simple data table; it is a structured workbook containing:

Multiple Worksheets: Over 15 worksheets with raw and processed data.

Pivot Tables & Charts: The workbook contains numerous pre-built Pivot Tables and Charts, suggesting that significant preliminary analysis is performed within Excel itself.

VBA Macros: The .xlsm format and the presence of a vbaProject.bin file indicate that macros may be used for data cleaning, automation, or custom calculations within the workbook. Macros must be enabled when opening this file for it to function correctly.

5. Getting Started & Data Refresh Workflow

To use this project, you must have both files and follow the correct workflow.

Prerequisites

You must have Microsoft Power BI Desktop installed.

You must have Microsoft Excel installed.

Setup Instructions

Download the Files: Download both ARSHATH_pbi.pbix and ARSHATH-1.xlsm.

Maintain Directory Structure: Place both files in the same folder. This is critical for Power BI to correctly locate its data source.

Open the Dashboard: Open the ARSHATH_pbi.pbix file using Power BI Desktop. The dashboard should load with the existing data.

How to Refresh Data (The Workflow)

If you need to make any changes to the data, you must edit the Excel file first.

Update the Excel File:

Open ARSHATH-1.xlsm.

Enable Macros/Content if prompted by Excel.

Make your data changes in the worksheets.

Save and close the Excel file.

Refresh in Power BI:

Open the ARSHATH_pbi.pbix file.

Navigate to the Home tab on the ribbon.

Click the Refresh button.

Power BI will reconnect to the ARSHATH-1.xlsm file, load your changes, and update all visualizations automatically.
