# Arshath's Power BI Dashboard

## 1. Project Overview

This project is an interactive business intelligence dashboard created with Microsoft Power BI. It is designed to provide insightful data visualizations and analysis based on a comprehensive data source managed in a macro-enabled Excel workbook. The dashboard allows users to explore key metrics, identify trends, and filter data dynamically.

## 2. Features

* **Interactive Visualizations**: The dashboard includes multiple charts, graphs, and tables that respond to user interactions.
* **Dynamic Filtering**: Utilizes slicers and filters to allow users to drill down into specific data segments.
* **Multi-Page Report**: The report is likely structured with several pages, each focusing on a different aspect of the analysis.
* **Custom Theme**: The dashboard uses a custom theme (`Innovate.json` and `CY24SU10.json`) for a consistent and professional appearance.
* **Complex Data Model**: The dashboard connects to a detailed data model which is managed and refreshed from the source Excel file.

## 3. Technology Stack

* **Visualization & Dashboarding**: `Microsoft Power BI`
* **Data Source & Preprocessing**: `Microsoft Excel` (.xlsm)

## 4. Data Source

The single source of truth for this dashboard is the Excel file **`ARSHATH-1.xlsm`**. This file is not just a simple data table; it is a structured workbook containing:

* **Multiple Worksheets**: Over 15 worksheets with raw and processed data.
* **Pivot Tables & Charts**: The workbook contains numerous pre-built Pivot Tables and Charts, suggesting that significant preliminary analysis is performed within Excel itself.
* **VBA Macros**: The `.xlsm` format and the presence of a `vbaProject.bin` file indicate that macros may be used for data cleaning, automation, or custom calculations within the workbook.

## 5. Getting Started

To view and interact with the dashboard, please follow these steps.

### Prerequisites

* You must have **[Microsoft Power BI Desktop](https://powerbi.microsoft.com/desktop/)** installed on your machine.

### Setup Instructions

1.  **Download the Files**: Download both `ARSHATH_pbi.pbix` and `ARSHATH-1.xlsm`.
2.  **Maintain Directory Structure**: Place both files in the **same folder**. This is critical for Power BI to correctly locate its data source.
3.  **Open the Dashboard**: Open the `ARSHATH_pbi.pbix` file using Power BI Desktop. The dashboard should load with the existing data.

### How to Refresh Data

If you make any changes to the data in the `ARSHATH-1.xlsm` file, you will need to refresh the Power BI dashboard to see the updates.

1.  **Update the Excel File**: Open `ARSHATH-1.xlsm`, make your data changes, and save the file.
2.  **Refresh in Power BI**: In Power BI Desktop, navigate to the **Home** tab on the ribbon and click the **Refresh** button. Power BI will reconnect to the Excel file and update all visualizations with the new data.
