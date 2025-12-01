# MacroFactSheet Project

---

## Project Overview

MacroFactSheet is Excel-based interactive macroeconomic factsheet built on **IMF World Economic Outlook, October 2025** data.  
The workbook lets you explore macro indicators for **194 countries** across **1980–2030** with dynamic charts and a country map.

<img width="1493" height="644" alt="image" src="https://github.com/user-attachments/assets/a31d5f4a-46d2-4743-989a-031a4f74a444" />
<img width="1491" height="720" alt="image" src="https://github.com/user-attachments/assets/ea3af7d2-a299-4a04-bc0b-ed697fb42afb" />

Users can:

- select any country (from 194 available) and a base year in the range **1980–2030**;
- view a panel of **19 macroeconomic indicators** grouped into **4 sectors**:
  - Production indicators
  - Employment indicators
  - Economic indicators
  - External sector indicators
- see an image map of the selected country;
- switch between sectors with buttons and automatically rebuild the chart matrix;
- analyse indicator dynamics in a **fixed window of -10 / +5 years** around the selected year  
  (to see longer-term trends and future perspective).

---

## Features

- Interactive country and base-year selection.
- 4 sector buttons that rebuild charts for all indicators in the chosen sector.
- Dynamic charts showing a -10 / +5 year window around the selected year.
- Country map display that updates together with the country selection.
- Power Query pipeline for loading and transforming IMF WEO data.
- VBA procedures to refresh only required queries and charts.
- Dark-theme dashboard layout suitable for visual comfort and data perception.

---

## Data & Coverage

- Source: **IMF World Economic Outlook, October 2025**.
- Countries: **194**.
- Time range: **1980–2030** (full grid, including future projections where available).
- Indicators: **19 unique macro indicators** across 4 sectors.
- Data size: **≈180,000 rows**

---

##  Performance notes

- The underlying data model contains about **180k rows**.
- When you change the **country**, **base year** or **sector**, Excel needs to:
  - refresh relevant Power Query tables,
  - recalculate the ChartData matrix,
  - rebuild and redraw charts.
- Because of this, updating the dashboard for a new selection may take **a few seconds** - 
  this is normal behaviour.


## Prerequisites

Before you begin, ensure you have:

- **Microsoft Excel Desktop**  
  - Excel 365 / 2019 / 2021 (Windows or Mac).  
  - The project uses **VBA**, **Power Query** and the **chart object model**.  
  - It will **not** work in Excel Online or other limited clients.
    
---

## ⚠️ Important Excel Settings

For the workbook to work correctly, you **must** configure the following.

### 1. Enable macros and active content

When you open `MacroFactSheet.xlsx`, Excel will show a security warning at the top.  
Click:

> **Enable Content** / **Enable Macros**

If you do not enable macros, buttons and automatic refresh will not work.

### 2. Decimal and thousands separators

The workbook expects the following separators:

- **Decimal separator:** `.`
- **Thousands separator:** `,`

In Excel:

1. Go to **File → Options → Advanced**.  
2. Under **Editing options**, **untick** `Use system separators`.  
3. Set:
   - `Decimal separator` = `.`
   - `Thousands separator` = `,`

This is required so that Power Query and numeric formatting work consistently.

---

## How to Use

1. On the **Menu** sheet:
   - Choose a **country** from the selector.
   - Choose a **base year** within **1980–2030**.
2. The dashboard will:
   - update the **indicator table** for the chosen country and year;
   - show the **map** of the selected country;
   - display **charts** for the active sector in a -10 / +5 year window around the selected year.
3. Use the **sector buttons** (“Production”, “Employment”, “Economic”, “External”) to switch sectors.  
   The chart matrix will rebuild automatically for the selected sector.

---
