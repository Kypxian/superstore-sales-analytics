# Power BI Sales Analytics — End-to-End Portfolio Project

A comprehensive sales analytics project built from **Sample Superstore data (25,035 orders, $5.6M revenue, 4 years of transactions)**. The pipeline spans raw CSV ingestion, Python-based data cleaning, star-schema normalisation, 38 DAX measures organised into 7 functional folders, and a 4-page interactive Power BI report with synced filters and custom theming.

---

## 📁 Repository Contents

| File | Description |
|---|---|
| `superstore.csv` | Raw source data — Sample Superstore with 42 columns across 4 years (2011–2014) |
| `clean_sales_final.xlsx` | Cleaned dataset after deduplication, column engineering, and type normalisation |
| `sales_powerbi_star.xlsx` | Pre-built star-schema model — 1 fact table + 5 dimension tables with surrogate keys |
| `Sales_Analytics_Portfolio.pbix` | Final Power BI report with all measures, pages, relationships, and custom theme applied |
| `Sales_Analytics_Report.tex` | LaTeX source for the full technical report — compile on Overleaf |

---

## 1. Data Cleaning & Feature Engineering

The raw `superstore.csv` contained 25,035 rows across 42 columns. The following cleaning pipeline was applied using **Python (pandas)**:

### Data Quality

- **Duplicates:** Row-level deduplication was performed by checking unique combinations of Order ID and Product ID — zero duplicates were found in the raw data.
- **Null Values:** All critical columns (Sales, Profit, Customer ID, Product ID) were validated for completeness. No nulls were present in key fields.
- **Date Integrity:** Order dates span 1 January 2011 to 31 December 2014. Ship dates extend slightly into 2015 (up to 7 January 2015). All dates are continuous with no gaps.

### Calculated Columns

Several analytical columns were derived from existing fields:

| Column | Formula | Purpose |
|---|---|---|
| `NetSales` | Sales − Discount | Revenue after discounts — the true top line |
| `ProfitMargin` | Profit ÷ Sales | Profitability ratio used in ratio analysis |
| `ShippingDays` | ShipDate − OrderDate | Delivery speed — key operational metric |
| `UnitPriceBeforeDiscount` | Sales ÷ Quantity | Pre-discount pricing for discount-rate calculation |
| `YearQuarter` | Year + Quarter | Simplified period key for time intelligence |
| `IsWeekend` | DayName ∈ {Saturday, Sunday} | Enables weekend vs weekday shipping analysis |

### Data Type Standardisation

Date columns were cast to `datetime64`, currency columns to `float64` with consistent rounding, and categorical columns (Category, Segment, Ship Mode, Market) to `category` dtype for compressed storage in the export.

---

## 2. Star-Schema Data Model

The flat cleaned table was normalised into a **star schema** — one fact table and five dimension tables — each linked by surrogate integer keys rather than natural keys. This approach improves query performance in Power BI and enforces referential integrity.

### Tables

| Table | Type | Rows | Columns | Description |
|---|---|---|---|---|
| `Fact_Sales` | Fact | 25,035 | 23 | Every order line — measures plus foreign keys |
| `DimDate` | Dimension | 1,461 | 14 | Calendar with attributes: Year, Quarter, Month, Week, Day, Weekend flag, Half-year period |
| `DimCustomer` | Dimension | 4,867 | 4 | CustomerID, CustomerName, Segment (Consumer / Corporate / Home Office) |
| `DimProduct` | Dimension | 8,790 | 5 | ProductID, ProductName, Category, SubCategory |
| `DimGeography` | Dimension | 3,596 | 7 | City, State, Region, Country, Continent, Market |
| `DimShipping` | Dimension | 4 | 3 | ShipMode (Standard Class / Second Class / First Class / Same Day), ShippingSpeed |

### Relationship Design

```
┌──────────┐     ┌───────────────────────────────────┐     ┌──────────────┐
│ DimDate  │────→│           Fact_Sales               │←────│ DimCustomer  │
│ (DateKey)│     │ (OrderDateKey, ShipDateKey,        │     │ (CustomerKey)│
└──────────┘     │  CustomerKey, ProductKey,          │     └──────────────┘
                 │  GeographyKey, ShippingKey)         │
┌──────────┐     │                                   │     ┌──────────────┐
│ DimProduct│←───│                                   │←────│ DimGeography │
│(ProductKey)│   └───────────────────────────────────┘     │(GeographyKey)│
└──────────┘                                               └──────────────┘
                                                    ┌──────────────┐
                                                    │ DimShipping  │
                                                    │ (ShippingKey)│
                                                    └──────────────┘
```

**Key design decisions:**

- **ShipDate as inactive relationship:** The `DimDate[DateKey]` → `Fact_Sales[ShipDateKey]` relationship is intentionally left **inactive**. Default measures aggregate by Order Date (the primary business calendar). The dedicated `Sales by Ship Date` measure activates this relationship on demand using `USERELATIONSHIP`, enabling side-by-side comparison of revenue recognition timing.

- **Surrogate keys:** Each dimension uses an auto-incrementing `Key` column as the primary key instead of natural IDs. This decouples the model from source-system changes and makes relationship traversal in Power BI faster.

- **DimDate as Date Table:** `DimDate[FullDate]` is marked as a **Date Table** in Power BI, enabling all time-intelligence functions (`TOTALYTD`, `SAMEPERIODLASTYEAR`, `DATESINPERIOD`, etc.) without requiring Power Query auto-date tables.

### Data Profile Summary

```
Data Range:      1 Jan 2011 – 31 Dec 2014 (4 years)
Total Orders:    25,035
Total Revenue:   $5,632,756
Total Profit:    $478,925 (8.5% margin)
Total Quantity:  88,255 units
Unique Products: 8,790
Unique Customers:4,867
Countries:       147
Categories:      Office Supplies (16,675), Furniture (4,682), Technology (3,678)
Segments:        Consumer (12,901), Corporate (7,529), Home Office (4,605)
Shipping Modes:  Standard Class (14,901), Second Class (5,041), First Class (3,764), Same Day (1,329)
Avg Order Value: $225
Avg Shipping:    4.0 days
```

---

## 3. DAX Measures — 38 Measures Across 7 Categories

All measures are defined on the `Fact_Sales` table and organised into **display folders** for clean model navigation. Each category addresses a specific analytical domain.

### 01. Base Measures — Foundation Layer

These are the primitive aggregations from which all derived measures are built. Every measure in categories 2–8 references one or more of these.

| Measure | DAX Pattern | Purpose |
|---|---|---|
| `Total Sales` | `SUM(Fact_Sales[Sales])` | Gross revenue before discounts |
| `Total Net Sales` | `SUM(Fact_Sales[NetSales])` | Revenue net of discounts applied |
| `Total Cost` | `SUM(Fact_Sales[Cost])` | Cost of goods sold |
| `Total Profit` | `SUM(Fact_Sales[Profit])` | Gross profit = NetSales − Cost |
| `Total Quantity` | `SUM(Fact_Sales[Quantity])` | Unit volume sold |
| `Total Discount` | `SUM(Fact_Sales[Discount])` | Total discount amount given |
| `Total Shipping Cost` | `SUM(Fact_Sales[ShippingCost])` | Freight cost incurred |
| `Order Count` | `DISTINCTCOUNT(Fact_Sales[OrderID])` | Number of unique transactions |

### 02. Ratio & Margin Measures — Profitability Analysis

These measures normalise raw aggregates into comparable ratios, enabling fair comparison across products, regions, and time periods regardless of scale.

- **`Profit Margin`** — `DIVIDE([Total Profit], [Total Sales], 0)`. Overall margin is 8.5%, with significant variation across categories (Technology margins are 2.1× higher than Furniture).
- **`Discount Rate`** — Discount as a percentage of pre-discount price. Reveals pricing strategy: Office Supplies receive the deepest discounts.
- **`Sales per Order`** and **`Profit per Order`** — Average transaction values. Useful for basket-size analysis.
- **`Avg Unit Price`** and **`Avg Discount per Unit`** — Per-unit economics at the item level.
- **`Cost Ratio`** and **`Shipping Cost % of Sales`** — Cost structure metrics; shipping costs represent 10.8% of total sales, a significant operational expense.

### 03. Time Intelligence — Period Comparisons

This is the most technically demanding category, using CALCULATE context transition and time-intelligence functions.

- **`Sales PY`** — Uses `SAMEPERIODLASTYEAR(DimDate[FullDate])` within a `CALCULATE` block to fetch the prior-year equivalent. Enables direct year-over-year comparison at any granularity (day, month, quarter, year).
- **`Sales YoY %`** — `DIVIDE([Total Sales] - [Sales PY], [Sales PY], 0)`. Growth rate that works at any date hierarchy level.
- **`Sales YTD` / `Sales MTD` / `Sales QTD`** — Uses `TOTALYTD`, `TOTALMTD`, and `TOTALQTD` respectively. These filter the date context to the current partial period, allowing real-time progress tracking.
- **`Profit YTD`** — Year-to-date profit, used on the Executive Dashboard to track profitability momentum.
- **`Profit vs PY`** — Absolute profit change from the prior year — a direct variance measure.
- **`Sales Rolling 3M` / `Sales Rolling 12M`** — Uses `DATESINPERIOD` with `-3` and `-12` month windows. Rolling averages smooth out seasonal spikes and reveal underlying trends.
- **`Sales by Ship Date`** — The only measure that activates the **inactive** ShipDate relationship: `CALCULATE([Total Sales], USERELATIONSHIP(Fact_Sales[ShipDateKey], DimDate[DateKey]))`. This enables the dual-line chart on the Shipping page comparing revenue by order date vs ship date.

### 04. Product Analysis — Portfolio Performance

- **`Products Sold`** — `DISTINCTCOUNT(Fact_Sales[ProductKey])`. Tracks product breadth sold in each period.
- **`Avg Sales per Product`** — Revenue concentration metric. Lower values indicate a more diversified product mix.
- **`Category Sales %`** — Uses `ALLEXCEPT` to compute each category's share within the current filter context. On the Product Analysis page, this measure enables the table showing sub-category contribution to parent category.
- **`Top Product by Sales`** — `TOPN(1, VALUES(DimProduct[ProductName]), [Total Sales])`. Returns the highest-revenue product for the current selection context.

### 05. Customer Analysis — Customer Value

- **`Active Customers`** — `DISTINCTCOUNT(Fact_Sales[CustomerKey])`. A core KPI on the Executive Dashboard.
- **`Revenue per Customer`** — Customer-level average revenue, used to segment high-value vs low-value customers.
- **`Orders per Customer`** — Purchase frequency. Higher values indicate customer loyalty and repeat business.
- **`New Customers`** — Uses a virtual table (`ADDCOLUMNS` + `FILTER`) to identify customers whose first-ever order falls within the current date range. This measure required careful handling of filter context to avoid counting returning customers as new.
- **`Repeat Purchase Rate`** — The percentage of customers who have placed more than one order. A critical retention metric.

### 06. Geography & Market Analysis — Regional Performance

- **`Market Sales %`** — Each market's contribution to global sales, computed with `ALLEXCEPT` to respect the Market-level granularity while ignoring lower geography levels.
- **`Regional Profitability`** — Profit margin at the Region level. Reveals which geographic areas are profitable and which may require strategic review.

### 07. Shipping Analysis — Operational Efficiency

- **`Avg Shipping Days`** — Currently 4.0 days across all modes. Broken down by `DimShipping[ShipMode]` in the Shipping page report.
- **`Avg Shipping Cost`** and **`Shipping Cost per Unit`** — Freight cost analysis. Standard Class shipping has the lowest cost per unit, as expected.
- **`On-Time Ship Rate`** — Calculates the proportion of orders shipped via express methods (`ShippingSpeed = "Express"`). Useful for logistics performance monitoring.

### 08. Dynamic KPIs — Performance Vs Targets

- **`KPI Sales vs Target`** — `VAR TargetAmount = 6000000 RETURN DIVIDE([Total Sales], TargetAmount, 0)`. Current progress is **93.9%** ($5.63M of $6M goal).
- **`KPI Profit vs Target`** — `VAR TargetProfit = 500000 RETURN DIVIDE([Total Profit], TargetProfit, 0)`. Current progress is **95.8%** ($478.9K of $500K goal).
- **`YTD Status`** — A text measure returning `"▲ Above Last Year"` or `"▼ Below Last Year"` by comparing `[Sales YTD]` against `[Sales PY]`. Used in KPI cards for instant directional context.

---

## 4. Report Design & Visual Architecture

The report uses a **custom navy-and-coral theme** with white card backgrounds and consistent Segoe UI typography. A single **Year slicer** on the Dashboard is synced across all four pages via Power BI's built-in slicer sync feature — changing the year on any page updates every visual in the report.

### Page 1: Executive Dashboard

Designed for a C-level audience needing a 30-second business pulse check.

| Visual Element | Configuration | Insight |
|---|---|---|
| **4 KPI Cards** | Total Sales ($5.63M), Total Profit ($478.9K), Active Customers (4,867), Profit Margin (8.5%) | Instant snapshot of business health |
| **Line Chart** | Monthly Sales (axis: DimDate[Month]) + Sales PY (secondary line) | Shows seasonality — Q4 spikes are consistent across all four years |
| **Donut Chart** | Sales by Category | Office Supplies = 49%, Technology = 29%, Furniture = 22% |
| **Year Slicer** | Vertical list, multi-select, synced to all pages | Single point of filter control |

The line chart includes drill-down from year to quarter to month, powered by the DimDate hierarchy. The donut chart uses the custom colour palette to highlight the three categories distinctly.

### Page 2: Product Analysis

Designed for product managers and category owners.

| Visual Element | Configuration | Insight |
|---|---|---|
| **Clustered Bar Chart** | Sales + Profit by Category, dual-axis | Technology generates the highest profit despite lower volume than Office Supplies |
| **Table** | Category → SubCategory drill-down — columns: Sales, Profit, Profit Margin, Category Sales % | Enables granular portfolio analysis |
| **Scatter Chart** | X-axis: Avg Unit Price, Y-axis: Profit Margin, Legend: Category | Reveals the "star quadrant" — Technology products cluster in the top-right (high price, high margin) |

The scatter chart is the key differentiator on this page. It visually separates premium products (high unit price, high margin) from low-margin commodities, enabling strategic pricing decisions.

### Page 3: Customer & Geography

Designed for regional sales managers and marketing teams.

| Visual Element | Configuration | Insight |
|---|---|---|
| **Bing Map** | Sales by Country, bubble size = Total Sales | APAC, LATAM, and US markets dominate — visual confirmation of geographic concentration |
| **Table** | Market breakdown — columns: Active Customers, Sales, Revenue per Customer, Orders per Customer | US market leads in revenue per customer ($1,157), while APAC has the highest customer count |
| **Clustered Bar Chart** | Sales + Profit by Region | Central and South regions are strong; EMEA and Africa show growth opportunity |

The `Revenue per Customer` and `Orders per Customer` measures differentiate this page from a simple sales-by-region view, shifting the focus from pure volume to customer value.

### Page 4: Shipping & Time Trends

Designed for operations and logistics stakeholders.

| Visual Element | Configuration | Insight |
|---|---|---|
| **Dual Line Chart** | Sales by Order Date vs Sales by Ship Date (using `USERELATIONSHIP`) | A 3–5 day lag between order date and ship date revenue recognition is visible |
| **Gauge** | Avg Shipping Days (4.0) target needle at 3.0 days | Standard Class shipping (4.2 days average) is the main contributor to the gap |
| **Clustered Bar Chart** | Order Count by Ship Mode | Standard Class accounts for 59.5% of all orders — the dominant shipping preference |

The dual-line chart is the technical highlight of this page. It contrasts the default aggregation (by Order Date, using the active relationship) against a `CALCULATE` with `USERELATIONSHIP` (by Ship Date, using the inactive relationship). Both lines are plotted on the same time axis, making the timing shift visible at a glance.

---

## Tools & Methods Used

| Tool | Application |
|---|---|
| **Power BI Desktop** | Data modelling, DAX measure authoring, report layout, visual configuration, slicer sync, custom theme import |
| **Tabular Editor 2** | Batch creation of all 38 measures via C# script; applied format strings (`$#,##0`, `0.0%`), display folders, and data types in a single operation |
| **Python (pandas, NumPy)** | Raw CSV ingestion, data quality checks, feature engineering, star-schema normalisation, Excel export with 6 sheets |
| **Excel (openpyxl)** | Multi-sheet .xlsx output with clean column ordering and no index columns |

---

## How to Rebuild

```bash
git clone https://github.com/your-username/sales-analytics-powerbi
cd sales-analytics-powerbi
```

**Option A — Open the completed report:**
1. Open `Sales_Analytics_Portfolio.pbix` in Power BI Desktop
2. All measures, relationships, and pages are pre-configured

**Option B — Rebuild from the star-schema data:**
1. Import `sales_powerbi_star.xlsx` (6 sheets → 6 Power BI tables)
2. Set up relationships as documented in Section 2
3. Mark `DimDate[FullDate]` as Date Table
4. Create the 38 measures from the DAX reference documentation
5. Apply the custom theme and build the 4 report pages

---

## Skills Demonstrated

- **Data Engineering:** Raw CSV ingestion, data quality validation, feature engineering, flat-to-star normalisation
- **DAX:** CALCULATE context transition, time intelligence (SAMEPERIODLASTYEAR, TOTALYTD, DATESINPERIOD), USERELATIONSHIP for inactive relationships, iterator patterns, virtual tables with ADDCOLUMNS/FILTER
- **Data Modelling:** Surrogate key design, star-schema methodology, granularity selection, active vs inactive relationship strategy
- **Power BI Reporting:** Multi-page layout, drill-down hierarchies, synced slicers, custom colour theming, dual-axis charts, KPI cards, map visualisation
- **Automation:** Tabular Editor C# scripting for bulk measure operations
- **Analytical Storytelling:** Choosing the right visual for each audience (KPI cards for executives, scatter plots for analysts, maps for geographic stakeholders)

---

## 📸 Report Screenshots

| Page | Preview |
|------|---------|
| 📊 **Executive Dashboard** | <img src="https://raw.githubusercontent.com/Kypxian/superstore-sales-analytics/main/screenshots/dashboard.png" alt="Dashboard" width="400"/> |
| 📦 **Product Analysis** | <img src="https://raw.githubusercontent.com/Kypxian/superstore-sales-analytics/main/screenshots/products.png" alt="Products" width="400"/> |
| 🌍 **Customer & Geography** | <img src="https://raw.githubusercontent.com/Kypxian/superstore-sales-analytics/main/screenshots/customers.png" alt="Customers" width="400"/> |
| 🚚 **Shipping & Time Trends** | <img src="https://raw.githubusercontent.com/Kypxian/superstore-sales-analytics/main/screenshots/shipping.png" alt="Shipping" width="400"/> |