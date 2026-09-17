# SuperStore-Retail-Dashboard
SuperStore Retail Data analysis Using Excel


# 📊 Superstore Sales Dashboard — Excel (Power Query + Power Pivot + DAX)

An end-to-end interactive business intelligence dashboard built entirely in Microsoft Excel. Raw Superstore data is cleaned and reshaped with **Power Query**, modelled as a **star schema** in **Power Pivot**, analysed with **DAX** measures, and visualised through **pivot charts, charts, slicers, and filters**.

**Dataset:** Sample Superstore · 9,994 order lines · 3 Jan 2014 – 30 Dec 2017 · US retail

---

## 🎯 Project Objective

Turn three disconnected raw sheets (`Orders`, `Returns`, `People`) into a single, refreshable, self-service dashboard that answers sales and profitability questions without formulas or manual rework.

Specific goals:

- Automate cleaning and joining of the raw sheets in **Power Query** so the whole report refreshes with one click.
- Replace VLOOKUP-based flat tables with a proper **star-schema data model** in Power Pivot.
- Build reusable **DAX measures** for KPIs, time intelligence, and growth analysis.
- Let stakeholders slice results by **region, category, segment, ship mode, and date** on demand.
- Identify where the business is **losing money despite growing revenue**.

---

## ❓ Business Questions Answered

| # | Business Question | Where It's Answered |
|---|-------------------|---------------------|
| 1 | What are total sales, profit, quantity, and profit margin for the selected period? | KPI cards |
| 2 | How have sales and profit trended across 2014–2017? | Year/month trend chart |
| 3 | Which category and sub-category generate the most sales vs the most profit? | Category pivot chart |
| 4 | Which sub-categories are **loss-making**? | Profit-by-sub-category bar |
| 5 | How do the four regions compare on sales and margin? | Region column chart |
| 6 | Which states destroy profit and should be reviewed? | State ranking table |
| 7 | What impact does discounting have on profitability? | Discount-band analysis |
| 8 | Which customer segment is most valuable? | Segment breakdown |
| 9 | Who are the top 10 customers and what share do they hold? | Top-N pivot chart |
| 10 | What is the return rate and the value of returned orders? | Returns KPI |
| 11 | How long does shipping take by ship mode? | Ship-mode analysis |
| 12 | How does each year compare to the previous year (YoY %)? | Time-intelligence measures |

---

## 🔄 Project Workflow

```
Raw Data: RawOrders (9,994) · RawReturns (296) · RawPeople (4)
          │
          ▼
┌──────────────────────────────────────┐
│  1. POWER QUERY (ETL)                │
│  • Promote headers, set data types   │
│  • Trim/clean text, fix Postal Code  │
│    as text (preserve leading zeros)  │
│  • Merge RawPeople → Regional Manager│
│  • Merge RawReturns → Returned flag  │
│  • Replace nulls with "No"           │
│  • Remove duplicates for dim tables  │
│  • Generate DimDate & DimShipDate    │
│    (List.Dates custom query)         │
└──────────────────────────────────────┘
          │  Load to Data Model
          ▼
┌──────────────────────────────────────┐
│  2. POWER PIVOT (Star Schema)        │
│  • 1 fact + 6 dimension tables       │
│  • One-to-many relationships         │
│  • DimDate marked as Date Table      │
│  • Category → Sub-Category hierarchy │
└──────────────────────────────────────┘
          │
          ▼
┌──────────────────────────────────────┐
│  3. DAX MEASURES (Business logic)    │
└──────────────────────────────────────┘
          │
          ▼
┌──────────────────────────────────────┐
│  4. PIVOT TABLES & PIVOT CHARTS      │
└──────────────────────────────────────┘
          │
          ▼
┌──────────────────────────────────────┐
│  5. SLICERS + TIMELINE + FILTERS     │
│  Region · Category · Segment ·       │
│  Ship Mode · Year · Returned         │
└──────────────────────────────────────┘
          │
          ▼
              📊 FINAL DASHBOARD
```

### Power Query transformation steps

| Query | Key applied steps |
|-------|-------------------|
| `RawOrders` | Promote headers → change types → trim text → set `Postal Code` to text → `Order Date`/`Ship Date` to date |
| `RawReturns` | Promote headers → remove duplicate Order IDs |
| `RawPeople` | Promote headers → rename `Person` to `Regional Manager` |
| `FactOrders` | Merge `RawOrders` with `RawPeople` on **Region** → merge with `RawReturns` on **Order ID** → expand columns → replace `null` with `"No"` in `Returned` |
| `DimProducts` | Reference `RawOrders` → keep Product ID / Name / Category / Sub-Category → **remove duplicates** (1,862 rows) |
| `DimCustomers` | Reference `RawOrders` → keep Customer ID / Name / Segment → remove duplicates (793 rows) |
| `DimRegion` | Reference `RawOrders` → keep Region / Country → merge Regional Manager → remove duplicates (4 rows) |
| `DimShipMode` | Reference `RawOrders` → keep Ship Mode → remove duplicates (4 rows) |
| `DimDate` / `DimShipDate` | Blank query → `List.Dates` → convert to table → add Year, Quarter, Month Number, Month, Weekday |

---

## 🗂️ Data Model

Star schema: one fact table, six dimensions, all one-to-many with single-direction filter flow.

```
                     ┌─────────────────────┐
                     │      DimDate        │  ← marked as Date Table
                     │─────────────────────│
                     │ Date (PK)           │
                     │ Year · Quarter      │
                     │ Month Number · Month│
                     │ Weekday             │
                     └──────────┬──────────┘
                                │ 1
                                │ *  (Order Date)
┌──────────────────┐   ┌────────┴──────────────┐   ┌──────────────────┐
│   DimProducts    │   │      FactOrders       │   │  DimCustomers    │
│──────────────────│ 1 │───────────────────────│ * │──────────────────│
│ Product ID (PK)  ├───┤ Order ID              ├───┤ Customer ID (PK) │
│ Product Name     │ * │ Order Date   (FK)     │ 1 │ Customer Name    │
│ Category         │   │ Ship Date    (FK)     │   │ Segment          │
│ Sub-Category     │   │ Ship Mode    (FK)     │   └──────────────────┘
└──────────────────┘   │ Customer ID  (FK)     │
                       │ Product ID   (FK)     │   ┌──────────────────┐
┌──────────────────┐ * │ Region       (FK)     │ * │   DimShipMode    │
│    DimRegion     ├───┤ City · State · Postal ├───┤──────────────────│
│──────────────────│ 1 │ Sales · Quantity      │ 1 │ Ship Mode (PK)   │
│ Region (PK)      │   │ Discount · Profit     │   └──────────────────┘
│ Country          │   │ Regional Manager      │
│ Regional Manager │   │ Returned              │   ┌──────────────────┐
└──────────────────┘   └───────────┬───────────┘ * │  DimShipDate     │
                                   └───────────────┤──────────────────│
                                                 1 │ Date (PK) + parts│
                                                   └──────────────────┘
```

| Table | Type | Rows | Grain | Key |
|-------|------|------|-------|-----|
| `FactOrders` | Fact | 9,994 | One row per order line item | Order ID + Product ID |
| `DimDate` | Dimension | 1,461 | One row per calendar date | Date *(marked as Date Table)* |
| `DimShipDate` | Dimension | 1,461 | One row per calendar date | Date *(role-playing date dim)* |
| `DimProducts` | Dimension | 1,862 | One row per product | Product ID |
| `DimCustomers` | Dimension | 793 | One row per customer | Customer ID |
| `DimRegion` | Dimension | 4 | One row per region | Region |
| `DimShipMode` | Dimension | 4 | One row per ship mode | Ship Mode |

> **Note on role-playing dimensions:** Excel allows only one active relationship between two tables. `DimShipDate` exists as a separate table so `Ship Date` can be analysed independently of `Order Date` without deactivating relationships.

---

## 🧮 DAX Measures

### Core KPIs

```dax
Total Sales =
SUM ( FactOrders[Sales] )
```

```dax
Total Profit =
SUM ( FactOrders[Profit] )
```

```dax
Total Quantity =
SUM ( FactOrders[Quantity] )
```

```dax
Profit Margin % =
DIVIDE ( [Total Profit], [Total Sales], 0 )
```

```dax
Total Orders =
DISTINCTCOUNT ( FactOrders[Order ID] )
```

```dax
Total Customers =
DISTINCTCOUNT ( FactOrders[Customer ID] )
```

```dax
Total Products Sold =
DISTINCTCOUNT ( FactOrders[Product ID] )
```

```dax
Average Order Value =
DIVIDE ( [Total Sales], [Total Orders], 0 )
```

```dax
Average Discount =
AVERAGE ( FactOrders[Discount] )
```

```dax
Sales per Customer =
DIVIDE ( [Total Sales], [Total Customers], 0 )
```

### Time Intelligence

```dax
Sales YTD =
TOTALYTD ( [Total Sales], DimDate[Date] )
```

```dax
Profit YTD =
TOTALYTD ( [Total Profit], DimDate[Date] )
```

```dax
Sales LY =
CALCULATE ( [Total Sales], SAMEPERIODLASTYEAR ( DimDate[Date] ) )
```

```dax
Sales YoY % =
DIVIDE ( [Total Sales] - [Sales LY], [Sales LY], 0 )
```

```dax
Profit LY =
CALCULATE ( [Total Profit], SAMEPERIODLASTYEAR ( DimDate[Date] ) )
```

```dax
Profit YoY % =
DIVIDE ( [Total Profit] - [Profit LY], [Profit LY], 0 )
```

```dax
Sales PM =
CALCULATE ( [Total Sales], DATEADD ( DimDate[Date], -1, MONTH ) )
```

```dax
Sales MoM % =
DIVIDE ( [Total Sales] - [Sales PM], [Sales PM], 0 )
```

```dax
Rolling 3M Sales =
CALCULATE (
    [Total Sales],
    DATESINPERIOD ( DimDate[Date], MAX ( DimDate[Date] ), -3, MONTH )
)
```

### Returns & Shipping

```dax
Returned Orders =
CALCULATE (
    DISTINCTCOUNT ( FactOrders[Order ID] ),
    FactOrders[Returned] = "Yes"
)
```

```dax
Return Rate % =
DIVIDE ( [Returned Orders], [Total Orders], 0 )
```

```dax
Returned Sales Value =
CALCULATE ( [Total Sales], FactOrders[Returned] = "Yes" )
```

```dax
Avg Shipping Days =
AVERAGEX (
    FactOrders,
    DATEDIFF ( FactOrders[Order Date], FactOrders[Ship Date], DAY )
)
```

### Ranking & Contribution

```dax
% of Total Sales =
DIVIDE (
    [Total Sales],
    CALCULATE ( [Total Sales], ALL ( FactOrders ) ),
    0
)
```

```dax
Product Rank =
RANKX ( ALL ( DimProducts[Product Name] ), [Total Sales],, DESC, DENSE )
```

```dax
Top 10 Customer Sales =
IF (
    RANKX ( ALL ( DimCustomers[Customer Name] ), [Total Sales],, DESC ) <= 10,
    [Total Sales],
    BLANK ()
)
```

```dax
Loss Making Sales =
CALCULATE ( [Total Sales], FILTER ( FactOrders, FactOrders[Profit] < 0 ) )
```

```dax
Loss Making Orders =
CALCULATE (
    DISTINCTCOUNT ( FactOrders[Order ID] ),
    FILTER ( FactOrders, FactOrders[Profit] < 0 )
)
```

---

## ✨ Dashboard Features

| Feature | Implementation | Purpose |
|---------|----------------|---------|
| **Power Query ETL** | 10 queries with full applied-steps history, including 2 merge joins | One-click refresh, zero manual cleaning |
| **Power Pivot Model** | Star schema, 7 tables, 6 relationships, `DimDate` marked as Date Table | Removes VLOOKUP dependency, enables time intelligence |
| **DAX Measures** | 25+ explicit measures across KPI, time intelligence, returns, and ranking | Reusable, filter-aware business logic |
| **KPI Cards** | Sales, Profit, Margin %, Orders, AOV, Return Rate | Headline numbers at a glance |
| **Pivot Charts** | Clustered column (region), bar (sub-category profit), line (monthly trend), combo (sales vs margin) | Fully dynamic, respond to every filter |
| **Standard Charts** | Scatter of Discount vs Profit; sparklines on the trend table | Shows the discount–loss relationship |
| **Slicers** | Region · Category · Sub-Category · Segment · Ship Mode · Year · Returned | Point-and-click cross-filtering |
| **Timeline** | Date timeline on `DimDate` | Filter by day / month / quarter / year |
| **Report Connections** | Every slicer wired to all pivot tables | One click filters the whole dashboard |
| **Filters** | Top-10 value filters, label filters, and profit-threshold filters | Isolate top performers and loss-makers |
| **Conditional Formatting** | Red/green data bars on profit; colour scales on the state table | Instantly spots negative margin |
| **Clean UI** | Gridlines hidden, layout locked, consistent palette, shapes as KPI cards | Presentation-ready |

---

## 🖼️ Dashboard Screenshot

![Dashboard Overview](images/dashboard_overview.png)

<details>
<summary>📷 More screenshots</summary>

**Power Query — Applied Steps & Merge**
![Power Query](images/power_query.png)

**Power Pivot — Diagram View (Star Schema)**
![Data Model](images/data_model.png)

**DAX Measure List**
![DAX Measures](images/dax_measures.png)

</details>

> Create an `images/` folder in the repo, drop your PNGs in, and the paths above will resolve.

---

## 🔍 Key Findings

*(All figures computed from the dataset in this repository.)*

### Headline numbers

| KPI | Value |
|-----|-------|
| Total Sales | **$2,297,201** |
| Total Profit | **$286,397** |
| Profit Margin | **12.47%** |
| Total Orders | **5,009** |
| Units Sold | **37,873** |
| Customers | **793** |
| Products | **1,862** |
| Average Order Value | **$458.61** |
| Returned Orders | **296 (5.9%)** — $180,504 in sales |
| Avg Shipping Time | **3.96 days** |

### 1. Furniture sells well but barely makes money

| Category | Sales | Profit | Margin |
|----------|-------|--------|--------|
| Technology | $836,154 | $145,455 | **17.4%** |
| Office Supplies | $719,047 | $122,491 | **17.0%** |
| Furniture | $741,999 | $18,451 | **2.5%** |

Furniture generates **32% of revenue but only 6% of profit**. Its margin is roughly one-seventh of the other two categories.

### 2. Three sub-categories lose money outright

| Sub-Category | Sales | Profit |
|--------------|-------|--------|
| Tables | $206,966 | **−$17,725** |
| Bookcases | $114,880 | **−$3,473** |
| Supplies | $46,674 | **−$1,189** |

Tables alone wipe out **$17.7K** of profit on $207K of revenue. Eliminating or repricing these three would lift total profit by roughly **8%**.

Best performers: **Copiers ($55,618 profit)**, Phones ($44,516), Accessories ($41,937), Paper ($34,054), Binders ($30,222).

### 3. Discounting above 20% is always loss-making

| Discount Band | Sales | Profit |
|---------------|-------|--------|
| 0% | $1,087,908 | **+$320,988** |
| 1–20% | $846,522 | **+$100,785** |
| 21–30% | $103,227 | **−$10,369** |
| 31–50% | $195,315 | **−$48,448** |
| >50% | $64,229 | **−$76,559** |

This is the single clearest finding in the dataset. Every discount band above 20% is **net negative**, together destroying **$135,376** of profit. A hard discount ceiling at 20% would be the highest-impact policy change available.

### 4. Central region underperforms badly

| Region | Sales | Profit | Margin |
|--------|-------|--------|--------|
| West | $725,458 | $108,418 | **14.9%** |
| East | $678,781 | $91,523 | **13.5%** |
| South | $391,722 | $46,749 | **11.9%** |
| Central | $501,240 | $39,706 | **7.9%** |

Central earns **28% more revenue than South but 15% less profit** — a margin problem, not a volume problem.

### 5. Ten states are net loss-makers

10 of 49 states post negative profit. The worst: **Texas (−$25,729)**, **Ohio (−$16,971)**, **Pennsylvania (−$15,560)**, **Illinois (−$12,608)**, **North Carolina (−$7,491)**. Texas alone erases 9% of company profit.

### 6. Growth is accelerating

| Year | Sales | Profit | Sales YoY |
|------|-------|--------|-----------|
| 2014 | $484,247 | $49,544 | — |
| 2015 | $470,533 | $61,619 | −2.8% |
| 2016 | $609,206 | $81,795 | +29.5% |
| 2017 | $733,215 | $93,439 | **+20.4%** |

Sales grew **51% from 2014 to 2017**, and profit grew **89%** — margin is improving even without policy changes.

### 7. Strong Q4 seasonality

November ($352K), December ($325K), and September ($308K) are the peak months. **February is the weakest at $59.8K** — under one-sixth of November. Inventory and staffing should follow this cycle.

### 8. Consumer segment dominates, no concentration risk

Consumer drives **$1.16M (51%)** of sales, Corporate $706K, Home Office $430K. The **top 10 customers account for only 6.7%** of revenue, so there is no meaningful customer-concentration risk.

---

### 💡 Recommendations

1. **Cap discounts at 20%.** Every band above it is loss-making; this alone addresses $135K of destroyed profit.
2. **Restructure the Furniture category.** Reprice or discontinue Tables and Bookcases, or renegotiate supplier cost.
3. **Audit the Central region** and the ten loss-making states — Texas first — for pricing and discount discipline.
4. **Double down on Technology**, particularly Copiers and Phones, which carry the strongest margins.
5. **Plan inventory around Q4**, and use the February trough for promotions on full-margin products only.
6. **Investigate the 5.9% return rate** on $180K of sales; identify which sub-categories drive it.

---

## 📁 Repository Structure

```
Superstore-Sales-Dashboard/
│
├── README.md                                  # Project documentation (this file)
│
├── data/
│   ├── raw/
│   │   ├── Sample_Superstore_Orders.csv       # Original order-line data
│   │   ├── Sample_Superstore_Returns.csv      # Returned order IDs
│   │   └── Sample_Superstore_People.csv       # Regional managers
│   └── processed/
│       └── Superstore_PowerQuery_DataModel.xlsx   # Cleaned queries + model
│
├── dashboard/
│   └── Superstore_Sales_Dashboard.xlsx        # Main workbook (PQ + PP + DAX + charts)
│
├── documentation/
│   ├── data_dictionary.md                     # Column definitions & data types
│   ├── dax_measures.md                        # Full DAX reference
│   └── power_query_steps.md                   # Documented M transformation steps
│
├── images/
│   ├── dashboard_overview.png                 # Main dashboard screenshot
│   ├── power_query.png                        # Power Query editor
│   ├── data_model.png                         # Power Pivot diagram view
│   └── dax_measures.png                       # Measure list
│
└── LICENSE
```

---

## 🛠️ Tools & Technologies

| Tool | Used For |
|------|----------|
| **Microsoft Excel** | Primary platform |
| **Power Query (M)** | Extraction, cleaning, merging, date-table generation |
| **Power Pivot** | Star-schema data modelling and relationships |
| **DAX** | Calculated measures and time intelligence |
| **Pivot Tables & Pivot Charts** | Aggregation and dynamic visualisation |
| **Charts** | Scatter and sparkline supporting visuals |
| **Slicers & Timeline** | Interactive cross-filtering |

---

## 🚀 How to Use

1. Clone or download this repository.
2. Open `dashboard/Superstore_Sales_Dashboard.xlsx` in **Excel 2016 or later** (Power Pivot required).
3. Enable content if prompted.
4. Go to `Data → Refresh All` to reload every query.
5. Use the slicers and timeline on the dashboard sheet to explore.

> **Note:** Power Pivot is available in Excel 2013+ (Professional Plus), Excel 2016+ (most editions), and Microsoft 365.

---

## 📊 Dataset

Sample Superstore — a widely used retail dataset of US office-supply, furniture, and technology orders.

- **Period:** 3 January 2014 – 30 December 2017
- **Records:** 9,994 order lines across 5,009 orders
- **Fields:** 21 original columns, extended to 22 with `Regional Manager` and `Returned`

---

## 👤 Author

**[Vismit Vikas Shrisunder]**
[LinkedIn](#) · [Portfolio](#) · [Email](#)

---

## 📄 License

This project is licensed under the MIT License — see the `LICENSE` file for details.

---

⭐ If you found this project useful, consider giving it a star.
