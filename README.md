# Bank Financing Performance Dashboard (Excel)

An interactive Excel dashboard for the financing department of **Horizon Islamic Bank** (a fictional bank in Kuwait).
A manager can filter by **branch, product, customer segment, channel and month**, and every KPI and chart updates instantly.

---

## Business Question
The financing manager needs one page to review every week:
- How much financing was approved, and how many deals?
- What is the approval rate, and where is it dropping?
- Which products, channels and segments drive the results?
- How long do customers wait for a decision?

## Dataset
- **1,329** raw financing applications (Jan – Sep 2026)
- **950** customers (segment, nationality, birth year)
- Products: Home Ijara, SME Murabaha, Personal Tawarruq, Auto Murabaha
- Channels: Branch, Mobile App, Online
- 6 branches: Kuwait City, Hawally, Salmiya, Farwaniya, Jahra, Ahmadi

## Process

### 1. Data Preparation
- Removed **24** duplicate records
- Standardised product and channel names (e.g. `sme murabaha`, `MobileApp` → consistent values) with TRIM, PROPER and Find & Replace
- Converted text amounts such as `"12,500 KWD"` into numbers
- Moved **12** applications with missing amounts to a separate `Rejected_Data` sheet
- Converted the clean data (**1,293** rows) into an Excel Table named `Deals`

### 2. Data Enrichment (calculated columns)
| Column | Logic |
|---|---|
| Segment, Nationality | `XLOOKUP` from the Customers sheet |
| AgeGroup | `IFS` → Under 30 / 30-44 / 45+ |
| DaysToDecision | DecisionDate − ApplicationDate (blank if Pending) |
| IsApproved | 1 if Approved, else 0 |
| ApprovedAmount | Amount if Approved, else 0 |
| ExpectedAnnualProfit | ApprovedAmount × ProfitRate |

### 3. Calculation Layer (`calc` sheet)
Five PivotTables built on one shared Pivot Cache:
- KPIs by status (deals, average days, approved amount)
- Approved amount by month (dates grouped by month)
- Approved amount by product (sorted largest first)
- Approval rate by channel (% of row total, Pending excluded)
- Approved amount by segment

KPI cards read from the pivots with `GETPIVOTDATA`, so they react to every filter.

### 4. Dashboard
- **5 KPI cards:** Total Approved Amount, Approved Deals, Approval Rate, Avg Approved Deal Size, Avg Days to Decision
- **4 PivotCharts:** Column (product), 100% stacked bar (channel), Line (month), Doughnut (segment)
- **4 Slicers + 1 Timeline** connected to all pivots through Report Connections
- Dark theme, custom slicer and timeline styles, field buttons hidden

## Key Insights
| Area | Finding |
|---|---|
| Overall | **34.8M KWD** approved across **890** deals, approval rate **71.8%** |
| Jahra branch | Approval rate fell from **72.8%** (Jan–May) to **43.1%** (Jun–Sep) |
| Mobile App | Lowest approval rate: **64.5%** vs Branch **75.0%** |
| Home Ijara | **20.2 days** to decision vs **5.7 days** for other products |
| Auto Murabaha | Peaks in **August (581,300 KWD)** and **July (578,000 KWD)** |
| Premium segment | **1,578,800 KWD** approved in Q3 (Jul–Sep) |

## Recommendations
1. **Jahra branch:** review what changed in June (staff, credit officer, local policy), audit a sample of rejected applications, and track the branch's approval rate weekly.
2. **Mobile App:** add eligibility checks inside the app before submission (minimum salary, required documents, debt ratio) to reduce rejections.
3. **Home Ijara:** target a decision time of 10 days or less with a dedicated team, pre-approved property valuers and a document checklist at application.

## How to Use
1. Download `Bank_Financing_Dashboard.xlsx`
2. Open the **Dashboard** sheet (Excel 2016 or later, Microsoft 365 recommended)
3. Use the slicers and the timeline to filter. Clear filters with the 🧹 button on each slicer.

## Tools
Microsoft Excel: Tables, XLOOKUP, IFS, PivotTables, Date Grouping, GETPIVOTDATA, PivotCharts, Slicers, Timeline, Report Connections

## Repository Structure
```
├── Bank_Financing_Dashboard.xlsx   # Full workbook (data, calc, dashboard, answers)
├── dashboard.png                   # Dashboard screenshot
└── README.md
```

## Author
**Abdulrahman Lehlouh**: Data Science & AI graduate, Kuwait
