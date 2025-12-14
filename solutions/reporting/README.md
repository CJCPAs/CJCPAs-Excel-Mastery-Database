# Reporting & Dashboard Solutions

> **Create professional reports, dashboards, and dynamic displays**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Create dynamic titles | [Dynamic Text](#dynamic-titles-and-labels) |
| Show live dates | [Auto-Updating Dates](#auto-updating-dates) |
| Build a dashboard | [Dashboard Basics](#dashboard-fundamentals) |
| Create KPI indicators | [KPI Displays](#kpi-indicators) |
| Show progress bars in cells | [In-Cell Charts](#in-cell-progress-bars) |
| Highlight variances | [Variance Reporting](#variance-analysis) |
| Create ranking displays | [Top N Reports](#top-n-rankings) |
| Build summary tables | [Executive Summaries](#executive-summary-tables) |
| Add sparklines | [Trend Indicators](#sparklines-for-trends) |
| Create interactive reports | [Dynamic Filtering](#interactive-reports) |

---

## Dynamic Titles and Labels

### The Challenge
Create report titles that update automatically with dates, filters, or parameters.

### Quick Answer
```excel
="Sales Report - " & TEXT(TODAY(), "MMMM YYYY")
```

### Full Examples

**Auto-Updating Title with Current Month:**
```excel
="Monthly Report: " & TEXT(TODAY(), "MMMM YYYY")
```
**Result:** `Monthly Report: December 2025`

**Title Reflecting Filter Selection:**
```excel
="Sales Report - " & IF(A1="All", "All Regions", A1 & " Region")
```

**Date Range Title:**
```excel
="Report Period: " & TEXT(B1,"MM/DD/YY") & " to " & TEXT(B2,"MM/DD/YY")
```

**Quarter-Based Title:**
```excel
="Q" & ROUNDUP(MONTH(TODAY())/3,0) & " " & YEAR(TODAY()) & " Performance"
```
**Result:** `Q4 2025 Performance`

---

## Auto-Updating Dates

### Current Date/Time Displays

| Purpose | Formula | Result |
|---------|---------|--------|
| Today's date | `=TODAY()` | 12/14/2025 |
| Current time | `=NOW()` | 12/14/2025 3:45 PM |
| Formatted date | `=TEXT(TODAY(),"dddd, MMMM d, yyyy")` | Saturday, December 14, 2025 |
| Report timestamp | `="Generated: "&TEXT(NOW(),"MM/DD/YY h:mm AM/PM")` | Generated: 12/14/25 3:45 PM |

### Fiscal Periods
```excel
Fiscal Year (July start):  =IF(MONTH(TODAY())>=7, YEAR(TODAY()), YEAR(TODAY())-1)
Fiscal Quarter:            =CEILING((MONTH(TODAY())-6)/3, 1) (adjust for fiscal start)
```

### Period Comparisons
```excel
Last Month:      =EOMONTH(TODAY(),-1)
Same Month LY:   =EDATE(TODAY(),-12)
YTD Start:       =DATE(YEAR(TODAY()),1,1)
```

---

## Dashboard Fundamentals

### Dashboard Layout Best Practices

```
┌─────────────────────────────────────────────────────┐
│  TITLE / LOGO                    Date: [Dynamic]    │
├─────────────────────────────────────────────────────┤
│  ┌─────────┐ ┌─────────┐ ┌─────────┐ ┌─────────┐   │
│  │  KPI 1  │ │  KPI 2  │ │  KPI 3  │ │  KPI 4  │   │
│  │  $1.2M  │ │   85%   │ │  +12%   │ │  4,521  │   │
│  └─────────┘ └─────────┘ └─────────┘ └─────────┘   │
├─────────────────────────────────────────────────────┤
│  [Main Chart Area]              │ [Secondary Chart] │
│                                 │                   │
│                                 ├───────────────────┤
│                                 │ [Data Table]      │
│                                 │                   │
└─────────────────────────────────────────────────────┘
```

### Key Dashboard Elements

**1. KPI Cards**
```excel
Value:      =SUMIFS(Sales, Region, $B$1)
Label:      ="Total Sales"
Trend:      =IF(Current>Previous, "▲", "▼")
```

**2. Summary Metrics**
```excel
Total:      =SUM(Data[Sales])
Average:    =AVERAGE(Data[Sales])
Count:      =COUNTA(Data[Customer])
% of Goal:  =SUM(Data[Sales])/Goal
```

**3. Conditional Formatting for RAG Status**
- Green: >=100% of target
- Amber: 80-99% of target
- Red: <80% of target

---

## KPI Indicators

### Traffic Light Indicators

**Using Symbols:**
```excel
=IF(A1>=100,"🟢",IF(A1>=80,"🟡","🔴"))
```

**Using Wingdings (Format cell as Wingdings):**
```excel
=IF(A1>=Target,"n",IF(A1>=Target*0.8,"l","n"))
```
(n=check, l=circle in Wingdings)

**Using Unicode Shapes:**
```excel
=IF(A1>=Target,"●","○")  & " " & TEXT(A1,"0%")
```

### Trend Arrows
```excel
=IF(Current>Previous,"▲ ",IF(Current<Previous,"▼ ","◆ ")) & TEXT(ABS(Current-Previous)/Previous,"0.0%")
```
**Result:** `▲ 5.2%` or `▼ 3.1%`

### Star Ratings (1-5 Stars)
```excel
=REPT("★",ROUND(A1,0)) & REPT("☆",5-ROUND(A1,0))
```
**Result:** `★★★★☆`

### Progress Percentage with Bar
```excel
=REPT("█",A1*10) & REPT("░",10-A1*10) & " " & TEXT(A1,"0%")
```
**Result:** `████████░░ 80%`

---

## In-Cell Progress Bars

### Using REPT Function
```excel
=REPT("|",A1*20) & REPT(".",20-A1*20)
```

### Percentage Bar with Value
```excel
=REPT("▓",ROUND(A1*10,0)) & REPT("░",10-ROUND(A1*10,0)) & " " & TEXT(A1,"0%")
```

### Horizontal Bar Chart Effect
Use conditional formatting → Data Bars for automatic in-cell bars.

**Custom Data Bar Settings:**
- Minimum: 0
- Maximum: 100 (or formula)
- Fill: Solid or gradient
- Border: Optional

---

## Variance Analysis

### Basic Variance
```excel
Variance ($):    =Actual - Budget
Variance (%):    =(Actual - Budget) / Budget
Favorable/Not:   =IF(Actual>=Budget, "Favorable", "Unfavorable")
```

### Variance Display with Formatting
```excel
=IF(A1-B1>=0, "+" & TEXT(A1-B1,"$#,##0"), TEXT(A1-B1,"$#,##0"))
```

### Full Variance Table

| Category | Budget | Actual | Var $ | Var % | Status |
|----------|--------|--------|-------|-------|--------|
| Revenue | $100,000 | $108,000 | +$8,000 | +8.0% | ▲ |
| Expenses | $80,000 | $82,000 | -$2,000 | +2.5% | ▼ |
| Profit | $20,000 | $26,000 | +$6,000 | +30.0% | ▲ |

**Formulas:**
```excel
Var $:     =C2-B2
Var %:     =(C2-B2)/B2
Status:    =IF(C2>=B2,"▲","▼")
```

### Conditional Formatting Rules
- Positive variance: Green fill
- Negative variance: Red fill
- Near zero (±2%): Yellow fill

---

## Top N Rankings

### Top 5 Values
```excel
=LARGE(Data, ROW()-StartRow+1)
```

### Top 5 with Names (Dynamic Array - Excel 365)
```excel
=SORT(FILTER(Data, Data[Sales]>=LARGE(Data[Sales],5)), 2, -1)
```

### Top N Using INDEX/MATCH
```excel
Value:  =LARGE(B:B, ROW(A1))
Name:   =INDEX(A:A, MATCH(LARGE(B:B, ROW(A1)), B:B, 0))
```

### Ranking Table Example

| Rank | Salesperson | Sales | % of Total |
|------|-------------|-------|------------|
| 1 | Sarah | $125,000 | 25% |
| 2 | Mike | $98,000 | 20% |
| 3 | Lisa | $87,000 | 17% |

**Formulas:**
```excel
Rank:        =ROW()-1
Salesperson: =INDEX(Names, MATCH(LARGE(Sales,A2), Sales, 0))
Sales:       =LARGE(SalesRange, A2)
% of Total:  =C2/SUM(SalesRange)
```

---

## Executive Summary Tables

### Summary Statistics Block
```excel
| Metric | This Period | Last Period | Change |
|--------|-------------|-------------|--------|
| Revenue | =SUM(...) | =SUM(...) | =B2-C2 |
| Units | =COUNT(...) | =COUNT(...) | =B3-C3 |
| Avg Price | =AVERAGE(...) | =AVERAGE(...) | =B4-C4 |
```

### Period-Over-Period Comparison
```excel
Current Month:    =SUMIFS(Sales, Date, ">="&StartDate, Date, "<="&EndDate)
Prior Month:      =SUMIFS(Sales, Date, ">="&EDATE(StartDate,-1), Date, "<="&EDATE(EndDate,-1))
YoY:              =SUMIFS(Sales, Date, ">="&EDATE(StartDate,-12), Date, "<="&EDATE(EndDate,-12))
```

### Metric Cards Formula Set
```excel
Current Value:    =SUMIFS(...)
Prior Value:      =SUMIFS(...prior period...)
Change:           =Current-Prior
Change %:         =(Current-Prior)/Prior
Trend Arrow:      =IF(Current>Prior,"▲","▼")
```

---

## Sparklines for Trends

### Insert Sparklines
1. Select destination cell(s)
2. Insert → Sparklines → Line/Column/Win-Loss
3. Select data range
4. Format as needed

### Sparkline Types

| Type | Best For |
|------|----------|
| **Line** | Trends over time |
| **Column** | Comparing values |
| **Win/Loss** | Positive/negative results |

### Sparkline Options
- **High/Low Points:** Highlight min/max
- **First/Last Points:** Emphasize endpoints
- **Negative Points:** Different color for negatives
- **Axis:** Set min/max for comparison

### Mini Trend Indicator (No Sparklines)
```excel
=REPT("▁",B1) & REPT("▂",C1) & REPT("▄",D1) & REPT("▆",E1) & REPT("█",F1)
```
(Adjust based on relative values)

---

## Interactive Reports

### Using Data Validation for Filters
```excel
1. Create dropdown: Data → Data Validation → List
2. Source: =UniqueRegions (named range)
3. Reference in formulas: =SUMIFS(Sales, Region, $B$1)
```

### Dynamic Chart Titles
Link chart title to a cell containing:
```excel
=SelectedRegion & " Sales - " & TEXT(TODAY(),"MMMM YYYY")
```

### Slicer Controls
1. Insert → Slicer
2. Select fields to filter
3. Format and position
4. Works with Tables and PivotTables

### Show/Hide Sections
Use grouping (Alt+Shift+Right) to collapse sections:
```
▼ Detailed Data
    Row 1
    Row 2
▶ [Click to expand]
```

---

## Report Formatting Best Practices

### Number Formats
| Type | Format Code | Example |
|------|-------------|---------|
| Currency | `$#,##0` | $1,234 |
| Currency (neg red) | `$#,##0;[Red]-$#,##0` | -$500 |
| Percentage | `0.0%` | 85.5% |
| Large numbers | `$#,##0,,"M"` | $1.2M |
| Variance | `+0.0%;-0.0%` | +5.2% |

### Consistent Styling
- Headers: Bold, larger font
- Numbers: Right-aligned
- Labels: Left-aligned
- Totals: Bold with top border
- Sections: Alternating shading

### Print Optimization
- Set print area: Page Layout → Print Area
- Repeat headers: Page Layout → Print Titles
- Fit to page: Scale to Fit options
- Page breaks: View → Page Break Preview

---

## Common Reporting Formulas

### Period Calculations
```excel
MTD:     =SUMIFS(Sales, Date, ">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1), Date, "<="&TODAY())
QTD:     =SUMIFS(Sales, Date, ">="&EOMONTH(TODAY(),-MOD(MONTH(TODAY())-1,3))+1, Date, "<="&TODAY())
YTD:     =SUMIFS(Sales, Date, ">="&DATE(YEAR(TODAY()),1,1), Date, "<="&TODAY())
```

### Running Totals
```excel
=SUM($B$2:B2)  → Copy down for cumulative total
```

### Percent of Total
```excel
=B2/SUM($B$2:$B$100)
```

---

## Related Solutions

- [Conditional Calculations](../conditional-calculations/README.md) - SUMIFS, COUNTIFS
- [Data Analysis](../data-analysis/README.md) - Statistical analysis
- [Lookups](../lookups/README.md) - Dynamic data retrieval

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
