# 🧮 Formulas & Examples - Practical Guide

> **Real-world formula patterns and examples for common Excel tasks**

## 📋 Table of Contents

- [Formula Basics](#formula-basics)
- [Common Formula Patterns](#common-formula-patterns)
- [Nested Formulas](#nested-formulas)
- [Array Formulas](#array-formulas)
- [Financial Formulas](#financial-formulas)
- [Date Calculations](#date-calculations)
- [Text Manipulation](#text-manipulation)
- [Lookup Formulas](#lookup-formulas)
- [Conditional Calculations](#conditional-calculations)
- [Advanced Techniques](#advanced-techniques)

---

## Formula Basics

### Formula Structure

**All formulas start with `=`**
```excel
=SUM(A1:A10)
=A1+B1
=IF(C1>100, "High", "Low")
```

### Cell References

**Relative Reference:**
```excel
=A1+B1
// Adjusts when copied: A2+B2, A3+B3, etc.
```

**Absolute Reference:**
```excel
=$A$1+B1
// $A$1 stays fixed when copied
```

**Mixed Reference:**
```excel
=$A1+B$1
// Column A fixed, Row 1 fixed
```

**Toggle Reference Type:**
- Press **F4** (Windows) or **⌘+T** (Mac) while editing formula

### Order of Operations (PEMDAS)

1. **Parentheses** ()
2. **Exponents** ^
3. **Multiplication** * and **Division** /
4. **Addition** + and **Subtraction** -

**Example:**
```excel
=2+3*4              → 14 (multiply first)
=(2+3)*4            → 20 (parentheses first)
```

---

## Common Formula Patterns

### Basic Calculations

**Total:**
```excel
=SUM(A1:A10)                    // Sum range
=A1+A2+A3                       // Add specific cells
```

**Average:**
```excel
=AVERAGE(A1:A10)                // Mean
=SUM(A1:A10)/COUNT(A1:A10)      // Alternative
```

**Percentage:**
```excel
=Part/Whole                     // Decimal (0.15)
=(Part/Whole)*100               // Percentage (15)
=Part/Whole                     // Format as % with Ctrl+Shift+%
```

**Percentage Change:**
```excel
=(New-Old)/Old                  // Growth rate
=(B2-A2)/A2                     // Example
```

**Markup/Margin:**
```excel
// Markup (cost to price)
=Cost*(1+MarkupRate)
=100*(1+0.2)                    → $120 (20% markup)

// Margin (revenue to profit)
=(Price-Cost)/Price
=(120-100)/120                  → 16.67%
```

---

### Running Calculations

**Running Total:**
```excel
// In B2, copy down
=SUM($A$2:A2)
// Range expands: $A$2:A2, $A$2:A3, $A$2:A4...
```

**Running Average:**
```excel
=AVERAGE($A$2:A2)
```

**Cumulative Percentage:**
```excel
=SUM($A$2:A2)/SUM($A$2:$A$100)
```

**Running Count:**
```excel
=COUNTA($A$2:A2)
```

---

### Ranking & Sorting

**Rank Values:**
```excel
=RANK.EQ(A2, $A$2:$A$100, 0)    // 0 = descending (1 is highest)
=RANK.EQ(A2, $A$2:$A$100, 1)    // 1 = ascending
```

**Top N:**
```excel
=LARGE(A:A, 1)                  // Largest
=LARGE(A:A, 10)                 // 10th largest
=SMALL(A:A, 5)                  // 5th smallest
```

**Percentile:**
```excel
=PERCENTILE.INC(A:A, 0.75)      // 75th percentile
=PERCENTILE.INC(A:A, 0.9)       // 90th percentile
```

---

## Nested Formulas

### IF Statements

**Simple IF:**
```excel
=IF(A1>100, "High", "Low")
```

**Nested IF (Multiple Conditions):**
```excel
=IF(A1>=90, "A",
    IF(A1>=80, "B",
        IF(A1>=70, "C",
            IF(A1>=60, "D", "F"))))
```

**IFS (Excel 2019+, Cleaner):**
```excel
=IFS(A1>=90, "A",
     A1>=80, "B",
     A1>=70, "C",
     A1>=60, "D",
     TRUE, "F")
```

**IF with AND:**
```excel
=IF(AND(A1>50, B1<100), "Valid", "Invalid")
// Both conditions must be true
```

**IF with OR:**
```excel
=IF(OR(A1="Red", A1="Blue", A1="Green"), "Primary", "Other")
// Any condition can be true
```

---

### Combining Functions

**SUMIF with Multiple Criteria (Alternative to SUMIFS):**
```excel
=SUMIF(A:A, "Apple", B:B) + SUMIF(A:A, "Orange", B:B)
```

**Average Excluding Zeros:**
```excel
=AVERAGEIF(A:A, "<>0")
```

**Count Unique Values:**
```excel
=SUMPRODUCT(1/COUNTIF(A1:A100, A1:A100))
```

**Conditional Average:**
```excel
=AVERAGEIFS(Sales, Region, "North", Month, "Jan")
```

---

## Array Formulas

### What Are Array Formulas?

**Pre-Excel 365:**
- Enter with **Ctrl+Shift+Enter**
- Shows {curly braces} in formula bar
- Single formula processes array

**Excel 365:**
- Dynamic arrays (auto-spill)
- No Ctrl+Shift+Enter needed
- Results spill to multiple cells

### Classic Array Examples

**Sum of Products:**
```excel
=SUM(A1:A10*B1:B10)
// Multiplies each pair, then sums
// Pre-365: Ctrl+Shift+Enter
```

**Multi-Criteria Count:**
```excel
=SUM((A1:A100="Apple")*(B1:B100>100))
// Counts rows where both conditions true
```

**Unique Count with Criteria:**
```excel
=SUM(IF(COUNTIFS(A:A, A:A, B:B, "North")=1, 1, 0))
```

### Dynamic Array Functions (Excel 365)

**FILTER:**
```excel
=FILTER(A1:C100, B1:B100>100)
// Returns all rows where B>100
```

**SORT:**
```excel
=SORT(A1:C100, 2, -1)
// Sort by column 2, descending
```

**UNIQUE:**
```excel
=UNIQUE(A1:A100)
// Unique values from A
```

**SEQUENCE:**
```excel
=SEQUENCE(10, 1, 1, 1)
// Numbers 1-10
```

**SORTBY:**
```excel
=SORTBY(A1:B100, B1:B100, -1)
// Sort A:B by column B descending
```

---

## Financial Formulas

### Loan Calculations

**Monthly Payment:**
```excel
=PMT(rate/12, months, -loan_amount)
=PMT(5%/12, 360, -200000)
// $1,073.64 monthly payment
// 5% annual rate, 30 years, $200k loan
```

**Total Interest Paid:**
```excel
=(PMT(rate/12, months, -loan)*months)-loan
=($1,073.64*360)-200000
// $186,511.57 total interest
```

**Principal and Interest Breakdown:**
```excel
=PPMT(rate/12, period, months, -loan)   // Principal
=IPMT(rate/12, period, months, -loan)   // Interest
```

### Investment Calculations

**Future Value:**
```excel
=FV(rate/12, months, -monthly_payment, 0)
=FV(6%/12, 360, -500, 0)
// $502,257.91 (save $500/mo for 30 years at 6%)
```

**Present Value:**
```excel
=PV(rate, periods, payment)
=PV(5%, 10, -1000)
// $7,721.73 needed today to receive $1000/year for 10 years
```

**Net Present Value:**
```excel
=NPV(discount_rate, value1, value2, ...)
=NPV(10%, B2:B11)-A1
// NPV of cash flows minus initial investment
```

**Internal Rate of Return:**
```excel
=IRR(values, [guess])
=IRR(A1:A11, 0.1)
// Annual return rate
```

---

## Date Calculations

### Age & Tenure

**Age in Years:**
```excel
=DATEDIF(Birthdate, TODAY(), "Y")
=DATEDIF(A2, TODAY(), "Y")
```

**Age in Years and Months:**
```excel
=DATEDIF(A2, TODAY(), "Y") & " years, " & DATEDIF(A2, TODAY(), "YM") & " months"
```

**Days Between Dates:**
```excel
=EndDate-StartDate
=B2-A2
```

**Business Days Between:**
```excel
=NETWORKDAYS(StartDate, EndDate, [Holidays])
=NETWORKDAYS(A2, B2, Holidays!A:A)
```

### Date Manipulation

**First Day of Month:**
```excel
=DATE(YEAR(A2), MONTH(A2), 1)
=EOMONTH(A2, -1) + 1
```

**Last Day of Month:**
```excel
=EOMONTH(A2, 0)
```

**Add Months:**
```excel
=EDATE(A2, 3)
// 3 months from A2
```

**Add Business Days:**
```excel
=WORKDAY(A2, 10)
// 10 business days from A2
```

**Extract Date Parts:**
```excel
=YEAR(A2)                       // 2025
=MONTH(A2)                      // 12
=DAY(A2)                        // 14
=TEXT(A2, "MMMM")              // "December"
=TEXT(A2, "DDDD")              // "Saturday"
```

**Quarter:**
```excel
="Q" & ROUNDUP(MONTH(A2)/3, 0)
```

**Week Number:**
```excel
=WEEKNUM(A2)
```

### Date Validation

**Is Weekend:**
```excel
=WEEKDAY(A2, 2) > 5
=OR(WEEKDAY(A2)=1, WEEKDAY(A2)=7)
```

**Is Overdue:**
```excel
=AND(A2<TODAY(), B2<>"Complete")
```

**Days Until:**
```excel
=A2-TODAY()
=IF(A2<TODAY(), "Overdue", A2-TODAY() & " days")
```

---

## Text Manipulation

### Combining Text

**Concatenate:**
```excel
=A2 & " " & B2
=CONCAT(A2, " ", B2)
=TEXTJOIN(", ", TRUE, A2:C2)
```

**Full Name:**
```excel
=PROPER(A2 & " " & B2)
// "john smith" → "John Smith"
```

**Full Address:**
```excel
=TEXTJOIN(", ", TRUE, Street, City, State, ZIP)
```

### Extracting Text

**First Name (from "Last, First"):**
```excel
=TRIM(RIGHT(A2, LEN(A2)-FIND(",", A2)))
```

**Last Name:**
```excel
=TRIM(LEFT(A2, FIND(",", A2)-1))
```

**Email Username:**
```excel
=LEFT(A2, FIND("@", A2)-1)
// "user@email.com" → "user"
```

**Email Domain:**
```excel
=MID(A2, FIND("@", A2)+1, LEN(A2))
// "user@email.com" → "email.com"
```

**Extract Numbers from Text:**
```excel
=SUMPRODUCT(MID(0&A2, LARGE(INDEX(ISNUMBER(--MID(A2, ROW($1:$25), 1))*ROW($1:$25), 0), ROW($1:$25))+1, 1)*10^ROW($1:$25)/10)
```

### Cleaning Text

**Remove Extra Spaces:**
```excel
=TRIM(A2)
```

**Clean All:**
```excel
=TRIM(CLEAN(A2))
```

**Remove Non-Numeric:**
```excel
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2, "-", ""), "(", ""), ")", "")
// Clean phone number
```

**Proper Case Fix:**
```excel
=PROPER(LOWER(A2))
// Better than PROPER alone for all-caps
```

---

## Lookup Formulas

### VLOOKUP Patterns

**Basic Lookup:**
```excel
=VLOOKUP(A2, PriceTable, 2, FALSE)
```

**With Error Handling:**
```excel
=IFERROR(VLOOKUP(A2, Table, 2, FALSE), "Not Found")
=IFNA(VLOOKUP(A2, Table, 2, FALSE), 0)
```

**Approximate Match (Price Tiers):**
```excel
=VLOOKUP(A2, TierTable, 2, TRUE)
// Table must be sorted ascending
```

### INDEX/MATCH (Better Than VLOOKUP)

**Basic:**
```excel
=INDEX(ReturnColumn, MATCH(LookupValue, LookupColumn, 0))
=INDEX(C:C, MATCH(A2, B:B, 0))
```

**Two-Way Lookup:**
```excel
=INDEX(DataRange, MATCH(RowValue, RowHeaders, 0), MATCH(ColValue, ColHeaders, 0))
=INDEX(B2:M13, MATCH(A15, A2:A13, 0), MATCH(B14, B1:M1, 0))
```

**Lookup Left:**
```excel
=INDEX(A:A, MATCH(E2, C:C, 0))
// Lookup in C, return from A
```

### XLOOKUP (Excel 365)

**Basic:**
```excel
=XLOOKUP(A2, LookupArray, ReturnArray, "Not Found")
```

**Multiple Columns:**
```excel
=XLOOKUP(A2, IDs, B1:E100)
// Returns entire row
```

**Last Occurrence:**
```excel
=XLOOKUP(A2, Range1, Range2, , 0, -1)
```

---

## Conditional Calculations

### SUMIF/SUMIFS

**Single Criteria:**
```excel
=SUMIF(A:A, "Apple", B:B)
=SUMIF(B:B, ">100", B:B)
```

**Multiple Criteria:**
```excel
=SUMIFS(Sales, Region, "North", Product, "Apple", Month, "Jan")
```

**OR Logic (Multiple SUMIF):**
```excel
=SUMIF(A:A, "Apple", B:B) + SUMIF(A:A, "Orange", B:B)
```

**Sum with Wildcards:**
```excel
=SUMIF(A:A, "*apple*", B:B)
// Contains "apple"
```

### COUNTIF/COUNTIFS

**Count Occurrences:**
```excel
=COUNTIF(A:A, "Complete")
```

**Count Greater Than:**
```excel
=COUNTIF(B:B, ">100")
```

**Count Between:**
```excel
=COUNTIFS(A:A, ">=50", A:A, "<=100")
```

**Count Non-Blank:**
```excel
=COUNTA(A:A)
```

**Count Unique:**
```excel
=SUMPRODUCT(1/COUNTIF(A1:A100, A1:A100))
```

### AVERAGEIF/AVERAGEIFS

**Conditional Average:**
```excel
=AVERAGEIF(A:A, "Apple", B:B)
=AVERAGEIFS(Sales, Region, "North", Year, 2025)
```

**Average Excluding Zero:**
```excel
=AVERAGEIF(A:A, "<>0")
```

---

## Advanced Techniques

### Conditional Formatting with Formulas

**Highlight Row Based on Status:**
```excel
// Apply to A2:E100
=$E2="Complete"
```

**Alternate Row Colors:**
```excel
=MOD(ROW(), 2)=0
```

**Highlight Duplicates:**
```excel
=COUNTIF($A$2:$A$100, A2)>1
```

**Weekend Highlighting:**
```excel
=WEEKDAY(A2, 2)>5
```

### Data Validation Formulas

**No Duplicates:**
```excel
=COUNTIF($A$1:$A$100, A1)=1
```

**Must Be Greater Than Another Cell:**
```excel
=A2>B2
```

**Email Validation:**
```excel
=AND(LEN(A2)>0, ISNUMBER(FIND("@", A2)), ISNUMBER(FIND(".", A2)))
```

### Dynamic Named Ranges

**Expand with Data:**
```excel
=OFFSET(Sheet1!$A$1, 0, 0, COUNTA(Sheet1!$A:$A), 1)
```

**Table Column:**
```excel
=Table1[ColumnName]
```

---

## Quick Reference

### Most Used Formulas

```excel
// Basic Math
=SUM(A1:A10)
=AVERAGE(A1:A10)
=COUNT(A1:A10)

// Conditional
=SUMIF(A:A, "Criteria", B:B)
=COUNTIF(A:A, ">100")
=AVERAGEIF(A:A, "Apple", B:B)

// Lookup
=VLOOKUP(A2, Table, 2, FALSE)
=INDEX(C:C, MATCH(A2, B:B, 0))
=XLOOKUP(A2, Lookup, Return)

// Logical
=IF(A1>100, "High", "Low")
=IFS(A1>90,"A", A1>80,"B", TRUE,"C")

// Text
=CONCAT(A2, " ", B2)
=LEFT(A2, 5)
=TRIM(A2)

// Date
=TODAY()
=DATEDIF(A2, TODAY(), "Y")
=EOMONTH(A2, 0)
```

---

**[⬆ Back to Main README](../../README.md)**
