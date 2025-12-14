# 🎯 Goal-Based Solutions Library

> **Find solutions by what you want to achieve - 50+ common Excel goals with step-by-step instructions**

## 📋 Table of Contents

- [Looking Up & Finding Data](#looking-up--finding-data)
- [Cleaning & Organizing Data](#cleaning--organizing-data)
- [Calculating & Analyzing](#calculating--analyzing)
- [Text Manipulation](#text-manipulation)
- [Date & Time Operations](#date--time-operations)
- [Conditional Operations](#conditional-operations)
- [List Management](#list-management)
- [Visualization & Formatting](#visualization--formatting)

---

## Looking Up & Finding Data

### "I want to... look up a value from another table"

**Solution: XLOOKUP (Excel 365) or VLOOKUP**

**Scenario:** Find price for a product code

**XLOOKUP (Recommended):**
```excel
=XLOOKUP(A2, ProductCodes, Prices, "Not Found")
```

**VLOOKUP (Older Excel):**
```excel
=IFERROR(VLOOKUP(A2, ProductTable, 2, FALSE), "Not Found")
```

**Step-by-Step:**
1. Identify lookup value (e.g., product code in A2)
2. Identify lookup range (where to search)
3. Identify return range (what to return)
4. Use XLOOKUP or VLOOKUP with FALSE for exact match
5. Add error handling with IFERROR or built-in parameter

---

### "I want to... find values that match multiple criteria"

**Solution: XLOOKUP with multiple criteria or INDEX/MATCH**

**Scenario:** Find price for specific product in specific region

**Method 1 - XLOOKUP (Excel 365):**
```excel
=XLOOKUP(1, (Products=A2)*(Regions=B2), Prices)
```

**Method 2 - INDEX/MATCH:**
```excel
=INDEX(Prices, MATCH(1, (Products=A2)*(Regions=B2), 0))
```
*Enter as array formula (Ctrl+Shift+Enter in older Excel)*

**Method 3 - Helper Column:**
1. Create helper column: `=A2&"-"&B2`
2. Use VLOOKUP: `=VLOOKUP(A2&"-"&B2, HelperTable, 3, FALSE)`

---

### "I want to... look up and return multiple values"

**Solution: FILTER (Excel 365)**

**Scenario:** Get all sales for a specific salesperson

```excel
=FILTER(SalesData, Salesperson=A2, "No sales found")
```

**Returns:** All matching rows

**Without FILTER (older Excel):**
Use Advanced Filter or Power Query

---

### "I want to... find the last occurrence of a value"

**Solution: XLOOKUP with reverse search or INDEX/MATCH**

**XLOOKUP (Excel 365):**
```excel
=XLOOKUP(A2, LookupRange, ReturnRange, , 0, -1)
```
The `-1` at end searches from bottom up

**INDEX/MATCH:**
```excel
=INDEX(ReturnRange, MATCH(2, 1/(LookupRange=A2), 1))
```

---

## Cleaning & Organizing Data

### "I want to... remove duplicates"

**Solution: UNIQUE function or Remove Duplicates feature**

**UNIQUE Function (Excel 365):**
```excel
=UNIQUE(A2:A100)
```

**Sort and get unique:**
```excel
=SORT(UNIQUE(A2:A100))
```

**Manual Method:**
1. Select data range
2. **Data** tab → **Remove Duplicates**
3. Choose columns to check
4. Click OK

**Count unique values:**
```excel
=COUNTA(UNIQUE(A2:A100))
```

---

### "I want to... split data into multiple columns"

**Solution: Text to Columns or TEXTSPLIT**

**TEXTSPLIT (Excel 365):**
```excel
=TEXTSPLIT(A2, ",")          // Split by comma
=TEXTSPLIT(A2, " ")          // Split by space
```

**Text to Columns (All versions):**
1. Select column
2. **Data** tab → **Text to Columns**
3. Choose **Delimited** or **Fixed Width**
4. Select delimiter (comma, space, tab, etc.)
5. Click **Finish**

**Extract specific parts:**
```excel
// First name from "John Smith"
=LEFT(A2, FIND(" ", A2)-1)

// Last name
=RIGHT(A2, LEN(A2)-FIND(" ", A2))
```

---

### "I want to... combine data from multiple columns"

**Solution: CONCAT, TEXTJOIN, or & operator**

**Simple concatenation:**
```excel
=A2 & " " & B2                      // "John" + "Smith" = "John Smith"
```

**TEXTJOIN (Excel 2019+):**
```excel
=TEXTJOIN(" ", TRUE, A2:C2)        // Join with space, ignore blanks
=TEXTJOIN(", ", TRUE, A2:A10)      // Join vertical list with comma
```

**Combine with formatting:**
```excel
=A2 & " - " & B2 & " ($" & C2 & ")"
// Result: "Product - Description ($99.99)"
```

**Full address:**
```excel
=TEXTJOIN(", ", TRUE, Street, City, State, ZIP)
```

---

### "I want to... clean up text with extra spaces"

**Solution: TRIM function**

```excel
=TRIM(A2)                          // Remove extra spaces
=TRIM(UPPER(A2))                   // Clean and uppercase
=TRIM(CLEAN(A2))                   // Remove spaces and non-printable chars
```

**Apply to entire column:**
1. Use formula in helper column
2. Copy results
3. Paste as values over original
4. Delete helper column

---

### "I want to... convert text case"

**Solution: UPPER, LOWER, PROPER**

```excel
=UPPER(A2)                         // "john smith" → "JOHN SMITH"
=LOWER(A2)                         // "JOHN SMITH" → "john smith"
=PROPER(A2)                        // "john smith" → "John Smith"
```

**Fix all caps names:**
```excel
=PROPER(LOWER(A2))                 // Better than just PROPER
```

---

## Calculating & Analyzing

### "I want to... calculate a running total"

**Solution: SUM with expanding range**

```excel
// In cell B2
=SUM($A$2:A2)
```
Copy down - the range expands as you copy

**Result:**
| Amount | Running Total |
|--------|---------------|
| 100 | 100 |
| 50 | 150 |
| 75 | 225 |

**Running total with criteria:**
```excel
=SUMIFS($B$2:B2, $A$2:A2, "Product A")
```

---

### "I want to... calculate percentage of total"

**Solution: Divide by sum with absolute reference**

```excel
=A2/SUM($A$2:$A$10)
```
Format as percentage (Ctrl+Shift+%)

**Percentage change:**
```excel
=(New-Old)/Old
=(B2-A2)/A2
```

**Percentage of goal:**
```excel
=Actual/Goal
=A2/$B$1
```

---

### "I want to... rank values"

**Solution: RANK.EQ or RANK.AVG**

```excel
=RANK.EQ(A2, $A$2:$A$10, 0)       // 0 = descending (1 is highest)
=RANK.EQ(A2, $A$2:$A$10, 1)       // 1 = ascending (1 is lowest)
```

**Dynamic ranking with ties:**
```excel
=RANK.AVG(A2, $A$2:$A$10, 0)      // Average rank for ties
```

**Top N with filter:**
```excel
=FILTER(Names, Sales>=LARGE(Sales, 10))  // Top 10
```

---

### "I want to... calculate subtotals for groups"

**Solution: SUBTOTAL function or PivotTable**

**SUBTOTAL (works with filters):**
```excel
=SUBTOTAL(9, B2:B10)               // 9 = SUM (ignores hidden rows)
=SUBTOTAL(101, B2:B10)             // 101 = AVERAGE (ignores hidden)
```

**By Category with SUMIF:**
```excel
=SUMIF($A$2:$A$100, A2, $B$2:$B$100)
```

**PivotTable (Recommended for complex grouping):**
1. Select data
2. **Insert** tab → **PivotTable**
3. Drag category to Rows
4. Drag values to Values
5. Subtotals appear automatically

---

### "I want to... find top or bottom N values"

**Solution: LARGE, SMALL, or FILTER**

**Top 5 values:**
```excel
=LARGE(A:A, 1)                     // Largest
=LARGE(A:A, 2)                     // 2nd largest
=LARGE(A:A, 5)                     // 5th largest
```

**Bottom 5 values:**
```excel
=SMALL(A:A, 1)                     // Smallest
=SMALL(A:A, 5)                     // 5th smallest
```

**Filter top N (Excel 365):**
```excel
=FILTER(Names, Sales>=LARGE(Sales, 10))
```

**Sort and take top N:**
```excel
=SORT(A2:B100, 2, -1)              // Sort by column 2, descending
// Then manually select top N rows
```

---

### "I want to... calculate weighted average"

**Solution: SUMPRODUCT**

```excel
=SUMPRODUCT(Values, Weights) / SUM(Weights)
```

**Example - Weighted grade:**
```excel
=SUMPRODUCT(B2:B5, C2:C5) / SUM(C2:C5)
// B2:B5 = Scores, C2:C5 = Weights
```

---

## Text Manipulation

### "I want to... extract first name from full name"

**Solution: LEFT and FIND**

```excel
=LEFT(A2, FIND(" ", A2)-1)
```

**If no space exists:**
```excel
=IFERROR(LEFT(A2, FIND(" ", A2)-1), A2)
```

---

### "I want to... extract last name from full name"

**Solution: RIGHT, LEN, and FIND**

```excel
=RIGHT(A2, LEN(A2)-FIND(" ", A2))
```

**Handle multiple spaces:**
```excel
=TRIM(RIGHT(SUBSTITUTE(A2," ", REPT(" ", 100)), 100))
```

---

### "I want to... extract email username or domain"

**Solution: LEFT/RIGHT with FIND**

**Username (before @):**
```excel
=LEFT(A2, FIND("@", A2)-1)
// "user@domain.com" → "user"
```

**Domain (after @):**
```excel
=MID(A2, FIND("@", A2)+1, LEN(A2))
// "user@domain.com" → "domain.com"
```

---

### "I want to... remove or replace characters"

**Solution: SUBSTITUTE**

```excel
=SUBSTITUTE(A2, "-", "")           // Remove hyphens
=SUBSTITUTE(A2, "old", "new")      // Replace text
=SUBSTITUTE(A2, " ", "")           // Remove all spaces
```

**Remove specific character (all occurrences):**
```excel
=SUBSTITUTE(A2, "X", "")
```

**Replace only first occurrence:**
```excel
=SUBSTITUTE(A2, "X", "Y", 1)
```

---

### "I want to... create initials from name"

**Solution: LEFT with FIND**

```excel
=LEFT(A2,1) & LEFT(MID(A2, FIND(" ", A2)+1, 50), 1)
// "John Smith" → "JS"
```

**With period separator:**
```excel
=LEFT(A2,1) & "." & LEFT(MID(A2, FIND(" ", A2)+1, 50), 1) & "."
// "John Smith" → "J.S."
```

---

## Date & Time Operations

### "I want to... calculate age from birthdate"

**Solution: DATEDIF or calculation**

```excel
=DATEDIF(A2, TODAY(), "Y")         // Age in years
=DATEDIF(A2, TODAY(), "YM")        // Months beyond years
=DATEDIF(A2, TODAY(), "MD")        // Days beyond months
```

**Age in years and months:**
```excel
=DATEDIF(A2, TODAY(), "Y") & " years, " & DATEDIF(A2, TODAY(), "YM") & " months"
```

**Simple calculation:**
```excel
=INT((TODAY()-A2)/365.25)
```

---

### "I want to... calculate days between dates"

**Solution: Subtraction**

```excel
=B2-A2                             // Days between
=NETWORKDAYS(A2, B2)               // Business days only
=NETWORKDAYS.INTL(A2, B2, 1)       // Custom weekends
```

**Weeks between:**
```excel
=(B2-A2)/7
```

**Months between:**
```excel
=DATEDIF(A2, B2, "M")
```

---

### "I want to... add or subtract days/months/years"

**Solution: DATE functions**

**Add days:**
```excel
=A2 + 30                           // Add 30 days
=TODAY() + 7                       // 7 days from today
```

**Add months:**
```excel
=EDATE(A2, 3)                      // Add 3 months
=EDATE(TODAY(), -1)                // 1 month ago
```

**Add years:**
```excel
=DATE(YEAR(A2)+1, MONTH(A2), DAY(A2))
```

**Add business days:**
```excel
=WORKDAY(A2, 10)                   // 10 business days from A2
```

---

### "I want to... extract year, month, or day from date"

**Solution: YEAR, MONTH, DAY**

```excel
=YEAR(A2)                          // 2024
=MONTH(A2)                         // 12
=DAY(A2)                           // 14
```

**Month name:**
```excel
=TEXT(A2, "MMMM")                  // "December"
=TEXT(A2, "MMM")                   // "Dec"
```

**Day of week:**
```excel
=TEXT(A2, "DDDD")                  // "Saturday"
=TEXT(A2, "DDD")                   // "Sat"
```

---

### "I want to... get first or last day of month"

**Solution: DATE and EOMONTH**

**First day of month:**
```excel
=DATE(YEAR(A2), MONTH(A2), 1)
```

**Last day of month:**
```excel
=EOMONTH(A2, 0)
```

**Last day of previous month:**
```excel
=EOMONTH(A2, -1)
```

**First day of next month:**
```excel
=EOMONTH(A2, 0) + 1
```

---

## Conditional Operations

### "I want to... count cells that meet criteria"

**Solution: COUNTIF or COUNTIFS**

**Single criteria:**
```excel
=COUNTIF(A:A, ">100")              // Count values >100
=COUNTIF(A:A, "Apple")             // Count "Apple"
=COUNTIF(A:A, "<>"&"")             // Count non-empty
```

**Multiple criteria:**
```excel
=COUNTIFS(A:A, "Apple", B:B, ">100")  // Apple AND >100
```

**Count unique values:**
```excel
=SUMPRODUCT(1/COUNTIF(A2:A100, A2:A100))
```

---

### "I want to... sum cells based on criteria"

**Solution: SUMIF or SUMIFS**

**Single criteria:**
```excel
=SUMIF(A:A, "Apple", B:B)          // Sum B where A="Apple"
=SUMIF(B:B, ">100", B:B)           // Sum values >100
```

**Multiple criteria:**
```excel
=SUMIFS(C:C, A:A, "Apple", B:B, ">100")  // Sum C where A="Apple" AND B>100
```

**Sum with OR logic:**
```excel
=SUMIF(A:A, "Apple", B:B) + SUMIF(A:A, "Orange", B:B)
```

---

### "I want to... apply different calculations based on conditions"

**Solution: IF, IFS, or SWITCH**

**Simple IF:**
```excel
=IF(A2>100, A2*0.1, A2*0.05)       // 10% if >100, else 5%
```

**Multiple conditions (IFS):**
```excel
=IFS(A2>=90, "A", A2>=80, "B", A2>=70, "C", A2>=60, "D", TRUE, "F")
```

**Switch for exact matches:**
```excel
=SWITCH(A2, 1, "Jan", 2, "Feb", 3, "Mar", "Unknown")
```

---

### "I want to... highlight cells based on conditions"

**Solution: Conditional Formatting**

1. Select range
2. **Home** tab → **Conditional Formatting**
3. Choose rule type:
   - **Highlight Cells Rules** - Greater than, Less than, etc.
   - **Top/Bottom Rules** - Top 10, Above Average, etc.
   - **Data Bars** - Visual bars
   - **Color Scales** - Gradient colors
   - **Icon Sets** - Traffic lights, arrows, etc.
4. Set criteria and format

**Custom formula rule:**
1. **New Rule** → **Use a formula**
2. Enter formula (e.g., `=MOD(ROW(),2)=0` for alternating rows)
3. Set format

---

## List Management

### "I want to... create a dropdown list"

**Solution: Data Validation**

1. Select cell(s)
2. **Data** tab → **Data Validation**
3. Allow: **List**
4. Source: Type items separated by commas OR select range
5. Click OK

**Dynamic dropdown from range:**
- Source: `=$A$1:$A$10`

**Dynamic dropdown from table:**
- Source: `=TableName[ColumnName]`

---

### "I want to... prevent duplicate entries"

**Solution: Data Validation with custom formula**

1. Select range
2. **Data** tab → **Data Validation**
3. Allow: **Custom**
4. Formula: `=COUNTIF($A$1:$A$100, A1)=1`
5. Error message: "Duplicate entry not allowed"

---

### "I want to... sort data dynamically"

**Solution: SORT function (Excel 365)**

```excel
=SORT(A2:C100, 1, 1)               // Sort by column 1, ascending
=SORT(A2:C100, 2, -1)              // Sort by column 2, descending
```

**Multiple sort levels:**
```excel
=SORTBY(A2:D100, C2:C100, 1, D2:D100, -1)  // Sort by C asc, then D desc
```

---

### "I want to... generate a unique list automatically"

**Solution: UNIQUE and SORT**

```excel
=SORT(UNIQUE(A2:A100))             // Unique sorted list
=UNIQUE(A2:A100)                   // Unique list (unsorted)
```

**Unique combinations:**
```excel
=UNIQUE(A2:B100)                   // Unique rows based on both columns
```

---

## Visualization & Formatting

### "I want to... create a progress bar in a cell"

**Solution: REPT function**

```excel
=REPT("█", A2/10)                  // Visual bar (0-100 scale)
=REPT("▓", INT(A2*10))             // Percentage bar
```

**With background:**
```excel
=REPT("█", A2/10) & REPT("░", 10-A2/10)
```

---

### "I want to... format numbers with custom text"

**Solution: TEXT function or Custom Number Format**

**TEXT function:**
```excel
=TEXT(A2, "$#,##0.00")             // "$1,234.56"
=TEXT(A2, "0.0%")                  // "15.5%"
="Sales: " & TEXT(A2, "$#,##0")    // "Sales: $1,234"
```

**Custom Number Format:**
1. Select cells
2. Ctrl+1 (Format Cells)
3. Number → Custom
4. Type: `$#,##0.00 "USD"`

---

### "I want to... alternate row colors"

**Solution: Conditional Formatting or Table**

**Conditional Formatting:**
1. Select range
2. Conditional Formatting → New Rule → Use formula
3. Formula: `=MOD(ROW(),2)=0`
4. Set format (background color)

**Convert to Table (Easier):**
1. Select range
2. Ctrl+T or Insert → Table
3. Table automatically has banded rows

---

**[⬆ Back to Main README](../../README.md)**
