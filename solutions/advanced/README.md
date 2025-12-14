# Advanced Excel Techniques

> **Power user formulas, dynamic arrays, automation, and expert-level solutions**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Create dynamic ranges | [Dynamic Arrays](#dynamic-arrays-excel-365) |
| Build dependent dropdowns | [Cascading Dropdowns](#cascading-dropdowns) |
| Create unique lists | [Unique Values](#extract-unique-values) |
| Sort with formulas | [Formula-Based Sorting](#sort-with-formulas) |
| Build dynamic tables | [LAMBDA Functions](#lambda-custom-functions) |
| Handle arrays | [Array Formulas](#array-formulas) |
| Create complex conditions | [Advanced Logic](#advanced-conditional-logic) |
| Work with 3D references | [Multi-Sheet Formulas](#3d-references) |
| Build self-updating ranges | [Structured References](#structured-references) |
| Create error-proof formulas | [Robust Formulas](#bulletproof-formulas) |

---

## Dynamic Arrays (Excel 365)

### What Are Dynamic Arrays?
Formulas that return multiple values and automatically "spill" into adjacent cells.

### Key Dynamic Array Functions

| Function | Purpose | Example |
|----------|---------|---------|
| FILTER | Filter rows by criteria | `=FILTER(A:D, B:B="Sales")` |
| SORT | Sort data | `=SORT(A1:C10, 2, -1)` |
| SORTBY | Sort by another column | `=SORTBY(A1:B10, C1:C10)` |
| UNIQUE | Extract unique values | `=UNIQUE(A1:A100)` |
| SEQUENCE | Generate number series | `=SEQUENCE(10, 1, 1, 1)` |
| RANDARRAY | Random array | `=RANDARRAY(5, 3, 1, 100, TRUE)` |

### FILTER Examples
```excel
Basic:           =FILTER(A2:D100, B2:B100="North")
Multiple AND:    =FILTER(A2:D100, (B2:B100="North")*(C2:C100>1000))
Multiple OR:     =FILTER(A2:D100, (B2:B100="North")+(B2:B100="South"))
With Default:    =FILTER(A2:D100, B2:B100="North", "No results")
Select Columns:  =FILTER(CHOOSECOLS(A2:D100,1,3), B2:B100="North")
```

### SORT and SORTBY
```excel
Sort ascending:        =SORT(A1:C10, 2, 1)
Sort descending:       =SORT(A1:C10, 2, -1)
Multi-column sort:     =SORT(A1:C10, {2,1}, {1,-1})
Sort by external col:  =SORTBY(A1:B10, C1:C10, -1)
```

### UNIQUE with Options
```excel
Unique column:         =UNIQUE(A2:A100)
Unique rows:           =UNIQUE(A2:C100)
Unique by column:      =UNIQUE(A2:C100, FALSE, FALSE)
Exactly once only:     =UNIQUE(A2:A100, FALSE, TRUE)
```

### SEQUENCE Applications
```excel
Numbers 1-10:          =SEQUENCE(10)
Row numbers:           =SEQUENCE(ROWS(A1:A100))
Dates (next 7 days):   =SEQUENCE(7, 1, TODAY(), 1)
Times (hourly):        =SEQUENCE(24, 1, 0, 1/24)
Multiplication table:  =SEQUENCE(10, 10, 1, 1) * SEQUENCE(1, 10, 1, 1)
```

### Combining Dynamic Functions
```excel
Top 5 by sales:
=TAKE(SORT(FILTER(A2:C100, B2:B100="North"), 3, -1), 5)

Unique sorted list:
=SORT(UNIQUE(A2:A100))

Filter and count:
=ROWS(FILTER(A2:A100, B2:B100="Active"))
```

---

## Cascading Dropdowns

### The Challenge
Create a dropdown where choices depend on another dropdown selection.

### Setup

**Named Ranges:**
| Name | Range |
|------|-------|
| Categories | {Fruit, Vegetable, Dairy} |
| Fruit | {Apple, Banana, Orange} |
| Vegetable | {Carrot, Broccoli, Spinach} |
| Dairy | {Milk, Cheese, Yogurt} |

**First Dropdown (B1):**
1. Data → Data Validation
2. Allow: List
3. Source: `=Categories`

**Second Dropdown (C1):**
1. Data → Data Validation
2. Allow: List
3. Source: `=INDIRECT(B1)`

### Dynamic with FILTER (Excel 365)
```excel
Source for dropdown 2: =FILTER(Items, Categories=B1)
```

---

## Extract Unique Values

### UNIQUE Function (Excel 365)
```excel
=UNIQUE(A2:A100)                    → Unique values
=UNIQUE(A2:B100)                    → Unique row combinations
=UNIQUE(A2:A100, FALSE, TRUE)       → Values appearing exactly once
```

### Legacy Method (No UNIQUE)
```excel
=IFERROR(INDEX($A$2:$A$100, MATCH(0, COUNTIF($C$1:C1, $A$2:$A$100), 0)), "")
```
(Enter as array formula, copy down)

### Count Unique Values
```excel
Excel 365:  =ROWS(UNIQUE(A2:A100))
Legacy:     =SUMPRODUCT(1/COUNTIF(A2:A100, A2:A100))
```

---

## Sort with Formulas

### SORT Function (Excel 365)
```excel
=SORT(range, sort_index, sort_order, by_col)
```

**Examples:**
```excel
Ascending:              =SORT(A2:C10, 1, 1)
Descending:             =SORT(A2:C10, 2, -1)
By multiple columns:    =SORT(A2:C10, {1,2}, {1,-1})
```

### Legacy Sort (Using INDEX/MATCH)
Rank first, then retrieve:
```excel
Rank:  =SUMPRODUCT((B$2:B$100>B2)*1)+1
Sort:  =INDEX($A$2:$A$100, MATCH(ROW()-1, RankColumn, 0))
```

### SORTBY for External Criteria
```excel
=SORTBY(Names, Scores, -1)
```
Sort Names by Scores descending.

---

## LAMBDA Custom Functions

### What is LAMBDA?
Create your own reusable functions without VBA.

### Basic Syntax
```excel
=LAMBDA(parameter1, parameter2, calculation)
```

### Examples

**Create a function called CELSIUS:**
Name Manager → New
- Name: CELSIUS
- Refers to: `=LAMBDA(fahrenheit, (fahrenheit-32)*5/9)`

**Usage:**
```excel
=CELSIUS(98.6)   → 37
```

**Tax Calculator:**
```excel
Name: ADDTAX
Formula: =LAMBDA(amount, rate, amount * (1 + rate))
Usage: =ADDTAX(100, 0.08)   → 108
```

**Conditional with LAMBDA:**
```excel
Name: GRADE
Formula: =LAMBDA(score, IF(score>=90,"A",IF(score>=80,"B",IF(score>=70,"C","F"))))
Usage: =GRADE(85)   → "B"
```

### LET Function (Named Variables)
```excel
=LET(
    data, A2:A100,
    avg, AVERAGE(data),
    stdev, STDEV.S(data),
    FILTER(data, ABS(data-avg) <= 2*stdev)
)
```
(Filter values within 2 standard deviations)

---

## Array Formulas

### Enter Array Formulas
- **Excel 365:** Just press Enter (auto-spill)
- **Legacy:** Press Ctrl+Shift+Enter

### Classic Array Examples

**Sum of products:**
```excel
=SUM(A1:A10 * B1:B10)
```

**Count with multiple criteria:**
```excel
=SUM((A1:A100="North") * (B1:B100>1000))
```

**Average of top 5:**
```excel
=AVERAGE(LARGE(A1:A100, {1,2,3,4,5}))
```

**Sum ignoring errors:**
```excel
=SUM(IF(ISERROR(A1:A100), 0, A1:A100))
```

### Comparing Arrays
```excel
All match:    =AND(A1:A10 = B1:B10)
Any match:    =OR(A1:A10 = B1:B10)
Count matches: =SUM(--(A1:A10 = B1:B10))
```

---

## Advanced Conditional Logic

### Nested IFS (Cleaner than Nested IF)
```excel
=IFS(
    A1>=90, "A",
    A1>=80, "B",
    A1>=70, "C",
    A1>=60, "D",
    TRUE, "F"
)
```

### SWITCH for Exact Matches
```excel
=SWITCH(A1,
    1, "January",
    2, "February",
    3, "March",
    "Unknown"
)
```

### Complex AND/OR
```excel
=IF(AND(A1>10, OR(B1="Yes", C1>100)), "Pass", "Fail")
```

### CHOOSE for Index-Based Selection
```excel
=CHOOSE(WEEKDAY(TODAY()), "Sun","Mon","Tue","Wed","Thu","Fri","Sat")
```

### Boolean Math
```excel
TRUE = 1, FALSE = 0

Count where both conditions: =SUM((A:A="Yes")*(B:B>100))
Count where either condition: =SUM(SIGN((A:A="Yes")+(B:B>100)))
```

---

## 3D References

### Sum Across Multiple Sheets
```excel
=SUM(Sheet1:Sheet12!B5)
```
Sums cell B5 from all sheets between Sheet1 and Sheet12.

### Average Across Sheets
```excel
=AVERAGE(Jan:Dec!C10)
```

### Count Across Sheets
```excel
=SUM(Sheet1:Sheet12!A:A)
```

### 3D Reference Requirements
- Sheets must be contiguous (next to each other)
- Same cell reference on each sheet
- Works with: SUM, AVERAGE, COUNT, MAX, MIN, PRODUCT

---

## Structured References

### Table References
When data is formatted as Table (Ctrl+T):

| Reference | Meaning |
|-----------|---------|
| `Table1[Sales]` | Entire Sales column |
| `Table1[@Sales]` | Sales in current row |
| `Table1[[#Headers],[Sales]]` | Header of Sales column |
| `Table1[[#Totals],[Sales]]` | Total of Sales column |
| `Table1[@]` | Entire current row |
| `Table1[#All]` | Entire table with headers |
| `Table1[#Data]` | Data without headers |

### Benefits
- Auto-expand when data added
- Self-documenting formulas
- No need to update ranges

### Example Formulas
```excel
=SUM(Table1[Sales])
=AVERAGE(Table1[Price])
=SUMIFS(Table1[Sales], Table1[Region], "North")
```

---

## Bulletproof Formulas

### Error Handling Wrapper
```excel
=IFERROR(your_formula, "Error")
=IFNA(VLOOKUP(...), "Not Found")
```

### Check Before Calculate
```excel
=IF(B1=0, 0, A1/B1)                    → Prevent #DIV/0!
=IF(ISBLANK(A1), "", your_formula)     → Handle blanks
=IF(ISNUMBER(A1), A1*2, 0)             → Verify type
```

### Defensive VLOOKUP
```excel
=IF(COUNTIF(LookupRange, Value)>0,
    VLOOKUP(Value, Table, 2, FALSE),
    "Not Found")
```

### Handle Missing Sheets
```excel
=IFERROR(INDIRECT("'"&SheetName&"'!A1"), "Sheet not found")
```

### Validate Inputs
```excel
=IF(AND(ISNUMBER(A1), A1>0, A1<1000),
    CalculationFormula,
    "Invalid input")
```

---

## Advanced INDIRECT Uses

### Dynamic Sheet Reference
```excel
=INDIRECT("'" & A1 & "'!B5")
```
Where A1 contains sheet name.

### Dynamic Range
```excel
=SUM(INDIRECT("A1:A" & B1))
```
Where B1 contains row number.

### Dynamic Column
```excel
=INDIRECT("R1C" & A1, FALSE)
```
R1C1 style for column number in A1.

### Dynamic Named Range
```excel
=SUM(INDIRECT(A1))
```
Where A1 contains named range name.

---

## AGGREGATE Function

### Ignore Errors, Hidden Rows, or Nested Functions
```excel
=AGGREGATE(function_num, options, range)
```

### Function Numbers
| Num | Function |
|-----|----------|
| 1 | AVERAGE |
| 2 | COUNT |
| 4 | MAX |
| 5 | MIN |
| 9 | SUM |
| 14 | LARGE |
| 15 | SMALL |

### Options
| Opt | Ignores |
|-----|---------|
| 5 | Hidden rows |
| 6 | Error values |
| 7 | Hidden rows and errors |

### Examples
```excel
=AGGREGATE(9, 6, A1:A100)       → SUM ignoring errors
=AGGREGATE(14, 6, A1:A100, 3)   → 3rd largest, ignoring errors
=AGGREGATE(5, 5, A1:A100)       → MIN of visible rows only
```

---

## Performance Tips for Complex Formulas

1. **Use Tables** - Auto-expanding, faster than whole-column references
2. **Avoid Volatile Functions** - INDIRECT, OFFSET, NOW, RAND recalculate constantly
3. **Replace OFFSET with INDEX** - INDEX is non-volatile
4. **Limit Array Size** - Don't use A:A when A1:A1000 works
5. **Use Helper Columns** - Break complex formulas into steps
6. **Manual Calculation** - For large workbooks, use Formulas → Calculation Options → Manual

---

## Related Solutions

- [Lookups](../lookups/README.md) - XLOOKUP, INDEX/MATCH
- [Data Analysis](../data-analysis/README.md) - Statistical functions
- [Conditional Calculations](../conditional-calculations/README.md) - SUMIFS, COUNTIFS
- [Error Handling](../error-handling/README.md) - IFERROR, IFNA

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
