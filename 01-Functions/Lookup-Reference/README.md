# 🔍 Lookup & Reference Functions

> **30+ functions for finding, referencing, and retrieving data - the most powerful tools in Excel**

## 📋 Table of Contents

- [Essential Lookup Functions](#essential-lookup-functions)
- [Modern Dynamic Arrays](#modern-dynamic-arrays)
- [Reference Functions](#reference-functions)
- [Advanced Lookup Techniques](#advanced-lookup-techniques)

---

## Essential Lookup Functions

### VLOOKUP
**Vertical lookup - searches first column and returns value from same row**

**Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

**Parameters:**
- `lookup_value`: Value to search for
- `table_array`: Table to search in
- `col_index_num`: Column number to return (1-based)
- `range_lookup`: FALSE (exact), TRUE (approximate, default)

**Examples:**
```excel
=VLOOKUP("Apple", A1:C10, 2, FALSE)     → Find "Apple" in column A, return column B
=VLOOKUP(D1, PriceTable, 3, FALSE)      → Lookup price from table
=VLOOKUP(E1, A:D, 4, 0)                 → Search entire column (0 = FALSE)
```

**Real-World Uses:**
- Price lookups from product list
- Employee information by ID
- Product details by SKU
- Customer data by account number

**Important Notes:**
- Searches ONLY first column
- Returns value from RIGHT of lookup column
- Case-insensitive
- Use FALSE for exact match (recommended)

**Common Errors:**
- `#N/A` - Value not found
- `#REF!` - Column index too large
- `#VALUE!` - Wrong data types

**With Error Handling:**
```excel
=IFERROR(VLOOKUP(A1, Table, 2, 0), "Not Found")
=IFNA(VLOOKUP(A1, Table, 2, 0), "Missing")
```

**Limitations:**
- Can't lookup left (use INDEX/MATCH)
- Column changes break formula
- Not dynamic (won't expand with table)

---

### HLOOKUP
**Horizontal lookup - searches first row and returns value from same column**

**Syntax:** `=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

**Examples:**
```excel
=HLOOKUP("Q1", A1:D10, 5, FALSE)        → Find "Q1" in row 1, return row 5
=HLOOKUP(A1, MonthlyData, 3, 0)         → Lookup from horizontal table
```

**Use:** When data organized in rows instead of columns

**Note:** VLOOKUP is more common as data usually organized vertically

---

### XLOOKUP
**Modern replacement for VLOOKUP/HLOOKUP (Excel 365/2021)**

**Syntax:** `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`

**Parameters:**
- `lookup_value`: Value to find
- `lookup_array`: Where to search
- `return_array`: Where to return value from
- `if_not_found`: (Optional) Value if not found
- `match_mode`: 0=exact (default), -1=exact or next smaller, 1=exact or next larger, 2=wildcard
- `search_mode`: 1=first to last (default), -1=last to first, 2=binary asc, -2=binary desc

**Examples:**
```excel
=XLOOKUP(A1, ProductIDs, Prices)                    → Basic lookup
=XLOOKUP(A1, IDs, Prices, "Not Found")              → With custom message
=XLOOKUP(A1, Names, A1:Z1)                          → Return entire row
=XLOOKUP(A1, Dates, Sales, , 1, 1)                  → Approximate match
```

**Advantages over VLOOKUP:**
✅ Can lookup left (any direction)
✅ Built-in error handling
✅ Returns multiple columns/rows
✅ Can search backwards
✅ Faster and more flexible

**Real-World Examples:**

**Two-Way Lookup:**
```excel
=XLOOKUP(A1, Names, XLOOKUP(B1, Months, DataTable))
```

**Return Multiple Columns:**
```excel
=XLOOKUP(A1, IDs, B1:E100)              → Returns entire row from B:E
```

**Last Occurrence:**
```excel
=XLOOKUP(A1, Transactions, Dates, , 0, -1)  → Find last transaction date
```

---

### INDEX
**Returns value at specific row and column intersection**

**Syntax:** `=INDEX(array, row_num, [column_num])`

**Examples:**
```excel
=INDEX(A1:C10, 5, 2)                    → Value at row 5, column 2
=INDEX(A:A, 10)                         → 10th value in column A
=INDEX(B1:B100, MATCH(A1, A1:A100, 0))  → Dynamic lookup
```

**Real-World Uses:**
- Dynamic lookups with MATCH
- Return entire row or column
- Two-way lookups
- More flexible than VLOOKUP

**Return Entire Row:**
```excel
=INDEX(A1:Z100, 5, 0)                   → Entire row 5
```

**Return Entire Column:**
```excel
=INDEX(A1:Z100, 0, 3)                   → Entire column 3 (C)
```

---

### MATCH
**Returns position of value in range**

**Syntax:** `=MATCH(lookup_value, lookup_array, [match_type])`

**Parameters:**
- `lookup_value`: Value to find
- `lookup_array`: Range to search
- `match_type`: 0=exact, 1=less than or equal (default), -1=greater than or equal

**Examples:**
```excel
=MATCH("Apple", A1:A10, 0)              → Returns 5 if "Apple" is in A5
=MATCH(100, B1:B50, 1)                  → Position of largest value ≤100
=MATCH(A1, 1:1, 0)                      → Find column position
```

**Real-World Uses:**
- Find position for INDEX
- Determine rank/order
- Validate existence
- Dynamic column/row reference

**INDEX/MATCH Combination (Better than VLOOKUP):**
```excel
=INDEX(ReturnRange, MATCH(LookupValue, LookupRange, 0))
```

**Advantages:**
- Can lookup left
- Column insertions don't break formula
- More flexible
- Faster for large datasets

**Two-Way Lookup:**
```excel
=INDEX(DataRange, MATCH(RowValue, RowHeaders, 0), MATCH(ColValue, ColHeaders, 0))
```

---

### XMATCH
**Modern MATCH with more options (Excel 365/2021)**

**Syntax:** `=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])`

**Examples:**
```excel
=XMATCH(A1, B1:B100)                    → Position of A1 in range
=XMATCH(A1, B1:B100, 0, -1)             → Search from bottom
=INDEX(C1:C100, XMATCH(A1, B1:B100))    → Dynamic lookup
```

**Advantages over MATCH:**
- More match modes
- Search direction control
- Better performance

---

## Modern Dynamic Arrays

### FILTER
**Returns filtered array based on criteria (Excel 365)**

**Syntax:** `=FILTER(array, include, [if_empty])`

**Examples:**
```excel
=FILTER(A1:C100, B1:B100>50)                    → Rows where column B > 50
=FILTER(A1:D100, C1:C100="Apple")               → Rows where column C = "Apple"
=FILTER(A1:D100, (B1:B100>100)*(C1:C100="North"), "No Results")
```

**Real-World Uses:**
- Filter sales by region
- Extract records by criteria
- Dynamic dashboards
- Create custom views

**Multiple Criteria (AND):**
```excel
=FILTER(Data, (Region="North")*(Sales>1000))
```

**Multiple Criteria (OR):**
```excel
=FILTER(Data, (Region="North")+(Region="South"))
```

**Top N Results:**
```excel
=FILTER(Data, Sales>=LARGE(Sales, 10))          → Top 10
```

---

### SORT
**Sorts array by specified columns (Excel 365)**

**Syntax:** `=SORT(array, [sort_index], [sort_order], [by_col])`

**Parameters:**
- `array`: Range to sort
- `sort_index`: Column/row number to sort by (default: 1)
- `sort_order`: 1=ascending (default), -1=descending
- `by_col`: FALSE=sort by row (default), TRUE=sort by column

**Examples:**
```excel
=SORT(A1:C100, 1, 1)                    → Sort by column 1, ascending
=SORT(A1:C100, 3, -1)                   → Sort by column 3, descending
=SORT(A1:C100)                          → Sort by first column, ascending
```

**Real-World Uses:**
- Dynamic sorted lists
- Leaderboards
- Alphabetical directories
- Sort by date/value

---

### SORTBY
**Sorts array by another array (Excel 365)**

**Syntax:** `=SORTBY(array, by_array1, [sort_order1], [by_array2], [sort_order2], ...)`

**Examples:**
```excel
=SORTBY(A1:A100, B1:B100, -1)           → Sort A by B descending
=SORTBY(A1:D100, C1:C100, -1, D1:D100, 1)  → Sort by C desc, then D asc
```

**Use:** Sort by column not in result range

---

### UNIQUE
**Returns unique values from range (Excel 365)**

**Syntax:** `=UNIQUE(array, [by_col], [exactly_once])`

**Parameters:**
- `array`: Range to extract unique values
- `by_col`: FALSE=unique rows (default), TRUE=unique columns
- `exactly_once`: FALSE=all unique (default), TRUE=values that occur only once

**Examples:**
```excel
=UNIQUE(A1:A100)                        → Unique values from A
=UNIQUE(A1:C100)                        → Unique rows
=UNIQUE(A1:A100, FALSE, TRUE)           → Values appearing exactly once
```

**Real-World Uses:**
- Remove duplicates dynamically
- Create dropdown lists
- Count distinct values
- Extract unique customers/products

**Count Unique:**
```excel
=COUNTA(UNIQUE(A1:A100))                → Count distinct values
```

**Unique with Sort:**
```excel
=SORT(UNIQUE(A1:A100))                  → Sorted unique list
```

---

### SEQUENCE
**Generates sequence of numbers (Excel 365)**

**Syntax:** `=SEQUENCE(rows, [columns], [start], [step])`

**Examples:**
```excel
=SEQUENCE(10)                           → 1 to 10
=SEQUENCE(10, 1, 0)                     → 0 to 9
=SEQUENCE(5, 5, 1)                      → 5x5 grid (1-25)
=SEQUENCE(10, 1, 100, 10)               → 100, 110, 120...190
```

**Real-World Uses:**
- Generate number series
- Create calendars
- Testing data
- Dynamic row numbers

**With Dates:**
```excel
=TODAY() + SEQUENCE(7) - 1              → Next 7 days
```

---

### RANDARRAY
**Generates array of random numbers (Excel 365)**

**Syntax:** `=RANDARRAY([rows], [columns], [min], [max], [integer])`

**Examples:**
```excel
=RANDARRAY(10, 1)                       → 10 random decimals 0-1
=RANDARRAY(5, 5, 1, 100, TRUE)          → 5x5 integers 1-100
=RANDARRAY(10, 1, 1, 6, TRUE)           → 10 dice rolls
```

**Use:** Testing, simulations, sampling

---

## Reference Functions

### OFFSET
**Returns reference offset from starting cell**

**Syntax:** `=OFFSET(reference, rows, cols, [height], [width])`

**Parameters:**
- `reference`: Starting cell
- `rows`: Rows to offset (+ down, - up)
- `cols`: Columns to offset (+ right, - left)
- `height`: (Optional) Height of range
- `width`: (Optional) Width of range

**Examples:**
```excel
=OFFSET(A1, 5, 0)                       → A6 (5 rows down)
=OFFSET(A1, 0, 3)                       → D1 (3 columns right)
=OFFSET(A1, 2, 2, 5, 3)                 → Range starting at C3, 5 tall, 3 wide
=SUM(OFFSET(A1, 0, 0, 10, 1))           → Sum of 10 cells starting A1
```

**Real-World Uses:**
- Dynamic ranges
- Rolling averages
- Expanding ranges
- Reference data by position

**Dynamic Range (Last 12 Months):**
```excel
=OFFSET(A1, COUNTA(A:A)-12, 0, 12, 1)
```

**Warning:** OFFSET is volatile (recalculates often) - use sparingly

---

### INDIRECT
**Returns reference specified by text string**

**Syntax:** `=INDIRECT(ref_text, [a1])`

**Examples:**
```excel
=INDIRECT("A1")                         → Value in A1
=INDIRECT("A" & ROW())                  → Reference based on current row
=INDIRECT(A1 & "!B5")                   → Cell B5 in sheet named in A1
=SUM(INDIRECT("Sheet" & A1 & "!A1:A10"))  → Sum from dynamic sheet
```

**Real-World Uses:**
- Dynamic sheet references
- Reference cells by formula
- Create flexible formulas
- Multi-sheet consolidation

**Dynamic Sheet Reference:**
```excel
=INDIRECT("'" & A1 & "'!B1")            → Cell B1 in sheet named in A1
```

**Warning:** Also volatile - use carefully

---

### CHOOSE
**Returns value from list based on index**

**Syntax:** `=CHOOSE(index_num, value1, [value2], ...)`

**Examples:**
```excel
=CHOOSE(2, "Red", "Blue", "Green")      → "Blue"
=CHOOSE(A1, 100, 200, 300)              → Returns based on A1 value (1-3)
=CHOOSE(WEEKDAY(TODAY()), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
```

**Real-World Uses:**
- Convert number to text
- Select from options
- Day/month names
- Dynamic references

**Dynamic Column:**
```excel
=CHOOSE(A1, B:B, C:C, D:D)              → Select column based on A1
```

---

### ROW & COLUMN
**Returns row or column number**

**Syntax:**
```excel
=ROW([reference])
=COLUMN([reference])
```

**Examples:**
```excel
=ROW()                                  → Current row number
=ROW(A10)                               → 10
=COLUMN()                               → Current column number
=COLUMN(D1)                             → 4
```

**Real-World Uses:**
- Sequential numbering
- Alternating row colors
- Array formulas
- Dynamic references

**Auto-Numbering:**
```excel
=ROW()-1                                → In A2: shows 1, A3: shows 2, etc.
```

**Stripe Rows:**
```excel
=MOD(ROW(), 2) = 0                      → TRUE for even rows
```

---

### ROWS & COLUMNS
**Returns count of rows or columns in reference**

**Syntax:**
```excel
=ROWS(array)
=COLUMNS(array)
```

**Examples:**
```excel
=ROWS(A1:A10)                           → 10
=COLUMNS(A1:E1)                         → 5
=ROWS(DataRange)                        → Number of rows in named range
```

**Use:** Dynamic range calculations

---

### ADDRESS
**Returns cell address as text**

**Syntax:** `=ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])`

**Examples:**
```excel
=ADDRESS(10, 5)                         → "$E$10"
=ADDRESS(10, 5, 4)                      → "E10" (relative)
=ADDRESS(10, 5, 1, TRUE, "Sheet2")      → "Sheet2!$E$10"
```

**Use:** Build cell references dynamically

---

### AREAS
**Returns number of areas in reference**

**Syntax:** `=AREAS(reference)`

**Examples:**
```excel
=AREAS(A1:B10)                          → 1
=AREAS((A1:A10,C1:C10))                 → 2
```

**Use:** Check multi-area ranges

---

## Advanced Lookup Techniques

### INDEX/MATCH (Better than VLOOKUP)

**Basic Lookup:**
```excel
=INDEX(ReturnColumn, MATCH(LookupValue, LookupColumn, 0))
```

**Example:**
```excel
=INDEX(C:C, MATCH(A1, B:B, 0))          → Look up A1 in column B, return from C
```

**Advantages:**
- Lookup left or right
- Column changes don't break formula
- Faster for large data
- More flexible

**Two-Way Lookup (Matrix):**
```excel
=INDEX(DataTable, MATCH(RowHeader, RowRange, 0), MATCH(ColHeader, ColRange, 0))
```

**Example:**
```excel
=INDEX(B2:M13, MATCH(A15, A2:A13, 0), MATCH(B14, B1:M1, 0))
// Look up value at intersection of row and column
```

---

### Multiple Criteria Lookup

**Using XLOOKUP (Excel 365):**
```excel
=XLOOKUP(1, (Criteria1=Range1)*(Criteria2=Range2), ReturnRange)
```

**Using INDEX/MATCH:**
```excel
=INDEX(ReturnRange, MATCH(1, (Criteria1=Range1)*(Criteria2=Range2), 0))
// Enter as array formula (Ctrl+Shift+Enter in older Excel)
```

**Example:**
```excel
=INDEX(Prices, MATCH(1, (Products=A1)*(Regions=B1), 0))
// Find price matching both product and region
```

---

### Reverse Lookup (Right to Left)

**INDEX/MATCH:**
```excel
=INDEX(A:A, MATCH(E1, C:C, 0))          → Lookup in C, return from A
```

**XLOOKUP:**
```excel
=XLOOKUP(E1, C:C, A:A)                  → Even simpler
```

---

### Approximate Match Lookup

**Find closest value ≤ lookup value:**
```excel
=VLOOKUP(A1, Table, 2, TRUE)            → Must be sorted ascending
=XLOOKUP(A1, Range1, Range2, , 1)       → Exact or next smaller
```

**Use:** Tax brackets, pricing tiers, grade cutoffs

---

### Lookup with Wildcards

**Partial Match:**
```excel
=VLOOKUP("*apple*", A:B, 2, FALSE)      → Find any cell containing "apple"
=XLOOKUP(A1&"*", Range1, Range2, , 2)   → Wildcard match
```

---

### Return Multiple Values

**XLOOKUP (Returns array):**
```excel
=XLOOKUP(A1, IDs, B1:E100)              → Returns entire row B:E
```

**FILTER:**
```excel
=FILTER(A1:D100, B1:B100=E1)            → All rows matching criteria
```

---

### Last Occurrence Lookup

**XLOOKUP:**
```excel
=XLOOKUP(A1, Range1, Range2, , 0, -1)   → Search from bottom
```

**INDEX/MATCH:**
```excel
=INDEX(ReturnRange, MATCH(2, 1/(LookupRange=LookupValue), 1))
```

---

## Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| VLOOKUP | Lookup in column | `=VLOOKUP(A1,Table,2,0)` |
| XLOOKUP | Modern lookup | `=XLOOKUP(A1,IDs,Prices)` |
| INDEX | Get value at position | `=INDEX(A:A,10)` |
| MATCH | Find position | `=MATCH(A1,B:B,0)` |
| FILTER | Filter by criteria | `=FILTER(A:C,B:B>100)` |
| SORT | Sort array | `=SORT(A:A,-1)` |
| UNIQUE | Unique values | `=UNIQUE(A:A)` |
| OFFSET | Dynamic reference | `=OFFSET(A1,5,0)` |
| INDIRECT | Text to reference | `=INDIRECT("A"&ROW())` |
| CHOOSE | Select from list | `=CHOOSE(2,"A","B","C")` |

---

## Best Practices

### Use XLOOKUP Instead of VLOOKUP (When Available)
```excel
✓ =XLOOKUP(A1, IDs, Prices, "Not Found")
✗ =IFERROR(VLOOKUP(A1, Table, 2, 0), "Not Found")
```

### Use INDEX/MATCH Instead of VLOOKUP (Pre-365)
```excel
✓ =INDEX(Prices, MATCH(A1, IDs, 0))
✗ =VLOOKUP(A1, A:B, 2, 0)
```

### Always Use FALSE/0 for Exact Match
```excel
✓ =VLOOKUP(A1, Table, 2, FALSE)
✗ =VLOOKUP(A1, Table, 2)
```

### Use Named Ranges
```excel
✓ =XLOOKUP(A1, ProductIDs, ProductNames)
✗ =XLOOKUP(A1, Sheet1!$B$2:$B$1000, Sheet1!$C$2:$C$1000)
```

### Avoid Volatile Functions When Possible
```excel
✓ Use Tables with structured references
✗ Excessive OFFSET/INDIRECT
```

---

**[⬆ Back to Main README](../../README.md)**
