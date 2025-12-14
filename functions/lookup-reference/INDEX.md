# INDEX

## What It Does (Plain English)
Returns the value at a specific row and column position in a range - like GPS coordinates for your data. Often paired with MATCH to create flexible lookups.

## Syntax
```
=INDEX(array, row_num, [column_num])
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| array | Yes | The range of cells to return a value from |
| row_num | Yes | Which row within the array (1 = first row) |
| column_num | No | Which column within the array (1 = first column). If omitted and array is one column, returns from that column |

## Returns
The value at the specified row and column intersection within the array.

## Examples

### Example 1: Get Value at Specific Position
**Data (A1:C4):**
| A | B | C |
|---|---|---|
| Apple | Red | 1.50 |
| Banana | Yellow | 0.75 |
| Cherry | Red | 3.00 |
| Date | Brown | 4.50 |

**Formula:** `=INDEX(A1:C4, 2, 3)`

**Result:** `0.75`

**Explanation:** Returns the value at row 2, column 3 of the range = the price of Banana.

---

### Example 2: Single Column (Common Use)
**Data (B1:B5):**
| B |
|---|
| Jan |
| Feb |
| Mar |
| Apr |
| May |

**Formula:** `=INDEX(B1:B5, 3)`

**Result:** `Mar`

**Explanation:** For a single column, you only need the row number.

---

### Example 3: INDEX + MATCH (The Power Combo)
**Data (A1:C5):**
| A | B | C |
|---|---|---|
| Product | Category | Price |
| Laptop | Electronics | 999 |
| Chair | Furniture | 199 |
| Phone | Electronics | 699 |
| Desk | Furniture | 349 |

**Formula:** `=INDEX(C2:C5, MATCH("Phone", A2:A5, 0))`

**Result:** `699`

**Explanation:** MATCH finds "Phone" at position 3, INDEX returns the 3rd value in the price column.

---

### Example 4: Two-Way Lookup with INDEX/MATCH
**Data - Sales by Region and Quarter (A1:E4):**
| | Q1 | Q2 | Q3 | Q4 |
|---|---|---|---|---|
| North | 100 | 120 | 130 | 150 |
| South | 80 | 90 | 95 | 110 |
| East | 90 | 100 | 115 | 125 |

**Formula:** `=INDEX(B2:E4, MATCH("South", A2:A4, 0), MATCH("Q3", B1:E1, 0))`

**Result:** `95`

**Explanation:**
- MATCH("South", A2:A4, 0) returns 2 (row position)
- MATCH("Q3", B1:E1, 0) returns 3 (column position)
- INDEX returns value at row 2, column 3 = 95

---

### Example 5: Return Entire Row
**Same data as Example 4**

**Formula:** `=INDEX(B2:E4, 2, 0)`

**Result:** `{80, 90, 95, 110}` (spills across 4 cells)

**Explanation:** Using 0 for column_num returns the entire row 2.

---

### Example 6: Return Entire Column
**Formula:** `=INDEX(B2:E4, 0, 3)`

**Result:** `{130, 95, 115}` (spills down 3 cells)

**Explanation:** Using 0 for row_num returns the entire column 3 (Q3 data).

---

### Example 7: Dynamic Range Reference
**Scenario:** Create a chart that always shows the last 12 months of data.

**Formula for chart data range:**
`=INDEX(Sales, ROWS(Sales)-11, 1):INDEX(Sales, ROWS(Sales), 1)`

**Explanation:** Uses INDEX to create a reference to the last 12 rows of the Sales range.

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #REF! | Row or column number exceeds array dimensions | Check that row_num ≤ ROWS(array) and column_num ≤ COLUMNS(array) |
| #VALUE! | Row or column number is 0 or negative | Use positive integers (1 or greater) |
| #N/A (with MATCH) | MATCH didn't find the value | Check MATCH separately, ensure lookup value exists |

## Pro Tips

- **INDEX/MATCH > VLOOKUP:** This combo can look left, is column-insert proof, and more flexible
- **Return a reference:** INDEX actually returns a reference, so you can use it in other functions like SUM: `=SUM(INDEX(Data,1,1):INDEX(Data,10,1))`
- **Array form:** INDEX has an array form for working with multiple areas - rarely needed
- **With SMALL/LARGE:** Combine to get nth smallest/largest: `=INDEX(A:A, MATCH(LARGE(B:B, 3), B:B, 0))`
- **Named ranges:** Use with named ranges for clearer formulas: `=INDEX(Products, 5, 2)`

## Why Use INDEX/MATCH Over VLOOKUP?

| Advantage | INDEX/MATCH | VLOOKUP |
|-----------|-------------|---------|
| Look left | ✅ Yes | ❌ No |
| Insert columns without breaking | ✅ Yes | ❌ No |
| Separate lookup and return ranges | ✅ Yes | ❌ No |
| Performance on large data | ✅ Faster | ❌ Slower |
| Return row/column reference | ✅ Yes | ❌ No |

## Related Functions

| Function | When to Use Together/Instead |
|----------|-------------------|
| [MATCH](./MATCH.md) | Find position to feed into INDEX - the classic combo |
| [XLOOKUP](./XLOOKUP.md) | Modern alternative (Excel 365+) - simpler syntax |
| [OFFSET](./OFFSET.md) | Similar but volatile (recalculates constantly) |
| [INDIRECT](./INDIRECT.md) | Create reference from text string |
| [CHOOSE](./CHOOSE.md) | Return value from a list by index |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Lookup & Reference Functions](./README.md) | [🏠 Back to Home](../../README.md)
