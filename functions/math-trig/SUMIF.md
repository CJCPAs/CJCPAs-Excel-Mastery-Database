# SUMIF

## What It Does (Plain English)
Adds up only the numbers that meet a specific condition you set - like saying "add all sales, but only from the North region."

## Syntax
```
=SUMIF(range, criteria, [sum_range])
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| range | Yes | The range of cells to check against your criteria |
| criteria | Yes | The condition that determines which cells to sum (number, text, expression, or cell reference) |
| sum_range | No | The cells to actually add up. If omitted, Excel sums the cells in 'range' |

## Returns
A single number representing the sum of all cells where the criteria was met.

## Examples

### Example 1: Sum Sales for a Specific Product
**Data:**
| A | B |
|---|---|
| Product | Sales |
| Laptop | 1200 |
| Mouse | 25 |
| Laptop | 800 |
| Monitor | 350 |
| Mouse | 30 |

**Formula:** `=SUMIF(A2:A6, "Laptop", B2:B6)`

**Result:** `2000`

**Explanation:** Looks for "Laptop" in column A, then sums the corresponding values in column B (1200 + 800 = 2000).

---

### Example 2: Sum Values Greater Than a Threshold
**Data:**
| A |
|---|
| Order Amount |
| 150 |
| 50 |
| 275 |
| 80 |
| 320 |

**Formula:** `=SUMIF(A2:A6, ">100")`

**Result:** `745`

**Explanation:** Adds only values greater than 100 (150 + 275 + 320 = 745). No sum_range needed because we're summing the same cells we're checking.

---

### Example 3: Sum Using Cell Reference as Criteria
**Data:**
| A | B | C |
|---|---|---|
| Region | Sales | Lookup Region |
| North | 5000 | North |
| South | 3000 | |
| North | 4500 | |
| East | 2800 | |

**Formula:** `=SUMIF(A2:A5, C1, B2:B5)`

**Result:** `9500`

**Explanation:** Uses the value in C1 ("North") as the criteria, summing sales where Region equals "North" (5000 + 4500 = 9500).

---

### Example 4: Sum with Wildcard for Partial Match
**Data:**
| A | B |
|---|---|
| Product Code | Revenue |
| PROD-001 | 500 |
| PROD-002 | 750 |
| SERVICE-001 | 200 |
| PROD-003 | 600 |

**Formula:** `=SUMIF(A2:A5, "PROD*", B2:B5)`

**Result:** `1850`

**Explanation:** The asterisk (*) wildcard matches any characters after "PROD", so it sums revenue for PROD-001, PROD-002, and PROD-003 (500 + 750 + 600 = 1850).

---

### Example 5: Sum Values Not Equal To
**Data:**
| A | B |
|---|---|
| Status | Amount |
| Paid | 100 |
| Pending | 250 |
| Paid | 175 |
| Cancelled | 50 |
| Pending | 300 |

**Formula:** `=SUMIF(A2:A6, "<>Paid", B2:B6)`

**Result:** `600`

**Explanation:** The `<>` operator means "not equal to", so this sums amounts where status is NOT "Paid" (250 + 50 + 300 = 600).

---

### Example 6: Sum Based on Dates
**Data:**
| A | B |
|---|---|
| Date | Sales |
| 2024-01-15 | 500 |
| 2024-02-20 | 750 |
| 2024-01-28 | 600 |
| 2024-03-05 | 800 |

**Formula:** `=SUMIF(A2:A5, ">=2024-02-01", B2:B5)`

**Result:** `1550`

**Explanation:** Sums sales for dates on or after February 1, 2024 (750 + 800 = 1550).

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #VALUE! | Criteria has more than 255 characters | Shorten criteria or use different approach |
| 0 (unexpected) | Criteria text doesn't match due to case differences | SUMIF is case-insensitive, check for extra spaces using TRIM |
| 0 (unexpected) | Numbers stored as text or vice versa | Ensure consistent data types in range |
| Wrong sum | Range and sum_range different sizes | Make sure both ranges have the same number of rows/columns |
| #NAME? | Criteria text not in quotes | Put text criteria in double quotes: "North" |

## Pro Tips

- **Wildcards are powerful:** Use `*` for any characters, `?` for single character. `"*Smith"` matches "John Smith", "Jane Smith"
- **Dynamic criteria:** Reference a cell for criteria `=SUMIF(A:A, F1, B:B)` lets users change the filter without editing the formula
- **Date comparisons:** Use comparison operators with dates: `">="&DATE(2024,1,1)` for dates on or after Jan 1, 2024
- **Approximate matches:** Use `>=`, `<=`, `>`, `<` for numeric ranges
- **Literal wildcards:** To find actual asterisk or question mark, use tilde: `"~*"` finds cells containing *
- **Empty cells:** Use `""` (empty quotes) to sum where criteria cells are blank, `"<>"` for non-blank

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [SUMIFS](./SUMIFS.md) | When you need multiple conditions (AND logic) |
| [SUM](./SUM.md) | When you want to sum everything without conditions |
| [COUNTIF](../statistical/COUNTIF.md) | When you want to count (not sum) based on criteria |
| [AVERAGEIF](../statistical/AVERAGEIF.md) | When you want the average of values meeting criteria |
| [SUMPRODUCT](./SUMPRODUCT.md) | When you need OR logic or more complex conditions |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Math & Trig Functions](./README.md) | [🏠 Back to Home](../../README.md)
