# SUMIFS

## What It Does (Plain English)
Adds up numbers that meet multiple conditions at once - like saying "add all sales from the North region AND from January AND over $100."

## Syntax
```
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| sum_range | Yes | The range of cells to add up |
| criteria_range1 | Yes | The first range to check against criteria1 |
| criteria1 | Yes | The condition for criteria_range1 |
| criteria_range2 | No | Additional range to check (up to 127 pairs) |
| criteria2 | No | Additional condition for criteria_range2 |

## Returns
A single number representing the sum of all cells where ALL criteria were met.

## Examples

### Example 1: Sum Sales by Region and Product
**Data:**
| A | B | C |
|---|---|---|
| Region | Product | Sales |
| North | Laptop | 1200 |
| South | Laptop | 950 |
| North | Mouse | 45 |
| North | Laptop | 1350 |
| South | Mouse | 38 |

**Formula:** `=SUMIFS(C2:C6, A2:A6, "North", B2:B6, "Laptop")`

**Result:** `2550`

**Explanation:** Sums Sales (column C) where Region is "North" AND Product is "Laptop" (1200 + 1350 = 2550).

---

### Example 2: Sum Within a Date Range
**Data:**
| A | B | C |
|---|---|---|
| Date | Category | Amount |
| 2024-01-15 | Supplies | 250 |
| 2024-02-03 | Supplies | 175 |
| 2024-01-22 | Equipment | 890 |
| 2024-02-28 | Supplies | 320 |
| 2024-03-10 | Supplies | 200 |

**Formula:** `=SUMIFS(C2:C6, A2:A6, ">=2024-01-01", A2:A6, "<=2024-02-28", B2:B6, "Supplies")`

**Result:** `745`

**Explanation:** Sums Amount for dates between Jan 1 and Feb 28, 2024, where Category is "Supplies" (250 + 175 + 320 = 745).

---

### Example 3: Sum with Numeric Conditions
**Data:**
| A | B | C |
|---|---|---|
| Salesperson | Units Sold | Commission |
| Alice | 45 | 450 |
| Bob | 32 | 320 |
| Carol | 58 | 580 |
| Alice | 67 | 670 |
| Bob | 41 | 410 |

**Formula:** `=SUMIFS(C2:C6, A2:A6, "Alice", B2:B6, ">=50")`

**Result:** `670`

**Explanation:** Sums Commission for Alice where Units Sold is 50 or more. Only Alice's second entry (67 units) qualifies.

---

### Example 4: Using Cell References for Dynamic Criteria
**Data:**
| A | B | C | D | E |
|---|---|---|---|---|
| Region | Month | Sales | Filter Region | Filter Month |
| North | Jan | 5000 | North | Jan |
| North | Feb | 5500 | | |
| South | Jan | 4200 | | |
| South | Feb | 4800 | | |

**Formula:** `=SUMIFS(C2:C5, A2:A5, D1, B2:B5, E1)`

**Result:** `5000`

**Explanation:** Uses D1 ("North") and E1 ("Jan") as dynamic filters. Changing these cells instantly updates the sum.

---

### Example 5: Sum Excluding Specific Values
**Data:**
| A | B | C |
|---|---|---|
| Product | Status | Revenue |
| Widget | Active | 1000 |
| Gadget | Discontinued | 500 |
| Widget | Active | 1200 |
| Gizmo | Active | 800 |
| Widget | Discontinued | 300 |

**Formula:** `=SUMIFS(C2:C6, B2:B6, "<>Discontinued", A2:A6, "Widget")`

**Result:** `2200`

**Explanation:** Sums Revenue for Widget products that are NOT Discontinued (1000 + 1200 = 2200).

---

### Example 6: Sum with Wildcard Patterns
**Data:**
| A | B | C |
|---|---|---|
| SKU | Warehouse | Quantity |
| ELEC-001 | Main | 150 |
| FURN-002 | Main | 75 |
| ELEC-003 | Branch | 200 |
| ELEC-004 | Main | 180 |

**Formula:** `=SUMIFS(C2:C5, A2:A5, "ELEC*", B2:B5, "Main")`

**Result:** `330`

**Explanation:** Sums Quantity for SKUs starting with "ELEC" in "Main" warehouse (150 + 180 = 330).

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #VALUE! | Ranges are different sizes | Ensure all criteria_ranges and sum_range have the same dimensions |
| 0 (unexpected) | No rows match all criteria | Verify your criteria - remember it's AND logic, not OR |
| 0 (unexpected) | Text/number mismatch | Check if numbers are stored as text in your data |
| #NAME? | Unquoted text criteria | Put text in quotes: "North" not North |
| Wrong result | Criteria ranges misaligned with sum_range | All ranges must start and end at corresponding rows |

## Pro Tips

- **Order matters for readability, not results:** Criteria order doesn't affect the calculation, but putting the most restrictive first makes formulas easier to understand
- **Same range, different criteria:** You can use the same range multiple times: `=SUMIFS(C:C, A:A, ">=10", A:A, "<=20")` for a range between 10 and 20
- **Blank cells:** Use `""` for blank, `"<>"` for non-blank
- **OR logic workaround:** SUMIFS only does AND. For OR, add multiple SUMIFS: `=SUMIFS(C:C,A:A,"North")+SUMIFS(C:C,A:A,"South")`
- **Performance:** Name your ranges for cleaner formulas and use specific ranges (A2:A1000) instead of entire columns (A:A) for better performance

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [SUMIF](./SUMIF.md) | When you only have one condition |
| [COUNTIFS](../statistical/COUNTIFS.md) | When you want to count rows meeting multiple criteria |
| [AVERAGEIFS](../statistical/AVERAGEIFS.md) | When you want the average of values meeting multiple criteria |
| [MAXIFS](../statistical/MAXIFS.md) | When you want the maximum value meeting criteria |
| [MINIFS](../statistical/MINIFS.md) | When you want the minimum value meeting criteria |
| [SUMPRODUCT](./SUMPRODUCT.md) | When you need OR logic or more complex calculations |

## Version Notes
- **Available in:** Excel 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Not available in:** Excel 2003 (use SUMPRODUCT instead)
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Math & Trig Functions](./README.md) | [🏠 Back to Home](../../README.md)
