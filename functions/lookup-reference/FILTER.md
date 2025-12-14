# FILTER

## What It Does (Plain English)
Returns only the rows (or columns) from a range that meet your specified conditions - like using a coffee filter to keep only what you want.

## Syntax
```
=FILTER(array, include, [if_empty])
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| array | Yes | The range of data to filter |
| include | Yes | A TRUE/FALSE condition that determines which rows to keep |
| if_empty | No | What to return if no rows match (otherwise returns #CALC! error) |

## Returns
An array of rows/columns that match your criteria. Results "spill" into adjacent cells automatically.

## Examples

### Example 1: Filter Sales by Region
**Data (A1:C6):**
| A | B | C |
|---|---|---|
| Region | Product | Sales |
| North | Widget | 1000 |
| South | Gadget | 800 |
| North | Gadget | 1200 |
| East | Widget | 950 |
| North | Widget | 1100 |

**Formula:** `=FILTER(A2:C6, A2:A6="North")`

**Result:**
| | | |
|---|---|---|
| North | Widget | 1000 |
| North | Gadget | 1200 |
| North | Widget | 1100 |

**Explanation:** Returns all rows where Region = "North". The result spills into 3 rows × 3 columns.

---

### Example 2: Filter with Multiple Criteria (AND)
**Same data as above**

**Formula:** `=FILTER(A2:C6, (A2:A6="North")*(B2:B6="Widget"))`

**Result:**
| | | |
|---|---|---|
| North | Widget | 1000 |
| North | Widget | 1100 |

**Explanation:** Multiplying conditions creates AND logic. TRUE × TRUE = 1 (include), TRUE × FALSE = 0 (exclude).

---

### Example 3: Filter with Multiple Criteria (OR)
**Formula:** `=FILTER(A2:C6, (A2:A6="North")+(A2:A6="East"))`

**Result:**
| | | |
|---|---|---|
| North | Widget | 1000 |
| North | Gadget | 1200 |
| East | Widget | 950 |
| North | Widget | 1100 |

**Explanation:** Adding conditions creates OR logic. Returns North OR East regions.

---

### Example 4: Filter by Numeric Condition
**Formula:** `=FILTER(A2:C6, C2:C6>1000)`

**Result:**
| | | |
|---|---|---|
| North | Gadget | 1200 |
| North | Widget | 1100 |

**Explanation:** Returns only rows where Sales > 1000.

---

### Example 5: Handle Empty Results
**Formula:** `=FILTER(A2:C6, A2:A6="West", "No data found")`

**Result:** `No data found`

**Explanation:** No rows have Region = "West", so the if_empty message is displayed instead of an error.

---

### Example 6: Filter + Sort Combination
**Formula:** `=SORT(FILTER(A2:C6, C2:C6>900), 3, -1)`

**Result:**
| | | |
|---|---|---|
| North | Gadget | 1200 |
| North | Widget | 1100 |
| North | Widget | 1000 |
| East | Widget | 950 |

**Explanation:** First filters for Sales > 900, then sorts by column 3 (Sales) in descending order.

---

### Example 7: Filter and Return Specific Columns
**Formula:** `=FILTER(B2:C6, A2:A6="North")`

**Result:**
| | |
|---|---|
| Widget | 1000 |
| Gadget | 1200 |
| Widget | 1100 |

**Explanation:** Filter condition uses column A, but returns only columns B and C.

---

### Example 8: Top N with FILTER
**Get top 3 sales:**

**Formula:** `=FILTER(A2:C6, C2:C6>=LARGE(C2:C6, 3))`

**Result:** Returns all rows where Sales is in top 3 values.

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #CALC! | No rows match criteria | Add if_empty argument: `FILTER(data, cond, "None")` |
| #SPILL! | Cells where results would appear aren't empty | Clear the cells below/beside your formula |
| #VALUE! | Include array different size than data array | Ensure the criteria range matches the data rows |
| Results wrong | Criteria comparing wrong types | Ensure numbers compared to numbers, text to text |

## Pro Tips

- **Dynamic lists:** FILTER is perfect for creating dropdown-driven reports
- **With UNIQUE:** `=UNIQUE(FILTER(A:A, B:B="Active"))` for unique active items
- **Count matches:** `=ROWS(FILTER(A:A, B:B="X"))` counts how many match
- **Single column filter:** `=FILTER(A2:A100, B2:B100>50)` returns just the names where B > 50
- **Combine with other functions:** Works great with SORT, SORTBY, TAKE, DROP

## Related Functions

| Function | When to Use Together/Instead |
|----------|-------------------|
| [SORT](./SORT.md) | Sort the filtered results |
| [UNIQUE](./UNIQUE.md) | Get unique values from filtered results |
| [COUNTIF](../statistical/COUNTIF.md) | When you just need a count, not the actual data |
| [XLOOKUP](./XLOOKUP.md) | When you need just one matching value |
| [TAKE](./TAKE.md) | Get first/last N rows from filtered results |

## Version Notes
- **Available in:** Excel 365, Excel 2021
- **NOT available in:** Excel 2019 and earlier (use Advanced Filter or helper columns)
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ❌ Not available

---
[📘 Back to Lookup & Reference Functions](./README.md) | [🏠 Back to Home](../../README.md)
