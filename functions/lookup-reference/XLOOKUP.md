# XLOOKUP

## What It Does (Plain English)
Searches for a value anywhere in a column (or row), then returns a corresponding value from another column (or row). It's the modern, powerful replacement for VLOOKUP that can look in any direction.

## Syntax
```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| lookup_value | Yes | The value to search for |
| lookup_array | Yes | The column or row to search in |
| return_array | Yes | The column or row to return values from |
| if_not_found | No | Value to return if no match (instead of #N/A) |
| match_mode | No | 0 = exact (default), -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard |
| search_mode | No | 1 = first to last (default), -1 = last to first, 2 = binary ascending, -2 = binary descending |

## Returns
The value(s) from return_array corresponding to the match found in lookup_array.

## Examples

### Example 1: Basic Lookup (Replacing VLOOKUP)
**Data (A1:C6):**
| A | B | C |
|---|---|---|
| ID | Name | Salary |
| E001 | Alice | 75000 |
| E002 | Bob | 68000 |
| E003 | Carol | 82000 |
| E004 | Dave | 71000 |
| E005 | Eve | 79000 |

**Formula:** `=XLOOKUP("E003", A2:A6, C2:C6)`

**Result:** `82000`

**Explanation:** Finds "E003" in column A, returns corresponding value from column C.

---

### Example 2: Built-in Error Handling
**Formula:** `=XLOOKUP("E999", A2:A6, C2:C6, "Employee not found")`

**Result:** `Employee not found`

**Explanation:** The fourth argument specifies what to return if no match - no IFERROR needed!

---

### Example 3: Return Multiple Columns (Entire Row)
**Formula:** `=XLOOKUP("E003", A2:A6, B2:C6)`

**Result:** Returns both columns: `Carol` and `82000` (spills to adjacent cells)

**Explanation:** When return_array has multiple columns, XLOOKUP returns all of them.

---

### Example 4: Lookup Left (Not Possible with VLOOKUP)
**Data:**
| A | B | C |
|---|---|---|
| Name | ID | Department |
| Alice | E001 | Sales |
| Bob | E002 | Marketing |

**Formula:** `=XLOOKUP("E001", B2:B3, A2:A3)`

**Result:** `Alice`

**Explanation:** Searches in column B, returns from column A (to the LEFT). VLOOKUP can't do this!

---

### Example 5: Find Last Match (Reverse Search)
**Data:**
| A | B |
|---|---|
| Date | Transaction |
| 2024-01-15 | Purchase |
| 2024-02-20 | Purchase |
| 2024-03-05 | Return |
| 2024-04-10 | Purchase |

**Formula:** `=XLOOKUP("Purchase", B2:B5, A2:A5, , 0, -1)`

**Result:** `2024-04-10`

**Explanation:** The search_mode of -1 searches from bottom to top, finding the last "Purchase".

---

### Example 6: Two-Way Lookup (Nested XLOOKUP)
**Data (A1:D4):**
| | Jan | Feb | Mar |
|---|---|---|---|
| North | 100 | 120 | 150 |
| South | 80 | 95 | 110 |
| East | 90 | 100 | 130 |

**Formula:** `=XLOOKUP("South", A2:A4, XLOOKUP("Feb", B1:D1, B2:D4))`

**Result:** `95`

**Explanation:** Inner XLOOKUP finds the "Feb" column, outer XLOOKUP finds the "South" row within that column.

---

### Example 7: Approximate Match (Greater Than or Equal)
**Data - Shipping Rates:**
| A | B |
|---|---|
| Weight (lbs) | Rate |
| 0 | 5.99 |
| 5 | 8.99 |
| 10 | 12.99 |
| 20 | 18.99 |

**Formula:** `=XLOOKUP(7, A2:A5, B2:B5, , -1)`

**Result:** `8.99`

**Explanation:** match_mode -1 finds exact match or next smaller value. 7 lbs uses the 5-lb rate.

---

### Example 8: Wildcard Search
**Data:**
| A | B |
|---|---|
| Product | Price |
| iPhone 13 | 799 |
| iPhone 14 | 899 |
| Galaxy S22 | 749 |

**Formula:** `=XLOOKUP("*14*", A2:A4, B2:B4, , 2)`

**Result:** `899`

**Explanation:** match_mode 2 enables wildcards. Finds first product containing "14".

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #N/A | No match found | Use if_not_found argument |
| #VALUE! | lookup_array and return_array different sizes | Ensure same row/column count |
| #SPILL! | Return area not empty | Clear cells where results will spill |
| #CALC! | Arrays too large | Reduce range sizes |

## Pro Tips

- **Always use if_not_found:** `=XLOOKUP(x, a, b, "Not found")` is cleaner than IFERROR
- **Return multiple columns:** Return arrays expand to adjacent cells automatically
- **Binary search for large data:** Use search_mode 2 or -2 for faster lookups on sorted data
- **Case insensitive by default:** Just like VLOOKUP. For case-sensitive, combine with EXACT()
- **Handles #N/A gracefully:** Unlike VLOOKUP, you can specify exactly what happens on no match

## Why XLOOKUP is Better Than VLOOKUP

| Feature | XLOOKUP | VLOOKUP |
|---------|---------|---------|
| Look left | ✅ Yes | ❌ No |
| Error handling | ✅ Built-in | ❌ Need IFERROR |
| Column reference | ✅ Actual column | ❌ Number only |
| Return multiple columns | ✅ Yes | ❌ No |
| Last match search | ✅ Yes | ❌ No |
| Default match type | ✅ Exact | ❌ Approximate |
| Insert columns safe | ✅ Yes | ❌ No |

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [VLOOKUP](./VLOOKUP.md) | Legacy spreadsheets or older Excel versions |
| [INDEX](./INDEX.md) + [MATCH](./MATCH.md) | When XLOOKUP isn't available |
| [FILTER](./FILTER.md) | When you need all matching rows, not just first |
| [XMATCH](./XMATCH.md) | When you only need the position, not the value |

## Version Notes
- **Available in:** Excel 365, Excel 2021
- **NOT available in:** Excel 2019 and earlier (use INDEX/MATCH)
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ❌ Not available (use VLOOKUP or INDEX/MATCH)

---
[📘 Back to Lookup & Reference Functions](./README.md) | [🏠 Back to Home](../../README.md)
