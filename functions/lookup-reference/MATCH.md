# MATCH

## What It Does (Plain English)
Finds the position of a value in a list and tells you what row or column number it's in. Like finding what page number a word appears on in a book's index.

## Syntax
```
=MATCH(lookup_value, lookup_array, [match_type])
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| lookup_value | Yes | The value to search for |
| lookup_array | Yes | The one-row or one-column range to search in |
| match_type | No | 0 = exact match, 1 = largest value ≤ lookup (default), -1 = smallest value ≥ lookup |

## Returns
A number representing the relative position of the match in the lookup_array (1 for first item, 2 for second, etc.).

## Examples

### Example 1: Find Position in a List
**Data (A1:A5):**
| A |
|---|
| Apple |
| Banana |
| Cherry |
| Date |
| Elderberry |

**Formula:** `=MATCH("Cherry", A1:A5, 0)`

**Result:** `3`

**Explanation:** "Cherry" is the 3rd item in the list.

---

### Example 2: MATCH with INDEX for Lookups
**Data (A1:C5):**
| A | B | C |
|---|---|---|
| ID | Name | Salary |
| E001 | Alice | 75000 |
| E002 | Bob | 68000 |
| E003 | Carol | 82000 |
| E004 | Dave | 71000 |

**Formula:** `=INDEX(C2:C5, MATCH("E003", A2:A5, 0))`

**Result:** `82000`

**Explanation:** MATCH returns 3 (E003 is in position 3), INDEX returns the 3rd value from salaries.

---

### Example 3: Find Column Position
**Data (A1:D1):**
| A | B | C | D |
|---|---|---|---|
| Q1 | Q2 | Q3 | Q4 |

**Formula:** `=MATCH("Q3", A1:D1, 0)`

**Result:** `3`

**Explanation:** "Q3" is in the 3rd column of the range.

---

### Example 4: Approximate Match - Find Tax Bracket
**Data (A1:B5) - Tax brackets:**
| A | B |
|---|---|
| Income | Rate |
| 0 | 10% |
| 10000 | 12% |
| 40000 | 22% |
| 85000 | 24% |

**Formula:** `=MATCH(50000, A2:A5, 1)`

**Result:** `3`

**Explanation:** match_type 1 finds the largest value ≤ 50000, which is 40000 at position 3. **Data must be sorted ascending!**

---

### Example 5: Find Minimum Value Position
**Data (B1:B5):**
| B |
|---|
| 45 |
| 23 |
| 67 |
| 12 |
| 56 |

**Formula:** `=MATCH(MIN(B1:B5), B1:B5, 0)`

**Result:** `4`

**Explanation:** MIN returns 12, MATCH finds 12 at position 4.

---

### Example 6: Case-Insensitive (Default Behavior)
**Data (A1:A3):**
| A |
|---|
| APPLE |
| Banana |
| cherry |

**Formula:** `=MATCH("apple", A1:A3, 0)`

**Result:** `1`

**Explanation:** MATCH is case-insensitive. "apple" matches "APPLE".

---

### Example 7: Wildcard Match
**Data (A1:A4):**
| A |
|---|
| Product-A |
| Product-B |
| Item-A |
| Product-C |

**Formula:** `=MATCH("Product*", A1:A4, 0)`

**Result:** `1`

**Explanation:** The asterisk wildcard matches any characters. Returns first match starting with "Product".

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #N/A | Value not found | Check spelling, ensure value exists in range |
| #N/A with match_type 1 or -1 | Data not sorted properly | Sort data ascending for 1, descending for -1, or use 0 |
| Wrong result | Multiple matches | MATCH returns the first match only |
| Wrong position | Range starts at different row | Remember: MATCH returns relative position in the array, not the worksheet row |

## Pro Tips

- **Always use 0 for exact match:** The default (1) requires sorted data and often gives unexpected results
- **Position ≠ row number:** `=MATCH(x, A5:A20, 0)` returns position within A5:A20, not the actual row. Add the starting row -1 if you need the actual row
- **Wildcards work with 0:** Use `*` for any characters, `?` for single character
- **Combine with INDEX:** This is the most powerful and flexible lookup method
- **Find last occurrence:** For last match, reverse your data or use more complex formulas

## Match Types Explained

| match_type | Finds | Data Must Be |
|------------|-------|--------------|
| 0 | Exact match | Any order |
| 1 (default) | Largest value ≤ lookup_value | Sorted ascending |
| -1 | Smallest value ≥ lookup_value | Sorted descending |

## Related Functions

| Function | When to Use Together/Instead |
|----------|-------------------|
| [INDEX](./INDEX.md) | Use with MATCH to return values (INDEX/MATCH combo) |
| [XMATCH](./XMATCH.md) | Modern version with more options (Excel 365+) |
| [XLOOKUP](./XLOOKUP.md) | Combines MATCH + INDEX functionality |
| [SEARCH](../text/SEARCH.md) | Find position of text within a string |
| [FIND](../text/FIND.md) | Case-sensitive text position |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Lookup & Reference Functions](./README.md) | [🏠 Back to Home](../../README.md)
