# VLOOKUP

## What It Does (Plain English)
Searches for a value in the first column of a table, then returns a value from any column in that same row. Think of it like looking up a person's name in a phone book and getting their phone number.

## Syntax
```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| lookup_value | Yes | The value to search for in the first column |
| table_array | Yes | The table containing your data |
| col_index_num | Yes | Which column to return the value from (1 = first column) |
| range_lookup | No | FALSE for exact match (recommended), TRUE for approximate match |

## Returns
The value from the specified column in the matching row, or an error if not found.

## Examples

### Example 1: Look Up Product Price
**Data (A1:C6):**
| A | B | C |
|---|---|---|
| Product ID | Product Name | Price |
| P001 | Laptop | 999 |
| P002 | Mouse | 29 |
| P003 | Keyboard | 79 |
| P004 | Monitor | 349 |
| P005 | Webcam | 89 |

**Formula:** `=VLOOKUP("P003", A2:C6, 3, FALSE)`

**Result:** `79`

**Explanation:** Searches for "P003" in the first column (A), finds it in row 4, returns the value from column 3 (Price) = 79.

---

### Example 2: Using Cell Reference for Lookup
**Same data as above, with lookup value in cell E2.**

**Formula:** `=VLOOKUP(E2, A2:C6, 2, FALSE)`

**If E2 = "P001"**

**Result:** `Laptop`

**Explanation:** Searches for the value in E2, returns column 2 (Product Name).

---

### Example 3: VLOOKUP with Error Handling
**Problem:** When the lookup value isn't found, you get an ugly #N/A error.

**Formula:** `=IFERROR(VLOOKUP(E2, A2:C6, 3, FALSE), "Product not found")`

**If E2 = "P999" (doesn't exist)**

**Result:** `Product not found`

**Explanation:** IFERROR catches the #N/A and returns your custom message instead.

---

### Example 4: Approximate Match (TRUE) for Grade Lookup
**Data (A1:B6):**
| A | B |
|---|---|
| Min Score | Grade |
| 0 | F |
| 60 | D |
| 70 | C |
| 80 | B |
| 90 | A |

**Formula:** `=VLOOKUP(75, A2:B6, 2, TRUE)`

**Result:** `C`

**Explanation:** With TRUE (approximate match), VLOOKUP finds the largest value ≤ 75, which is 70. **Important:** Data MUST be sorted ascending for approximate match!

---

### Example 5: Lookup from Another Sheet
**Data on sheet named "PriceList"**

**Formula:** `=VLOOKUP(A2, PriceList!A:C, 3, FALSE)`

**Result:** Returns price from column C of PriceList sheet where column A matches A2.

**Explanation:** Use `SheetName!Range` to reference data on other sheets.

---

### Example 6: Using Named Ranges for Clarity
**First, name your data range "Products" (A2:C100)**

**Formula:** `=VLOOKUP("P003", Products, 3, FALSE)`

**Result:** Same as Example 1, but much easier to read and maintain.

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #N/A | Value not found in first column | Check spelling, use IFERROR for graceful handling |
| #REF! | col_index_num is larger than table width | Reduce col_index_num or expand table_array |
| #VALUE! | col_index_num is less than 1 | Use positive integer for column number |
| Wrong result | Numbers vs text mismatch | Ensure lookup_value matches data type (both numbers or both text) |
| Wrong result with TRUE | Data not sorted | Sort first column ascending, or use FALSE |

## Pro Tips

- **Always use FALSE** for exact match unless you specifically need approximate matching (like tax tables or grade lookups)
- **VLOOKUP can only look RIGHT:** The lookup column must be the leftmost. For left lookups, use INDEX/MATCH or XLOOKUP
- **Column numbers are fragile:** If you insert a column, your col_index_num breaks. Consider using MATCH to find the column dynamically
- **Use tables:** If your data is in an Excel Table, use `=VLOOKUP("P003", Table1, 3, FALSE)` - the reference updates automatically
- **Wildcards:** Use * and ? for partial matches: `=VLOOKUP("*laptop*", A:C, 2, FALSE)` (only works with exact match)

## Limitations

1. **Cannot look left** - Lookup column must be leftmost
2. **Only returns first match** - If duplicate values exist, returns first found
3. **Column reference by number** - Prone to breaking when columns are inserted
4. **Case-insensitive** - "ABC" and "abc" are treated as the same
5. **Limited to one lookup value** - For multiple criteria, need helper column or use INDEX/MATCH

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [XLOOKUP](./XLOOKUP.md) | Modern replacement - can look in any direction, better error handling |
| [INDEX](./INDEX.md) + [MATCH](./MATCH.md) | More flexible, can look left, column-insert proof |
| [HLOOKUP](./HLOOKUP.md) | When your data is organized horizontally |
| [LOOKUP](./LOOKUP.md) | Simpler syntax for sorted data approximate matches |
| [FILTER](./FILTER.md) | When you need to return multiple matching rows |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible
- **Note:** Consider using XLOOKUP (Excel 365/2021) as the modern replacement

---
[📘 Back to Lookup & Reference Functions](./README.md) | [🏠 Back to Home](../../README.md)
