# COUNT

## What It Does (Plain English)
Counts how many cells contain numbers - ignores text, blanks, and errors.

## Syntax
```
=COUNT(value1, [value2], ...)
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| value1 | Yes | The first cell, range, or value to count |
| value2 | No | Additional cells, ranges, or values to count (up to 255 total) |

## Returns
A single number representing how many numeric values were found.

## Examples

### Example 1: Count Numbers in a List
**Data:**
| A |
|---|
| 100 |
| 200 |
| Hello |
| 300 |
| (blank) |
| 400 |

**Formula:** `=COUNT(A1:A6)`

**Result:** `4`

**Explanation:** Counts only the numeric values: 100, 200, 300, and 400. Ignores "Hello" and the blank cell.

---

### Example 2: Count Across Multiple Columns
**Data:**
| A | B | C |
|---|---|---|
| 10 | 20 | Text |
| 30 | N/A | 40 |

**Formula:** `=COUNT(A1:C2)`

**Result:** `4`

**Explanation:** Counts 10, 20, 30, and 40. Ignores "Text" and "N/A".

---

### Example 3: Count with Mixed Arguments
**Formula:** `=COUNT(1, 2, "three", 4, TRUE)`

**Result:** `3`

**Explanation:** Counts only 1, 2, and 4. Text "three" and logical TRUE are not counted as numbers.

---

### Example 4: Count to Find Data Entries
**Data:**
| A | B |
|---|---|
| Employee | Sales |
| Alice | 5000 |
| Bob | 4200 |
| Carol | 6100 |
| Dave | 3800 |

**Formula:** `=COUNT(B2:B5)`

**Result:** `4`

**Explanation:** Useful for determining how many employees have sales recorded.

---

### Example 5: Dynamic Row Counter
**Scenario:** Create a running count that updates as you add data.

**Data:**
| A | B |
|---|---|
| Entry | Count |
| 100 | =COUNT($A$2:A2) → 1 |
| 200 | =COUNT($A$2:A3) → 2 |
| (blank) | =COUNT($A$2:A4) → 2 |
| 300 | =COUNT($A$2:A5) → 3 |

**Explanation:** The anchor on $A$2 means the range expands as you copy down, counting only numeric entries.

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| 0 when expecting count | Numbers stored as text | Convert to numbers using VALUE() or Text to Columns |
| Unexpected count | Dates/times are counted | Dates are numbers in Excel, so they're counted |
| Too high count | Hidden cells included | COUNT includes hidden rows/columns |

## Pro Tips

- **Dates are numbers:** Excel stores dates as numbers, so COUNT will include dates
- **Logical values in cells aren't counted:** TRUE/FALSE in cells aren't counted, but in array formulas they may be treated as 1/0
- **Quick count:** Select a range and look at the status bar - it shows "Count" for selected values
- **Difference between functions:**
  - COUNT = counts numbers only
  - COUNTA = counts non-empty cells (any content)
  - COUNTBLANK = counts empty cells

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [COUNTA](../statistical/COUNTA.md) | When you want to count any non-empty cell (including text) |
| [COUNTBLANK](../statistical/COUNTBLANK.md) | When you want to count empty cells |
| [COUNTIF](../statistical/COUNTIF.md) | When you want to count cells meeting a condition |
| [COUNTIFS](../statistical/COUNTIFS.md) | When you have multiple conditions |
| [ROWS](../lookup-reference/ROWS.md) | When you want to count rows in a range (regardless of content) |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Math & Trig Functions](./README.md) | [🏠 Back to Home](../../README.md)
