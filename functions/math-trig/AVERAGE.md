# AVERAGE

## What It Does (Plain English)
Calculates the arithmetic mean of a group of numbers - adds them all up and divides by how many there are.

## Syntax
```
=AVERAGE(number1, [number2], ...)
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| number1 | Yes | The first number, cell reference, or range to average |
| number2 | No | Additional numbers, cell references, or ranges (up to 255 total) |

## Returns
A single number representing the arithmetic mean of all values provided.

## Examples

### Example 1: Average of Test Scores
**Data:**
| A | B |
|---|---|
| Student | Score |
| Emma | 85 |
| James | 92 |
| Olivia | 78 |
| Liam | 88 |
| Sophia | 95 |

**Formula:** `=AVERAGE(B2:B6)`

**Result:** `87.6`

**Explanation:** (85 + 92 + 78 + 88 + 95) ÷ 5 = 438 ÷ 5 = 87.6

---

### Example 2: Average Ignores Text and Empty Cells
**Data:**
| A |
|---|
| 100 |
| 200 |
| N/A |
| 300 |
| (empty) |

**Formula:** `=AVERAGE(A1:A5)`

**Result:** `200`

**Explanation:** AVERAGE only considers 100, 200, and 300. The text "N/A" and empty cell are ignored. (100 + 200 + 300) ÷ 3 = 200.

---

### Example 3: Average Multiple Ranges
**Data:**
| A | B | C |
|---|---|---|
| Q1 | Q2 | Q3 |
| 1500 | 1800 | 2100 |
| 1200 | 1600 | 1900 |

**Formula:** `=AVERAGE(A2:C2, A3:C3)`

**Result:** `1683.33`

**Explanation:** Averages all 6 values from both rows: (1500 + 1800 + 2100 + 1200 + 1600 + 1900) ÷ 6 = 1683.33

---

### Example 4: Average with Direct Numbers
**Formula:** `=AVERAGE(10, 20, 30, 40, 50)`

**Result:** `30`

**Explanation:** You can pass numbers directly to AVERAGE: (10 + 20 + 30 + 40 + 50) ÷ 5 = 30

---

### Example 5: Weighted Average (Alternative Approach)
**Data:**
| A | B | C |
|---|---|---|
| Assignment | Score | Weight |
| Homework | 85 | 20% |
| Midterm | 78 | 30% |
| Final | 92 | 50% |

**Note:** AVERAGE treats all values equally. For weighted average, use SUMPRODUCT:

**Formula:** `=SUMPRODUCT(B2:B4, C2:C4)/SUM(C2:C4)`

**Result:** `86.1`

**Explanation:** (85×0.20 + 78×0.30 + 92×0.50) ÷ (0.20 + 0.30 + 0.50) = 86.1

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #DIV/0! | Range contains no numbers (only text or empty cells) | Ensure at least one numeric value exists |
| Wrong average | Zeros included when they shouldn't be | Use AVERAGEIF to exclude zeros |
| Result seems high/low | Numbers stored as text aren't counted | Convert text to numbers (Data > Text to Columns) |
| #VALUE! | Range includes an error value | Use AGGREGATE(1,6,range) to average ignoring errors |

## Pro Tips

- **Zeros ARE counted:** Unlike blank cells, zero values count toward the average. Use `=AVERAGEIF(A:A,"<>0")` to exclude zeros
- **AVERAGE vs AVERAGEA:** AVERAGE ignores text; AVERAGEA treats text as 0 and TRUE/FALSE as 1/0
- **Quick average:** Select cells and look at the status bar at the bottom of Excel - it shows the average
- **Median vs Average:** For data with outliers, MEDIAN might be more representative than AVERAGE
- **Precision:** Excel calculates to 15 significant digits; what you see depends on cell formatting

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [AVERAGEIF](../statistical/AVERAGEIF.md) | When you need to average only values meeting a condition |
| [AVERAGEIFS](../statistical/AVERAGEIFS.md) | When you need multiple conditions |
| [MEDIAN](../statistical/MEDIAN.md) | When you want the middle value (better for outliers) |
| [AVERAGEA](../statistical/AVERAGEA.md) | When you want text treated as 0 and TRUE/FALSE counted |
| [TRIMMEAN](../statistical/TRIMMEAN.md) | When you want to exclude extreme values from the average |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Math & Trig Functions](./README.md) | [🏠 Back to Home](../../README.md)
