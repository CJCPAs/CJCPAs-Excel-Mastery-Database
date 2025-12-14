# Conditional Calculations Solutions

> **Sum, count, and average based on criteria**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Sum based on one condition | [SUMIF](#sum-with-one-condition) |
| Sum based on multiple conditions | [SUMIFS](#sum-with-multiple-conditions) |
| Count items matching criteria | [COUNTIF](#count-with-conditions) |
| Average with conditions | [AVERAGEIF/AVERAGEIFS](#average-with-conditions) |
| Find max/min with conditions | [MAXIFS/MINIFS](#max-min-with-conditions) |
| Sum across sheets | [3D References](#sum-across-sheets) |
| Sum visible cells only | [SUBTOTAL](#sum-visible-cells-only) |

---

## Sum with One Condition

### The Challenge
Sum only values that meet a specific criterion (like sum all sales for "North" region).

### Quick Answer
```excel
=SUMIF(criteria_range, criteria, sum_range)
```

### Full Example

**Data (A1:C6):**
| A | B | C |
|---|---|---|
| Region | Product | Sales |
| North | Widget | 1000 |
| South | Widget | 800 |
| North | Gadget | 1500 |
| South | Gadget | 1200 |
| North | Widget | 900 |

**Formula:** `=SUMIF(A2:A6, "North", C2:C6)`

**Result:** `3400`

**Explanation:** Sums Sales (column C) where Region (column A) equals "North" (1000 + 1500 + 900).

### Criteria Options

| Criteria | Meaning | Example |
|----------|---------|---------|
| "North" | Equals "North" | `=SUMIF(A:A, "North", B:B)` |
| ">100" | Greater than 100 | `=SUMIF(B:B, ">100", B:B)` |
| "<>0" | Not equal to 0 | `=SUMIF(B:B, "<>0", B:B)` |
| "*widget*" | Contains "widget" | `=SUMIF(A:A, "*widget*", B:B)` |
| A1 | Equals value in A1 | `=SUMIF(Region, A1, Sales)` |

---

## Sum with Multiple Conditions

### The Challenge
Sum values that meet multiple criteria (like North region AND Q1).

### Quick Answer
```excel
=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)
```

### Full Example

**Data (A1:D6):**
| A | B | C | D |
|---|---|---|---|
| Region | Quarter | Product | Sales |
| North | Q1 | Widget | 1000 |
| South | Q1 | Widget | 800 |
| North | Q2 | Gadget | 1500 |
| North | Q1 | Gadget | 1200 |
| South | Q2 | Widget | 900 |

**Formula:** `=SUMIFS(D2:D6, A2:A6, "North", B2:B6, "Q1")`

**Result:** `2200`

**Explanation:** Sums Sales where Region = "North" AND Quarter = "Q1" (1000 + 1200).

### Dynamic Criteria with Cell References

**Formula:** `=SUMIFS(D:D, A:A, F1, B:B, F2)`

Where F1 contains "North" and F2 contains "Q1". Changing these cells updates the result.

### Date Range Example
```excel
=SUMIFS(Sales, Dates, ">="&DATE(2024,1,1), Dates, "<="&DATE(2024,12,31))
```

---

## Count with Conditions

### The Challenge
Count how many cells match specific criteria.

### Quick Answer
```excel
=COUNTIF(range, criteria)                    // One condition
=COUNTIFS(range1, criteria1, range2, crit2)  // Multiple conditions
```

### Full Example

**Data (A1:B6):**
| A | B |
|---|---|
| Status | Amount |
| Complete | 500 |
| Pending | 300 |
| Complete | 750 |
| Failed | 200 |
| Complete | 400 |

**Count Complete:** `=COUNTIF(A2:A6, "Complete")`
**Result:** `3`

**Count Complete AND > 400:** `=COUNTIFS(A2:A6, "Complete", B2:B6, ">400")`
**Result:** `2`

### Count Unique Values
```excel
=SUMPRODUCT(1/COUNTIF(A2:A100, A2:A100))     // Pre-365
=ROWS(UNIQUE(A2:A100))                        // Excel 365
```

---

## Average with Conditions

### The Challenge
Calculate average only for values meeting criteria.

### Quick Answer
```excel
=AVERAGEIF(criteria_range, criteria, average_range)
=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
```

### Full Example

**Data:**
| A | B |
|---|---|
| Department | Salary |
| Sales | 55000 |
| Engineering | 75000 |
| Sales | 62000 |
| Engineering | 80000 |

**Formula:** `=AVERAGEIF(A2:A5, "Sales", B2:B5)`

**Result:** `58500`

**Explanation:** Average of Sales department salaries (55000 + 62000) / 2.

### Exclude Zeros from Average
```excel
=AVERAGEIF(B:B, "<>0")
```

---

## Max/Min with Conditions

### The Challenge
Find the largest or smallest value meeting criteria.

### Quick Answer (Excel 2019+)
```excel
=MAXIFS(max_range, criteria_range1, criteria1, ...)
=MINIFS(min_range, criteria_range1, criteria1, ...)
```

### Full Example

**Data:**
| A | B |
|---|---|
| Category | Value |
| A | 100 |
| B | 150 |
| A | 200 |
| B | 120 |

**Max for Category A:** `=MAXIFS(B2:B5, A2:A5, "A")`
**Result:** `200`

**Min for Category A:** `=MINIFS(B2:B5, A2:A5, "A")`
**Result:** `100`

### Pre-2019 Alternative
```excel
=MAX(IF(A2:A100="A", B2:B100))  // Enter with Ctrl+Shift+Enter
```

---

## Sum Across Sheets

### The Challenge
Sum the same cell or range from multiple worksheets.

### Quick Answer
```excel
=SUM(Sheet1:Sheet12!B5)        // Sum B5 from Sheet1 through Sheet12
=SUM(Jan:Dec!A1:A10)           // Sum A1:A10 from Jan through Dec sheets
```

### Full Example

**Sheets:** Jan, Feb, Mar, Apr (each has sales in B2:B10)

**Formula:** `=SUM(Jan:Apr!B2:B10)`

**Result:** Total of B2:B10 from all four sheets.

### Requirements
- Sheet names must be contiguous (in order in tab bar)
- Same cell/range reference applies to all sheets

---

## Sum Visible Cells Only

### The Challenge
Sum only cells that are visible (ignore filtered or hidden rows).

### Quick Answer
```excel
=SUBTOTAL(9, range)      // 9 = SUM, ignores manually hidden
=SUBTOTAL(109, range)    // 109 = SUM, ignores filtered AND hidden
=AGGREGATE(9, 5, range)  // More options available
```

### Full Example

**Data with filter applied:**
| A | B |
|---|---|
| Product | Sales |
| Widget | 100 |
| (hidden) | 200 |
| Gadget | 150 |

**Formula:** `=SUBTOTAL(9, B2:B4)`

**Result:** `250` (ignores the hidden row)

### SUBTOTAL Function Codes
| Code | Function | +100 for filtered |
|------|----------|-------------------|
| 1 | AVERAGE | 101 |
| 2 | COUNT | 102 |
| 3 | COUNTA | 103 |
| 4 | MAX | 104 |
| 5 | MIN | 105 |
| 9 | SUM | 109 |

---

## Running Totals

### The Challenge
Create a cumulative sum that grows as you go down the rows.

### Quick Answer
```excel
=SUM($A$2:A2)    // In row 2, copy down
```

### Full Example

**Data:**
| A | B |
|---|---|
| Amount | Running Total |
| 100 | =SUM($A$2:A2) → 100 |
| 50 | =SUM($A$2:A3) → 150 |
| 75 | =SUM($A$2:A4) → 225 |
| 125 | =SUM($A$2:A5) → 350 |

**Explanation:** The $ anchors row 2, but the second reference grows as you copy down.

---

## Weighted Calculations

### The Challenge
Calculate weighted average or weighted sum.

### Quick Answer
```excel
=SUMPRODUCT(values, weights) / SUM(weights)    // Weighted average
=SUMPRODUCT(values, weights)                    // Weighted sum
```

### Full Example - Grade Calculation

**Data:**
| A | B | C |
|---|---|---|
| Assignment | Score | Weight |
| Homework | 85 | 20% |
| Midterm | 78 | 30% |
| Final | 92 | 50% |

**Formula:** `=SUMPRODUCT(B2:B4, C2:C4)`

**Result:** `86.1`

**Explanation:** (85×0.20) + (78×0.30) + (92×0.50) = 86.1

---

## Percentage Calculations

### Percentage of Total
```excel
=A2/SUM($A$2:$A$10)
```
Format as percentage (Ctrl+Shift+%).

### Percentage Change
```excel
=(New-Old)/Old
=(B2-A2)/A2
```

### Percentage Difference
```excel
=ABS(A2-B2)/((A2+B2)/2)    // Percentage difference between two values
```

---

## Related Solutions

- [Lookups](../lookups/README.md) - Find values to use in calculations
- [Data Analysis](../data-analysis/README.md) - More complex analysis
- [Error Handling](../error-handling/README.md) - Handle calculation errors

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
