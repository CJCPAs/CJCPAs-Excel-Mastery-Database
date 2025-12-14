# SUMPRODUCT

## What It Does (Plain English)
Multiplies corresponding values in multiple arrays together, then adds up all those products. It's incredibly versatile and can act as a powerful conditional calculator.

## Syntax
```
=SUMPRODUCT(array1, [array2], [array3], ...)
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| array1 | Yes | The first array of values |
| array2 | No | Additional arrays (must be same dimensions as array1) |
| array3... | No | Up to 255 total arrays |

## Returns
The sum of the products of corresponding values in the arrays.

## Examples

### Example 1: Calculate Total Revenue (Quantity × Price)
**Data:**
| A | B | C |
|---|---|---|
| Product | Quantity | Price |
| Laptop | 5 | 999 |
| Mouse | 25 | 29 |
| Keyboard | 15 | 79 |
| Monitor | 8 | 349 |

**Formula:** `=SUMPRODUCT(B2:B5, C2:C5)`

**Result:** `10,522`

**Explanation:** (5×999) + (25×29) + (15×79) + (8×349) = 4995 + 725 + 1185 + 2792 = 10,522

---

### Example 2: Weighted Average
**Data:**
| A | B | C |
|---|---|---|
| Category | Score | Weight |
| Homework | 85 | 0.20 |
| Midterm | 78 | 0.30 |
| Final | 92 | 0.50 |

**Formula:** `=SUMPRODUCT(B2:B4, C2:C4)/SUM(C2:C4)`

**Result:** `86.1`

**Explanation:** (85×0.20 + 78×0.30 + 92×0.50) / (0.20 + 0.30 + 0.50) = 86.1

---

### Example 3: Conditional Sum (Like SUMIF but More Flexible)
**Data:**
| A | B | C |
|---|---|---|
| Region | Product | Sales |
| North | Widget | 1000 |
| South | Widget | 800 |
| North | Gadget | 1500 |
| North | Widget | 1200 |

**Formula:** `=SUMPRODUCT((A2:A5="North")*(B2:B5="Widget")*(C2:C5))`

**Result:** `2200`

**Explanation:** TRUE becomes 1, FALSE becomes 0. Only rows where both conditions are TRUE have a non-zero multiplier. (1×1×1000) + (0×1×800) + (1×0×1500) + (1×1×1200) = 2200

---

### Example 4: Count with Multiple Conditions (Like COUNTIFS)
**Data:**
| A | B |
|---|---|
| Status | Priority |
| Open | High |
| Closed | Low |
| Open | High |
| Open | Low |
| Closed | High |

**Formula:** `=SUMPRODUCT((A2:A6="Open")*(B2:B6="High"))`

**Result:** `2`

**Explanation:** Counts rows where Status is "Open" AND Priority is "High".

---

### Example 5: OR Logic with SUMPRODUCT
**Data:**
| A | B |
|---|---|
| Region | Sales |
| North | 1000 |
| South | 800 |
| East | 1200 |
| West | 950 |

**Formula:** `=SUMPRODUCT(((A2:A5="North")+(A2:A5="South"))*(B2:B5))`

**Result:** `1800`

**Explanation:** The + creates OR logic. Sums sales for North OR South (1000 + 800 = 1800).

---

### Example 6: Count Unique Values
**Data:**
| A |
|---|
| Apple |
| Banana |
| Apple |
| Cherry |
| Banana |
| Apple |

**Formula:** `=SUMPRODUCT(1/COUNTIF(A2:A7, A2:A7))`

**Result:** `3`

**Explanation:** Each unique value contributes 1 to the total. Apple appears 3 times, so each Apple contributes 1/3. Total: 3×(1/3) + 2×(1/2) + 1×(1/1) = 1 + 1 + 1 = 3.

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #VALUE! | Arrays are different sizes | Ensure all arrays have identical dimensions |
| #DIV/0! | Division by zero in unique count formula | Add IFERROR or filter out blanks first |
| Wrong result | Comparing numbers to text-numbers | Convert text to numbers or use VALUE() |
| 0 (unexpected) | Logical conditions not properly formed | Wrap conditions in parentheses |

## Pro Tips

- **Parentheses are crucial:** Each condition needs its own parentheses: `(A1:A10="X")*(B1:B10="Y")`
- **No array formula needed:** Unlike some array functions, SUMPRODUCT doesn't need Ctrl+Shift+Enter
- **Mixed conditions:** Combine AND (*) and OR (+) logic: `((A="X")+(A="Y"))*(B>100)`
- **Case-insensitive:** SUMPRODUCT comparisons are case-insensitive by default
- **For case-sensitive:** Use EXACT: `=SUMPRODUCT((EXACT(A1:A10,"test"))*1)`
- **Performance:** For simple conditions, SUMIFS is faster. Use SUMPRODUCT for complex logic

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [SUMIFS](./SUMIFS.md) | Simpler for basic AND conditions - faster too |
| [COUNTIFS](../statistical/COUNTIFS.md) | When just counting with multiple conditions |
| [MMULT](./MMULT.md) | For true matrix multiplication |
| [SUM](./SUM.md) | When no multiplication is needed |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Math & Trig Functions](./README.md) | [🏠 Back to Home](../../README.md)
