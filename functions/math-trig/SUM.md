# SUM

## What It Does (Plain English)
Adds up all the numbers you give it - like having a calculator that instantly totals any list of values you point to.

## Syntax
```
=SUM(number1, [number2], [number3], ...)
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| number1 | Yes | The first number, cell reference, or range to add |
| number2 | No | Additional numbers, cell references, or ranges to add (up to 255 total) |

## Returns
A single number representing the total of all values provided.

## Examples

### Example 1: Sum a Column of Sales
**Data:**
| A | B |
|---|---|
| Product | Sales |
| Laptop | 1200 |
| Mouse | 25 |
| Keyboard | 75 |
| Monitor | 350 |

**Formula:** `=SUM(B2:B5)`

**Result:** `1650`

**Explanation:** The formula adds all values in cells B2 through B5 (1200 + 25 + 75 + 350 = 1650).

---

### Example 2: Sum Multiple Separate Ranges
**Data:**
| A | B | C | D |
|---|---|---|---|
| Q1 Sales | 5000 | Q2 Sales | 6200 |
| Q3 Sales | 5800 | Q4 Sales | 7100 |

**Formula:** `=SUM(B1, D1, B2, D2)`

**Result:** `24100`

**Explanation:** Adds individual cells from different locations: 5000 + 6200 + 5800 + 7100 = 24100.

---

### Example 3: Sum with Direct Numbers and Cell References
**Data:**
| A | B |
|---|---|
| Base Price | 100 |
| Tax | 8 |

**Formula:** `=SUM(B1, B2, 50)`

**Result:** `158`

**Explanation:** Combines cell values (100 + 8) with a direct number (50) to get 158.

---

### Example 4: Sum Entire Column (Dynamic)
**Scenario:** You have sales data that grows over time and want to always sum the entire column.

**Formula:** `=SUM(A:A)`

**Result:** Sum of all numbers in column A, automatically including new entries.

**Explanation:** Using `A:A` references the entire column, so new data added to column A is automatically included.

---

### Example 5: Sum Non-Contiguous Ranges
**Data:**
| A | B | C | D | E |
|---|---|---|---|---|
| Region | Jan | Feb | Mar | Apr |
| North | 1000 | 1100 | 1200 | 1300 |
| South | 800 | 900 | 950 | 1000 |

**Formula:** `=SUM(B2:B3, D2:D3)`

**Result:** `4150`

**Explanation:** Adds January totals (1000 + 800 = 1800) and March totals (1200 + 950 = 2150), giving 4150.

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #VALUE! | Text included in range that can't be interpreted as a number | Use SUMIF to filter out text, or clean your data |
| Result shows 0 | Numbers stored as text | Select cells, use Data > Text to Columns > Finish to convert |
| Formula shows instead of result | Cell formatted as text | Format cell as Number, then re-enter formula |
| Unexpected result | Hidden rows included in sum | Use SUBTOTAL(9,...) to ignore hidden rows |

## Pro Tips

- **SUM ignores text and logical values** - If your range contains "N/A" or TRUE/FALSE, these are skipped (not an error, just ignored)
- **Use named ranges for clarity** - Instead of `=SUM(B2:B50)`, use `=SUM(MonthlySales)` after naming your range
- **SUM is faster than adding cells** - `=SUM(A1:A1000)` calculates faster than `=A1+A2+A3+...+A1000`
- **Combine with other functions** - `=SUM(A:A)/COUNT(A:A)` gives average (though AVERAGE is simpler)
- **3D References** - Sum across sheets: `=SUM(Sheet1:Sheet12!B5)` adds cell B5 from all 12 sheets

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [SUMIF](./SUMIF.md) | When you need to sum only cells that meet a condition |
| [SUMIFS](./SUMIFS.md) | When you need multiple conditions for your sum |
| [SUMPRODUCT](./SUMPRODUCT.md) | When you need to multiply ranges and then sum the products |
| [SUBTOTAL](./SUBTOTAL.md) | When you need to sum visible cells only (works with filters) |
| [AGGREGATE](./AGGREGATE.md) | When you need to ignore errors or hidden rows |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Math & Trig Functions](./README.md) | [🏠 Back to Home](../../README.md)
