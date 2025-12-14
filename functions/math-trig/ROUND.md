# ROUND

## What It Does (Plain English)
Rounds a number to a specific number of decimal places using standard rounding rules (5 and above rounds up).

## Syntax
```
=ROUND(number, num_digits)
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| number | Yes | The number you want to round |
| num_digits | Yes | The number of decimal places to round to. Can be negative to round to tens, hundreds, etc. |

## Returns
The number rounded to the specified number of digits.

## Examples

### Example 1: Round to 2 Decimal Places (Currency)
**Data:**
| A | B |
|---|---|
| Price | Rounded |
| 19.9549 | =ROUND(A2, 2) |

**Formula:** `=ROUND(19.9549, 2)`

**Result:** `19.95`

**Explanation:** Rounds to 2 decimal places. The 4 in the third decimal position rounds down.

---

### Example 2: Round to Whole Number
**Data:**
| A | B |
|---|---|
| Amount | Rounded |
| 156.78 | =ROUND(A2, 0) |

**Formula:** `=ROUND(156.78, 0)`

**Result:** `157`

**Explanation:** num_digits of 0 rounds to the nearest whole number. Since .78 is above .5, it rounds up.

---

### Example 3: Round to Tens, Hundreds, Thousands
**Data:**
| A | B | C | D |
|---|---|---|---|
| Original | To Tens | To Hundreds | To Thousands |
| 12,847 | 12,850 | 12,800 | 13,000 |

**Formulas:**
- To tens: `=ROUND(12847, -1)` → `12850`
- To hundreds: `=ROUND(12847, -2)` → `12800`
- To thousands: `=ROUND(12847, -3)` → `13000`

**Explanation:** Negative num_digits rounds to the left of the decimal point.

---

### Example 4: Rounding Invoice Calculations
**Data:**
| A | B | C | D |
|---|---|---|---|
| Quantity | Price | Subtotal | Rounded Total |
| 7 | 12.99 | =A2*B2 | =ROUND(C2, 2) |

**Formula:** `=ROUND(7*12.99, 2)`

**Result:** `90.93`

**Explanation:** 7 × 12.99 = 90.93 exactly, but for complex calculations, ROUND ensures clean currency values.

---

### Example 5: Standard Rounding Rules Demonstrated
**Data showing rounding behavior:**
| A | B |
|---|---|
| Number | ROUND(A, 0) |
| 2.4 | 2 |
| 2.5 | 3 |
| 2.6 | 3 |
| -2.4 | -2 |
| -2.5 | -3 |
| -2.6 | -3 |

**Explanation:** 0.5 rounds away from zero (up for positive, down for negative numbers).

## Common Errors

| Error | Cause | Fix |
|-------|-------|-----|
| #VALUE! | Non-numeric value passed | Ensure number argument is numeric |
| Unexpected precision | Very large numbers | Excel has 15-digit precision limit |
| Still shows decimals | Cell formatting shows more decimals | Reduce decimal places in cell format, or result is exact |

## Pro Tips

- **ROUND vs Format:** ROUND actually changes the value; formatting only changes display. For calculations, ROUND is more reliable
- **Banker's rounding alternative:** Excel uses standard rounding. For "round half to even" (banker's rounding), use a custom formula
- **Preserve original:** Keep original values and round in a separate column for calculations
- **Chained rounding:** Don't nest ROUND functions unnecessarily - round once at the end for accuracy
- **Currency tip:** Always ROUND currency calculations to 2 decimals to avoid penny discrepancies

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [ROUNDUP](./ROUNDUP.md) | When you always want to round up (away from zero) |
| [ROUNDDOWN](./ROUNDDOWN.md) | When you always want to round down (toward zero) |
| [MROUND](./MROUND.md) | When you want to round to a specific multiple (like 0.05 or 5) |
| [CEILING](./CEILING.md) | When you want to round up to a multiple |
| [FLOOR](./FLOOR.md) | When you want to round down to a multiple |
| [INT](./INT.md) | When you want to round down to the nearest integer (floor function) |
| [TRUNC](./TRUNC.md) | When you want to truncate (remove decimals without rounding) |

## Version Notes
- **Available in:** Excel 2003, 2007, 2010, 2013, 2016, 2019, 2021, 365
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Math & Trig Functions](./README.md) | [🏠 Back to Home](../../README.md)
