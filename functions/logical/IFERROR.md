# IFERROR

## What It Does (Plain English)
Catches any error in a formula and returns a value you specify instead - like having a safety net that catches you when something goes wrong.

## Syntax
```
=IFERROR(value, value_if_error)
```

## Arguments

| Argument | Required? | Description |
|----------|-----------|-------------|
| value | Yes | The formula or value to check for errors |
| value_if_error | Yes | What to return if an error occurs |

## Returns
The result of the formula if no error; the value_if_error if any error occurs.

## Examples

### Example 1: Safe Division
**Data:**
| A | B |
|---|---|
| Revenue | Units |
| 1000 | 50 |
| 800 | 0 |
| 1200 | 40 |

**Formula:** `=IFERROR(A2/B2, "N/A")`

**Results:**
| Revenue | Units | Price |
|---|---|---|
| 1000 | 50 | 20 |
| 800 | 0 | N/A |
| 1200 | 40 | 30 |

**Explanation:** Division by zero would give #DIV/0!, but IFERROR shows "N/A" instead.

---

### Example 2: Safe VLOOKUP
**Formula:** `=IFERROR(VLOOKUP(A2, PriceTable, 2, FALSE), "Not found")`

**Without IFERROR:** Would show #N/A if product not in table
**With IFERROR:** Shows "Not found"

**Explanation:** When VLOOKUP can't find a match, catch the error gracefully.

---

### Example 3: Return Zero Instead of Error
**Formula:** `=IFERROR(A2/B2, 0)`

**Result:** Returns 0 for any error, allowing SUM/AVERAGE to work properly.

**Explanation:** Use 0 when you want errors to be treated as zero in calculations.

---

### Example 4: Return Empty String (Blank)
**Formula:** `=IFERROR(VLOOKUP(A2, Data, 3, FALSE), "")`

**Result:** Cell appears blank instead of showing an error.

**Explanation:** Use `""` (empty quotes) when you want the cell to look empty.

---

### Example 5: Chain IFERROR with Default Calculation
**Formula:** `=IFERROR(A2/B2, IFERROR(A2/C2, 0))`

**Explanation:** Try A/B first, if error try A/C, if still error return 0.

---

### Example 6: Nested Formula Protection
**Complex formula with multiple potential errors:**

**Formula:** `=IFERROR(INDEX(Data, MATCH(A2, IDs, 0), MATCH(B2, Headers, 0)), "Check inputs")`

**Explanation:** Wrapping the entire formula protects against errors from either MATCH.

## Common Errors

| Issue | Cause | Fix |
|-------|-------|-----|
| Hides real problems | IFERROR catches ALL errors | Use IFNA for just #N/A, or troubleshoot before adding IFERROR |
| Wrong result hidden | Data issue not visible | Test formula without IFERROR first |
| Can't debug | Error message suppressed | Temporarily remove IFERROR to see actual error |

## Pro Tips

- **Don't hide data problems:** IFERROR can mask legitimate data issues. Use sparingly and intentionally
- **IFNA is more specific:** For VLOOKUP/MATCH, use IFNA to catch only #N/A, not calculation errors
- **Testing:** Build formulas without IFERROR first, add it after confirming the formula works
- **Alternatives in XLOOKUP:** XLOOKUP has built-in if_not_found - no IFERROR needed
- **With AGGREGATE:** AGGREGATE function can ignore errors: `=AGGREGATE(9,6,A:A)` sums ignoring errors

## Errors Caught by IFERROR

| Error | Typical Cause |
|-------|---------------|
| #DIV/0! | Division by zero |
| #N/A | Lookup found nothing |
| #VALUE! | Wrong argument type |
| #REF! | Invalid cell reference |
| #NAME? | Unknown function/range name |
| #NUM! | Invalid numeric value |
| #NULL! | Incorrect range operator |

## Related Functions

| Function | When to Use Instead |
|----------|-------------------|
| [IFNA](./IFNA.md) | Catch only #N/A errors (more specific) |
| [IF](./IF.md) | When you need to test a condition, not catch an error |
| [ISERROR](../information/ISERROR.md) | When you need TRUE/FALSE for "is this an error?" |
| [ERROR.TYPE](../information/ERROR.TYPE.md) | When you need to identify which error occurred |

## Version Notes
- **Available in:** Excel 2007, 2010, 2013, 2016, 2019, 2021, 365
- **NOT available in:** Excel 2003 (use `=IF(ISERROR(...), alt_value, formula)`)
- **Google Sheets:** ✅ Fully compatible
- **LibreOffice Calc:** ✅ Fully compatible

---
[📘 Back to Logical Functions](./README.md) | [🏠 Back to Home](../../README.md)
