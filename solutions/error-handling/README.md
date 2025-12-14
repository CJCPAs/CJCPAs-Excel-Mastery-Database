# Error Handling Solutions

> **Catch, prevent, and gracefully handle errors in your Excel formulas**

## Quick Solutions

| Error | Meaning | Solution |
|-------|---------|----------|
| #N/A | Value not found | [IFNA](#handle-na-errors) or [IFERROR](#handle-any-error) |
| #DIV/0! | Division by zero | [Check before dividing](#prevent-division-by-zero) |
| #VALUE! | Wrong argument type | [Validate data types](#handle-value-errors) |
| #REF! | Invalid reference | [Fix cell references](#prevent-ref-errors) |
| #NAME? | Unknown function/name | [Check spelling](#fix-name-errors) |
| #NUM! | Invalid number | [Validate numeric inputs](#handle-num-errors) |
| Any error | Catch all | [IFERROR](#handle-any-error) |

---

## Handle #N/A Errors

### The Challenge
VLOOKUP, XLOOKUP, and MATCH return #N/A when they can't find the value.

### Quick Answer
```excel
=IFNA(formula, value_if_na)           // Catches only #N/A
=IFERROR(formula, value_if_error)     // Catches all errors
```

### Full Example

**Problem Formula:**
```excel
=VLOOKUP("Unknown", A2:B10, 2, FALSE)
```
**Result:** `#N/A`

**Solution 1 - IFNA:**
```excel
=IFNA(VLOOKUP("Unknown", A2:B10, 2, FALSE), "Not found")
```
**Result:** `Not found`

**Solution 2 - XLOOKUP (has built-in handling):**
```excel
=XLOOKUP("Unknown", A2:A10, B2:B10, "Not found")
```

### When to Use IFNA vs IFERROR

| Use IFNA when... | Use IFERROR when... |
|------------------|---------------------|
| Only catching lookup failures | Catching any calculation error |
| Want other errors visible | Want to hide all errors |
| Troubleshooting formulas | Final production formulas |

---

## Prevent Division by Zero

### The Challenge
Dividing by zero or an empty cell causes #DIV/0! error.

### Quick Answer
```excel
=IF(B2=0, 0, A2/B2)                   // Return 0 if divisor is 0
=IF(B2=0, "", A2/B2)                  // Return blank if divisor is 0
=IFERROR(A2/B2, 0)                    // Catch the error after it happens
```

### Full Example

**Data:**
| A | B |
|---|---|
| Revenue | Units |
| 1000 | 50 |
| 500 | 0 |
| 750 | 25 |

**Problem Formula:** `=A2/B2`

**Row 2 Result:** `20`
**Row 3 Result:** `#DIV/0!` ← Problem!

**Solution Formula:** `=IF(B2=0, "N/A", A2/B2)`

**Row 3 Result:** `N/A`

### Best Practice
```excel
=IF(B2=0, "No units sold", ROUND(A2/B2, 2))
```
- Checks for zero first
- Provides meaningful message
- Rounds result for readability

---

## Handle Any Error

### The Challenge
You want to catch ANY error and replace it with something useful.

### Quick Answer
```excel
=IFERROR(formula, value_if_error)
```

### Full Example

**Complex formula that might fail:**
```excel
=INDEX(Data, MATCH(A2, IDs, 0), MATCH(B2, Headers, 0))
```

**Error-proof version:**
```excel
=IFERROR(INDEX(Data, MATCH(A2, IDs, 0), MATCH(B2, Headers, 0)), "Check inputs")
```

### Error Types Caught by IFERROR

| Error | Typical Cause |
|-------|---------------|
| #DIV/0! | Division by zero |
| #N/A | Lookup not found |
| #VALUE! | Wrong data type |
| #REF! | Invalid reference |
| #NAME? | Unknown function name |
| #NUM! | Invalid number |
| #NULL! | Incorrect range |

### Warning
IFERROR hides ALL errors - this can mask real problems. Use sparingly!

---

## Handle #VALUE! Errors

### The Challenge
#VALUE! occurs when a formula expects a number but gets text (or vice versa).

### Quick Answer
```excel
=IFERROR(A2*B2, "Invalid data")
=IF(ISNUMBER(A2), A2*B2, "Not a number")
```

### Common Causes

| Cause | Example | Fix |
|-------|---------|-----|
| Text in calculation | `="5"*10` | `=VALUE("5")*10` |
| Date as text | `=DATEVALUE("date")+1` | Check date format |
| Space in "empty" cell | `=A1*2` (A1 has space) | `=TRIM(A1)*2` |

### Validate Before Calculating
```excel
=IF(AND(ISNUMBER(A2), ISNUMBER(B2)), A2*B2, "Invalid input")
```

---

## Prevent #REF! Errors

### The Challenge
#REF! occurs when a formula references a cell that was deleted.

### Prevention Tips

1. **Use Tables:** Table references automatically adjust when rows/columns are deleted
2. **Use Named Ranges:** Named ranges are more resilient than cell references
3. **Use INDIRECT carefully:** Can cause #REF! if the text reference is invalid

### Quick Fix
```excel
=IFERROR(formula, "Reference error - check data")
```

### Example - Safe Reference
Instead of:
```excel
=SUM(Sheet2!A1:A100)
```
Use:
```excel
=IFERROR(SUM(Sheet2!A1:A100), "Sheet2 data not available")
```

---

## Fix #NAME? Errors

### The Challenge
#NAME? means Excel doesn't recognize something in your formula.

### Common Causes & Fixes

| Cause | Example | Fix |
|-------|---------|-----|
| Misspelled function | `=VLOKUP(...)` | `=VLOOKUP(...)` |
| Missing quotes | `=VLOOKUP(North,...)` | `=VLOOKUP("North",...)` |
| Undefined name | `=SUM(SalesData)` | Define the named range |
| Missing add-in | `=XIRR(...)` | Enable Analysis ToolPak |

### Text Must Be in Quotes
```excel
Wrong:  =IF(A1=Yes, "OK", "No")
Right:  =IF(A1="Yes", "OK", "No")
```

---

## Handle #NUM! Errors

### The Challenge
#NUM! occurs when a calculation produces an invalid numeric result.

### Common Causes

| Cause | Example | Fix |
|-------|---------|-----|
| Negative square root | `=SQRT(-4)` | Check input is positive |
| IRR can't converge | `=IRR(bad_data)` | Check cash flows |
| Number too large | `=FACT(200)` | Reduce input |

### Example - Safe Square Root
```excel
=IF(A2>=0, SQRT(A2), "Cannot compute")
```

---

## Create Error-Proof Formulas

### The Complete Error-Handling Pattern

**Basic Pattern:**
```excel
=IFERROR(formula, fallback_value)
```

**With Specific Message:**
```excel
=IFERROR(VLOOKUP(A2, Table, 2, FALSE), "ID not found in database")
```

**With Alternative Calculation:**
```excel
=IFERROR(A2/B2, A2/AVERAGE(B:B))
```

**Nested Error Handling:**
```excel
=IFERROR(VLOOKUP(A2, Table1, 2, FALSE),
         IFERROR(VLOOKUP(A2, Table2, 2, FALSE), "Not in either table"))
```

### Validate Inputs First (Better Practice)
```excel
=IF(ISBLANK(A2), "Enter ID",
    IF(ISNUMBER(A2), VLOOKUP(A2, Data, 2, FALSE), "ID must be a number"))
```

---

## Error Checking Functions

### Test for Specific Errors
```excel
=ISERROR(A2)      // TRUE if any error
=ISNA(A2)         // TRUE if #N/A
=ISERR(A2)        // TRUE if error but NOT #N/A
=ISNUMBER(A2)     // TRUE if number (not error, not text)
=ISTEXT(A2)       // TRUE if text
=ISBLANK(A2)      // TRUE if empty
```

### Identify Error Type
```excel
=ERROR.TYPE(A2)
```
Returns: 1=#NULL!, 2=#DIV/0!, 3=#VALUE!, 4=#REF!, 5=#NAME?, 6=#NUM!, 7=#N/A

### Create Custom Error Messages
```excel
=IF(ISERROR(A2/B2),
    SWITCH(ERROR.TYPE(A2/B2),
           2, "Division by zero",
           3, "Invalid value",
           7, "Data not found",
           "Unknown error"),
    A2/B2)
```

---

## Best Practices

1. **Don't over-use IFERROR** - It can hide real problems
2. **Use IFNA for lookups** - More specific than IFERROR
3. **Validate inputs first** - Better than catching errors after
4. **Test without error handling** - Debug formulas before wrapping
5. **Provide meaningful messages** - "Error" is less helpful than "Product ID not found"

---

## Related Solutions

- [Lookups](../lookups/README.md) - Lookup formulas that often need error handling
- [Conditional Calculations](../conditional-calculations/README.md) - Using IF to prevent errors
- [Troubleshooting](../../troubleshooting/README.md) - Debug formula problems

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
