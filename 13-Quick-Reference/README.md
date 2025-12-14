# 🚀 Quick Reference Guides

> **Fast lookup guides for instant help**

## 📋 Available Quick References

1. [Top 50 Most Used Functions](#top-50-most-used-functions)
2. [Common Error Messages](#common-error-messages)
3. [Function Categories Quick Finder](#function-categories-quick-finder)
4. [Excel Version Comparison](#excel-version-comparison)
5. [Best Practices Checklist](#best-practices-checklist)
6. [Performance Optimization Tips](#performance-optimization-tips)

---

## Top 50 Most Used Functions

### Lookup & Reference (Most Critical)
1. **VLOOKUP** - `=VLOOKUP(lookup_value, table, col_index, FALSE)` - Vertical lookup
2. **XLOOKUP** - `=XLOOKUP(value, lookup_array, return_array)` - Modern lookup (365)
3. **INDEX** - `=INDEX(array, row, col)` - Get value at position
4. **MATCH** - `=MATCH(value, array, 0)` - Find position
5. **FILTER** - `=FILTER(array, criteria)` - Filter data (365)

### Math & Aggregation
6. **SUM** - `=SUM(A1:A10)` - Add numbers
7. **SUMIF** - `=SUMIF(range, criteria, sum_range)` - Conditional sum
8. **SUMIFS** - `=SUMIFS(sum_range, criteria_range1, criteria1, ...)` - Multi-criteria sum
9. **AVERAGE** - `=AVERAGE(A1:A10)` - Calculate mean
10. **AVERAGEIF** - `=AVERAGEIF(range, criteria, avg_range)` - Conditional average
11. **COUNT** - `=COUNT(A1:A10)` - Count numbers
12. **COUNTA** - `=COUNTA(A1:A10)` - Count non-empty cells
13. **COUNTIF** - `=COUNTIF(range, criteria)` - Conditional count
14. **COUNTIFS** - `=COUNTIFS(range1, criteria1, range2, criteria2)` - Multi-criteria count
15. **SUMPRODUCT** - `=SUMPRODUCT(array1, array2)` - Multiply and sum

### Logical Functions
16. **IF** - `=IF(test, value_if_true, value_if_false)` - Conditional logic
17. **IFS** - `=IFS(test1, value1, test2, value2, ...)` - Multiple conditions (365)
18. **AND** - `=AND(logical1, logical2, ...)` - All must be true
19. **OR** - `=OR(logical1, logical2, ...)` - Any can be true
20. **IFERROR** - `=IFERROR(value, value_if_error)` - Handle errors
21. **IFNA** - `=IFNA(value, value_if_na)` - Handle #N/A

### Text Functions
22. **CONCATENATE/CONCAT** - `=CONCAT(text1, text2)` - Join text
23. **TEXTJOIN** - `=TEXTJOIN(delimiter, ignore_empty, text1, ...)` - Join with separator
24. **LEFT** - `=LEFT(text, num_chars)` - Extract from left
25. **RIGHT** - `=RIGHT(text, num_chars)` - Extract from right
26. **MID** - `=MID(text, start, num_chars)` - Extract from middle
27. **LEN** - `=LEN(text)` - Text length
28. **TRIM** - `=TRIM(text)` - Remove extra spaces
29. **UPPER/LOWER/PROPER** - `=UPPER(text)` - Change case
30. **SUBSTITUTE** - `=SUBSTITUTE(text, old, new)` - Replace text
31. **TEXT** - `=TEXT(value, format)` - Format as text
32. **FIND/SEARCH** - `=FIND(find_text, within_text)` - Find position

### Date & Time
33. **TODAY** - `=TODAY()` - Current date
34. **NOW** - `=NOW()` - Current date and time
35. **DATE** - `=DATE(year, month, day)` - Create date
36. **YEAR/MONTH/DAY** - `=YEAR(date)` - Extract date part
37. **EOMONTH** - `=EOMONTH(date, months)` - End of month
38. **DATEDIF** - `=DATEDIF(start, end, "Y")` - Date difference
39. **NETWORKDAYS** - `=NETWORKDAYS(start, end)` - Business days
40. **WORKDAY** - `=WORKDAY(start, days)` - Add business days

### Statistical
41. **MAX/MIN** - `=MAX(A1:A10)` - Maximum/minimum value
42. **LARGE/SMALL** - `=LARGE(array, k)` - Kth largest/smallest
43. **MEDIAN** - `=MEDIAN(A1:A10)` - Middle value
44. **MODE** - `=MODE.SNGL(A1:A10)` - Most frequent value
45. **STDEV** - `=STDEV.S(A1:A10)` - Standard deviation

### Array Functions (Excel 365)
46. **UNIQUE** - `=UNIQUE(array)` - Unique values
47. **SORT** - `=SORT(array, col, order)` - Sort data
48. **SORTBY** - `=SORTBY(array, by_array, order)` - Sort by another array
49. **SEQUENCE** - `=SEQUENCE(rows, cols, start, step)` - Generate sequence

### Special
50. **ROUND** - `=ROUND(number, decimals)` - Round number

---

## Common Error Messages

### #DIV/0!
**Meaning:** Division by zero  
**Cause:** Formula divides by zero or empty cell  
**Solution:**
```excel
=IFERROR(A1/B1, 0)                 // Return 0 instead of error
=IF(B1=0, "", A1/B1)               // Return blank if denominator is 0
```

---

### #N/A
**Meaning:** Value not available  
**Cause:** Lookup function can't find the value  
**Solution:**
```excel
=IFNA(VLOOKUP(A1, Table, 2, 0), "Not Found")
=XLOOKUP(A1, Range1, Range2, "Not Found")  // Built-in default
```

**Common Causes:**
- Lookup value doesn't exist in table
- Extra spaces in data (use TRIM)
- Different data types (text vs number)
- Approximate match when need exact

---

### #VALUE!
**Meaning:** Wrong type of argument or operand  
**Cause:** Wrong data type in formula  
**Solution:**
- Check that numbers aren't stored as text
- Verify all arguments are correct type
- Use VALUE() to convert text to number

**Common Causes:**
```excel
=A1+B1                             // Error if B1 contains text
=VALUE(B1)                         // Convert text to number first
```

---

### #REF!
**Meaning:** Invalid cell reference  
**Cause:** Reference to deleted cell or invalid range  
**Solution:**
- Fix formula to reference valid cells
- Undo deletion if recent
- Use named ranges (more resilient)

**Common Causes:**
- Deleted row/column referenced in formula
- VLOOKUP column index > table columns
- Copy formula outside valid range

---

### #NAME?
**Meaning:** Excel doesn't recognize text in formula  
**Cause:** Misspelled function, missing quotes, or undefined name  
**Solution:**
- Check function spelling
- Verify function exists in your Excel version (IFS, XLOOKUP require 365)
- Add quotes around text: `"text"` not `text`
- Define named range if using one

---

### #NUM!
**Meaning:** Invalid numeric value  
**Cause:** Invalid number in function  
**Solution:**
- Check for negative numbers where not allowed (e.g., SQRT)
- Verify result is within Excel limits
- Check iteration formulas

**Examples:**
```excel
=SQRT(-1)                          // Error - can't take sqrt of negative
=SQRT(ABS(A1))                     // Fix - use absolute value
```

---

### #NULL!
**Meaning:** Invalid cell intersection  
**Cause:** Incorrect reference operator (space instead of comma/colon)  
**Solution:**
```excel
=SUM(A1 A10)                       // Error - missing colon
=SUM(A1:A10)                       // Correct
```

---

### ##### (Hash Marks)
**Meaning:** Column too narrow or invalid date/time  
**Cause:** Number/date too wide for column  
**Solution:**
- Double-click column border to auto-fit
- Or drag column border wider

**For dates:**
- Check if date is negative (before 1/1/1900)
- Verify date format is valid

---

### Circular Reference Warning
**Meaning:** Formula refers to its own cell  
**Cause:** Formula creates circular reference  
**Solution:**
- Check formula doesn't reference itself
- Use iterative calculations if intentional (File → Options → Formulas)
- Redesign formula logic

**Example:**
```excel
// In A1:
=A1+B1                             // Error - references itself
=B1+C1                             // Correct
```

---

## Function Categories Quick Finder

### When to Use Each Category

**Math & Trig** - Use for:
- Calculations and arithmetic
- Rounding numbers
- Trigonometry
- Examples: SUM, ROUND, PRODUCT

**Statistical** - Use for:
- Data analysis
- Averages and distributions
- Forecasting
- Examples: AVERAGE, MEDIAN, STDEV

**Logical** - Use for:
- Decision making
- Conditional values
- Boolean logic
- Examples: IF, AND, OR

**Text** - Use for:
- String manipulation
- Data cleaning
- Parsing text
- Examples: LEFT, CONCAT, TRIM

**Date & Time** - Use for:
- Date calculations
- Time tracking
- Age calculations
- Examples: TODAY, DATEDIF, EOMONTH

**Lookup & Reference** - Use for:
- Finding data
- Table lookups
- Dynamic references
- Examples: VLOOKUP, INDEX, MATCH

**Financial** - Use for:
- Loans and investments
- NPV and IRR
- Depreciation
- Examples: PMT, NPV, FV

**Information** - Use for:
- Data type checking
- Error checking
- Cell information
- Examples: ISBLANK, ISERROR, CELL

**Database** - Use for:
- Database-style operations
- Criteria-based calculations
- Examples: DSUM, DAVERAGE

---

## Excel Version Comparison

### Feature Availability

| Feature | Excel 365 | Excel 2021 | Excel 2019 | Excel 2016 | Excel 2013 |
|---------|-----------|------------|------------|------------|------------|
| XLOOKUP | ✅ | ✅ | ❌ | ❌ | ❌ |
| FILTER | ✅ | ✅ | ❌ | ❌ | ❌ |
| SORT | ✅ | ✅ | ❌ | ❌ | ❌ |
| UNIQUE | ✅ | ✅ | ❌ | ❌ | ❌ |
| IFS | ✅ | ✅ | ✅ | ❌ | ❌ |
| SWITCH | ✅ | ✅ | ✅ | ❌ | ❌ |
| TEXTJOIN | ✅ | ✅ | ✅ | ❌ | ❌ |
| CONCAT | ✅ | ✅ | ✅ | ❌ | ❌ |
| MAXIFS/MINIFS | ✅ | ✅ | ✅ | ❌ | ❌ |
| Power Query | ✅ | ✅ | ✅ | ✅ | ✅ |
| Power Pivot | ✅ | ✅ | ✅ | ✅ (Pro+) | ✅ (Pro+) |
| Dynamic Arrays | ✅ | ✅ | ❌ | ❌ | ❌ |

### Alternatives for Older Versions

**Instead of XLOOKUP:**
```excel
// Use INDEX/MATCH
=INDEX(ReturnRange, MATCH(LookupValue, LookupRange, 0))
```

**Instead of FILTER:**
```excel
// Use Advanced Filter or PivotTable
```

**Instead of SORT:**
```excel
// Use Sort feature (Data → Sort)
```

**Instead of UNIQUE:**
```excel
// Use Remove Duplicates feature
```

**Instead of IFS:**
```excel
// Use nested IF
=IF(test1, value1, IF(test2, value2, IF(test3, value3, default)))
```

---

## Best Practices Checklist

### Formula Design
- ✅ Use absolute references ($) when copying formulas
- ✅ Use named ranges for clarity and maintenance
- ✅ Keep formulas simple - use helper columns if needed
- ✅ Add error handling (IFERROR, IFNA)
- ✅ Use exact match (FALSE/0) in lookups unless you need approximate
- ✅ Document complex formulas with comments

### Data Organization
- ✅ Use Excel Tables (Ctrl+T) for structured data
- ✅ Keep one table per worksheet
- ✅ Use headers in first row
- ✅ No blank rows or columns in data
- ✅ Consistent data types in each column
- ✅ No merged cells in data ranges

### Performance
- ✅ Avoid entire column references if possible (A:A)
- ✅ Limit use of volatile functions (NOW, TODAY, OFFSET, INDIRECT)
- ✅ Use SUMIFS instead of SUMPRODUCT for simple criteria
- ✅ Turn off automatic calculation for large workbooks
- ✅ Use manual PivotTable refresh
- ✅ Limit conditional formatting rules

### Maintainability
- ✅ Use consistent naming conventions
- ✅ Color code for purpose (blue=input, black=formula, green=output)
- ✅ Protect worksheets with formulas
- ✅ Document assumptions and sources
- ✅ Version control for important workbooks
- ✅ Regular backups

### Accuracy
- ✅ Test formulas with edge cases (zero, negative, blank)
- ✅ Verify lookup tables are sorted correctly (if using approximate match)
- ✅ Check for hidden rows/columns affecting results
- ✅ Use data validation to prevent bad input
- ✅ Cross-check totals
- ✅ Audit formulas regularly (Trace Precedents/Dependents)

---

## Performance Optimization Tips

### Slow Calculation? Try These:

**1. Reduce Volatile Functions**
- Replace OFFSET with INDEX
- Replace INDIRECT with direct references
- Calculate NOW() once and reference
- Use static ranges instead of dynamic

**2. Optimize Formulas**
```excel
// Slow
=SUMPRODUCT((A:A="Apple")*(B:B>100)*C:C)

// Faster
=SUMIFS(C:C, A:A, "Apple", B:B, ">100")
```

**3. Use Tables Instead of Ranges**
```excel
// Slow
=VLOOKUP(A2, Sheet2!$A:$Z, 5, 0)

// Faster (with Table)
=VLOOKUP(A2, TableName, 5, 0)
```

**4. Limit Array Formulas**
- Use helper columns instead
- In Excel 365, dynamic arrays are optimized

**5. Turn Off Automatic Calculation**
- Formulas tab → Calculation Options → Manual
- Calculate with F9 when needed

**6. Reduce Conditional Formatting**
- Limit rules to necessary ranges only
- Delete unused rules
- Avoid formulas in conditional formatting if possible

**7. Clean Up Workbook**
- Delete unused worksheets
- Clear unused formatting (select all, Clear Formats)
- Remove external links if not needed
- Save as new file to reduce size

**8. Use Efficient Lookups**
```excel
// Slow - VLOOKUP in large data
=VLOOKUP(A2, LargeTable, 10, 0)

// Faster - INDEX/MATCH
=INDEX(ReturnColumn, MATCH(A2, LookupColumn, 0))

// Fastest - XLOOKUP (365)
=XLOOKUP(A2, LookupColumn, ReturnColumn)
```

---

## Quick Troubleshooting Guide

### Formula Not Calculating?
1. Check if calculation is set to Manual (Formulas → Calculation Options)
2. Press F9 to force calculation
3. Check if cell formatted as Text
4. Look for circular reference warning

### VLOOKUP Returning Wrong Values?
1. Verify range_lookup is FALSE (0) for exact match
2. Check if table is sorted (required for approximate match)
3. Ensure lookup column is first column in table
4. Check for extra spaces (use TRIM)

### Formula Dragging Not Working?
1. Check if absolute references ($) are correct
2. Verify AutoFill is enabled (File → Options → Advanced)
3. Ensure cells aren't merged

### Numbers Not Adding Up?
1. Check if numbers stored as text
2. Look for hidden rows/columns
3. Verify SUBTOTAL function number (9 vs 109)
4. Check for rounding issues

---

**[⬆ Back to Main README](../../README.md)**
