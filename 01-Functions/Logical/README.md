# 🔀 Logical Functions

> **Master conditional logic and decision-making with Excel's 15+ logical functions**

## 📋 Table of Contents

- [Core Logical Functions](#core-logical-functions)
- [Advanced Logical Functions](#advanced-logical-functions)
- [Error Handling](#error-handling)
- [Practical Examples](#practical-examples)

---

## Core Logical Functions

### IF
**Returns one value if condition is TRUE, another if FALSE**

**Syntax:** `=IF(logical_test, value_if_true, [value_if_false])`

**Parameters:**
- `logical_test`: Condition to evaluate (returns TRUE or FALSE)
- `value_if_true`: Value to return when TRUE
- `value_if_false`: (Optional) Value to return when FALSE (default: FALSE)

**Examples:**
```excel
=IF(A1>100, "High", "Low")              → "High" if A1>100, else "Low"
=IF(B1="", "Empty", "Has Value")        → Check if cell is empty
=IF(C1>=60, "Pass", "Fail")             → Grade pass/fail
=IF(D1>0, D1, 0)                        → Return value or zero
=IF(AND(A1>50, B1<100), "Yes", "No")    → Nested with AND
```

**Real-World Uses:**
- Grade calculations: Pass/Fail
- Sales commissions: Different rates based on performance
- Status indicators: Active/Inactive
- Conditional pricing

**Nested IF (Multiple Conditions):**
```excel
=IF(A1>=90, "A", IF(A1>=80, "B", IF(A1>=70, "C", IF(A1>=60, "D", "F"))))
```

**Tips:**
- Can nest up to 64 IF functions (but use IFS for readability)
- Empty cells evaluate to 0 in numeric comparisons
- Text comparisons are not case-sensitive by default
- Use "" for blank result

---

### AND
**Returns TRUE if all conditions are TRUE**

**Syntax:** `=AND(logical1, [logical2], ...)`

**Examples:**
```excel
=AND(A1>50, B1<100)                     → TRUE if both conditions met
=AND(A1="Apple", B1="Red")              → TRUE if both exact matches
=AND(A1>0, A1<100, A1<>50)              → Multiple conditions
=AND(A1:A10>0)                          → TRUE if all values >0
```

**Real-World Uses:**
- Validate multiple criteria
- Check if all conditions met for approval
- Verify data completeness

**With IF:**
```excel
=IF(AND(A1>=18, B1="Yes"), "Eligible", "Not Eligible")
```

**Tips:**
- Returns FALSE if any argument is FALSE
- Ignores text and empty cells
- Maximum 255 arguments

---

### OR
**Returns TRUE if any condition is TRUE**

**Syntax:** `=OR(logical1, [logical2], ...)`

**Examples:**
```excel
=OR(A1>100, B1>100)                     → TRUE if either >100
=OR(A1="Red", A1="Blue", A1="Green")    → TRUE if any color matches
=OR(ISBLANK(A1), ISBLANK(B1))           → TRUE if either cell blank
```

**Real-World Uses:**
- Flag items meeting any criteria
- Multiple valid options
- Exception handling

**With IF:**
```excel
=IF(OR(A1="Urgent", A1="Critical"), "Priority", "Normal")
```

---

### NOT
**Reverses the logical value**

**Syntax:** `=NOT(logical)`

**Examples:**
```excel
=NOT(A1>100)                            → TRUE if A1 is NOT >100
=NOT(ISBLANK(A1))                       → TRUE if cell is NOT blank
=NOT(AND(A1>50, B1<100))                → Reverse AND result
```

**Real-World Uses:**
- Invert conditions
- Check for "not equal"
- Reverse TRUE/FALSE flags

**Tips:**
- `NOT(TRUE)` = FALSE
- `NOT(FALSE)` = TRUE
- Useful with ISBLANK, ISERROR functions

---

### XOR
**Exclusive OR - TRUE if odd number of arguments are TRUE**

**Syntax:** `=XOR(logical1, [logical2], ...)`

**Examples:**
```excel
=XOR(A1>50, B1>50)                      → TRUE if only one >50
=XOR(TRUE, FALSE)                       → TRUE
=XOR(TRUE, TRUE)                        → FALSE
=XOR(TRUE, FALSE, FALSE)                → TRUE (odd number)
```

**Real-World Uses:**
- Either/or scenarios
- Validation: exactly one option selected
- Toggle switches

---

## Advanced Logical Functions

### IFS
**Modern replacement for nested IFs (Excel 2019+)**

**Syntax:** `=IFS(logical_test1, value1, [logical_test2, value2], ...)`

**Examples:**
```excel
=IFS(A1>=90, "A", A1>=80, "B", A1>=70, "C", A1>=60, "D", TRUE, "F")
=IFS(B1="Small", 5, B1="Medium", 10, B1="Large", 20)
```

**Real-World Uses:**
- Grade calculations
- Pricing tiers
- Status categorization

**Advantages over nested IF:**
- Much more readable
- Easier to maintain
- Less prone to errors

**Tips:**
- Tests evaluated left to right
- Returns value of first TRUE test
- Use `TRUE` as last test for "else" condition
- Generates #N/A if no conditions TRUE

---

### SWITCH
**Returns value matching an expression (Excel 2019+)**

**Syntax:** `=SWITCH(expression, value1, result1, [value2, result2], ..., [default])`

**Examples:**
```excel
=SWITCH(A1, 1, "Jan", 2, "Feb", 3, "Mar", "Unknown")
=SWITCH(B1, "S", "Small", "M", "Medium", "L", "Large", "Invalid")
=SWITCH(C1, "Red", "Stop", "Yellow", "Caution", "Green", "Go")
```

**Real-World Uses:**
- Convert codes to descriptions
- Map values to categories
- Lookup tables in formulas

**SWITCH vs IFS:**
- Use SWITCH for exact matches
- Use IFS for ranges/conditions

**Tips:**
- More efficient than nested IFs for exact matches
- Default value is optional
- Maximum 126 pairs

---

### IFERROR
**Returns custom value if formula results in error**

**Syntax:** `=IFERROR(value, value_if_error)`

**Examples:**
```excel
=IFERROR(A1/B1, 0)                      → Return 0 if division error
=IFERROR(VLOOKUP(A1,Table,2,0), "Not Found")  → Custom message
=IFERROR(INDEX(MATCH(...)), "")         → Return blank on error
```

**Real-World Uses:**
- Handle #DIV/0! errors
- Manage #N/A from lookups
- Clean error displays

**Errors it catches:**
- `#DIV/0!` - Division by zero
- `#N/A` - Value not available
- `#VALUE!` - Wrong type
- `#REF!` - Invalid reference
- `#NUM!` - Invalid number
- `#NAME?` - Unrecognized name
- `#NULL!` - Invalid intersection

**Tips:**
- Use sparingly - can hide real errors
- Consider IFNA for lookups specifically
- Don't use to bypass data issues

---

### IFNA
**Returns custom value only for #N/A errors**

**Syntax:** `=IFNA(value, value_if_na)`

**Examples:**
```excel
=IFNA(VLOOKUP(A1,Table,2,0), "Not Found")
=IFNA(MATCH(A1,Range,0), "Missing")
=IFNA(XLOOKUP(A1,Range1,Range2), 0)
```

**IFNA vs IFERROR:**
- IFNA: Only catches #N/A (better for lookups)
- IFERROR: Catches all errors

**Recommended:**
```excel
=IFNA(XLOOKUP(...), "Not Found")        ✓ Recommended
=IFERROR(XLOOKUP(...), "Not Found")     ✗ Too broad
```

---

### TRUE / FALSE
**Returns the logical value TRUE or FALSE**

**Syntax:** 
```excel
=TRUE()
=FALSE()
```

**Examples:**
```excel
=TRUE()                                 → TRUE
=FALSE()                                → FALSE
=IF(A1>50, TRUE(), FALSE())             → Returns logical value
```

**Tips:**
- Rarely needed - can type TRUE or FALSE directly
- Sometimes required for compatibility
- Useful in array formulas

---

## Practical Examples

### Example 1: Grade Calculator
```excel
=IFS(A1>=90, "A", A1>=80, "B", A1>=70, "C", A1>=60, "D", TRUE, "F")
```

### Example 2: Discount Calculation
```excel
=IF(B1>=1000, B1*0.9, IF(B1>=500, B1*0.95, B1))  // 10% off if ≥1000, 5% off if ≥500
```

### Example 3: Status Indicator
```excel
=IFS(C1<TODAY()-30, "Overdue", C1<TODAY(), "Due Soon", TRUE, "On Time")
```

### Example 4: Validate Complete Data
```excel
=IF(AND(NOT(ISBLANK(A1)), NOT(ISBLANK(B1)), NOT(ISBLANK(C1))), "Complete", "Incomplete")
```

### Example 5: Convert Size Codes
```excel
=SWITCH(A1, "S", "Small", "M", "Medium", "L", "Large", "XL", "Extra Large", "Unknown")
```

### Example 6: Price Lookup with Error Handling
```excel
=IFERROR(VLOOKUP(A1, PriceTable, 2, FALSE), "Price Not Found")
```

### Example 7: Multi-Criteria Approval
```excel
=IF(AND(B1>=18, C1="Yes", D1>50000, E1="Good"), "Approved", "Denied")
```

### Example 8: Flag Outliers
```excel
=IF(OR(A1<AVERAGE($A$1:$A$100)-2*STDEV($A$1:$A$100), 
       A1>AVERAGE($A$1:$A$100)+2*STDEV($A$1:$A$100)), 
   "Outlier", "Normal")
```

### Example 9: Nested Conditions
```excel
=IF(A1="", "No Data", 
    IF(A1<0, "Negative", 
        IF(A1=0, "Zero", 
            IF(A1<100, "Small", "Large"))))
```

### Example 10: Year Quarter
```excel
=SWITCH(MONTH(A1), 1,2,3, "Q1", 4,5,6, "Q2", 7,8,9, "Q3", 10,11,12, "Q4")
```

---

## Combining Logical Functions

### AND with IF
```excel
=IF(AND(A1>=18, B1="Active", C1>0), "Eligible", "Not Eligible")
```

### OR with IF
```excel
=IF(OR(A1="Urgent", A1="Critical", B1<TODAY()), "Priority", "Standard")
```

### Nested AND/OR
```excel
=IF(OR(AND(A1="Premium", B1>1000), AND(A1="VIP", B1>500)), "Discount", "Regular")
```

### NOT with AND
```excel
=IF(NOT(AND(ISBLANK(A1), ISBLANK(B1))), "Has Data", "Empty")
```

### Complex Logic
```excel
=IF(AND(OR(A1="Red", A1="Blue"), B1>100, NOT(ISBLANK(C1))), "Match", "No Match")
```

---

## Truth Tables

### AND Truth Table
| A | B | AND(A,B) |
|---|---|----------|
| TRUE | TRUE | TRUE |
| TRUE | FALSE | FALSE |
| FALSE | TRUE | FALSE |
| FALSE | FALSE | FALSE |

### OR Truth Table
| A | B | OR(A,B) |
|---|---|---------|
| TRUE | TRUE | TRUE |
| TRUE | FALSE | TRUE |
| FALSE | TRUE | TRUE |
| FALSE | FALSE | FALSE |

### XOR Truth Table
| A | B | XOR(A,B) |
|---|---|----------|
| TRUE | TRUE | FALSE |
| TRUE | FALSE | TRUE |
| FALSE | TRUE | TRUE |
| FALSE | FALSE | FALSE |

### NOT Truth Table
| A | NOT(A) |
|---|--------|
| TRUE | FALSE |
| FALSE | TRUE |

---

## Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| IF | Conditional value | `=IF(A1>100,"High","Low")` |
| IFS | Multiple conditions | `=IFS(A1>90,"A",A1>80,"B",TRUE,"C")` |
| SWITCH | Match expression | `=SWITCH(A1,1,"Jan",2,"Feb",3,"Mar")` |
| AND | All TRUE | `=AND(A1>50,B1<100)` |
| OR | Any TRUE | `=OR(A1>100,B1>100)` |
| NOT | Reverse | `=NOT(ISBLANK(A1))` |
| XOR | Exclusive OR | `=XOR(A1>50,B1>50)` |
| IFERROR | Handle errors | `=IFERROR(A1/B1,0)` |
| IFNA | Handle #N/A | `=IFNA(VLOOKUP(...),"Not Found")` |

---

## Common Patterns

### Check if cell is blank
```excel
=IF(A1="", "Blank", "Not Blank")
=IF(ISBLANK(A1), "Blank", "Not Blank")
```

### Check if number is between values
```excel
=IF(AND(A1>=50, A1<=100), "In Range", "Out of Range")
```

### Convert Yes/No to 1/0
```excel
=IF(A1="Yes", 1, 0)
=IF(A1="Yes", TRUE, FALSE)
```

### Multiple OR conditions
```excel
=IF(OR(A1="Apple", A1="Orange", A1="Banana"), "Fruit", "Not Fruit")
```

### Check multiple cells not blank
```excel
=IF(AND(A1<>"", B1<>"", C1<>""), "Complete", "Incomplete")
```

---

## Best Practices

### Readability
- Use IFS instead of nested IFs when possible
- Break complex formulas into helper columns
- Add comments for complex logic
- Use named ranges for clarity

### Performance
- Put most likely conditions first
- Use SWITCH for exact matches (faster than IFs)
- Avoid unnecessary nested functions

### Error Prevention
- Always include value_if_false in IF
- Use IFERROR/IFNA for lookups
- Test edge cases (blank, zero, negative)
- Validate data types

### Maintenance
- Document complex logic
- Use consistent patterns across workbook
- Avoid deeply nested formulas (use helper columns)
- Test thoroughly before deploying

---

## Common Errors & Solutions

### #VALUE! Error
**Cause:** Wrong data type
**Solution:** Check that logical tests return TRUE/FALSE

### #NAME? Error
**Cause:** Function not recognized (old Excel version)
**Solution:** IFS, SWITCH require Excel 2019+

### Unexpected Results
**Cause:** Order of operations in nested functions
**Solution:** Use parentheses to control evaluation order

### TRUE/FALSE showing instead of values
**Cause:** Missing value_if_false parameter
**Solution:** `=IF(A1>100, "High", "Low")` not `=IF(A1>100, "High")`

---

## Advanced Tips

### Short-Circuit Evaluation
Excel evaluates AND left to right and stops at first FALSE:
```excel
=AND(A1>0, B1/A1>10)  // Safe - won't divide if A1≤0
```

### Boolean Arithmetic
```excel
=(A1>50)+(B1>50)      // Count how many TRUE (0, 1, or 2)
=(A1>50)*(B1>50)      // Returns 1 only if both TRUE
```

### Array Formulas
```excel
=IF(A1:A10>50, "High", "Low")  // Returns array in modern Excel
```

### Conditional Aggregation
```excel
=SUM(IF(A1:A10>50, B1:B10, 0))  // Sum B where A>50
```

---

**[⬆ Back to Main README](../../README.md)**
