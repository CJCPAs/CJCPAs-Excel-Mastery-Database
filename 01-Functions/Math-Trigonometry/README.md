# 🔢 Math & Trigonometry Functions

> **80+ functions for mathematical calculations, rounding operations, and trigonometric computations**

## 📋 Table of Contents

- [Basic Math Functions](#basic-math-functions)
- [Aggregation Functions](#aggregation-functions)
- [Rounding Functions](#rounding-functions)
- [Trigonometric Functions](#trigonometric-functions)
- [Advanced Math](#advanced-math)
- [Matrix Operations](#matrix-operations)

---

## Basic Math Functions

### SUM
**Adds all numbers in a range**

**Syntax:** `=SUM(number1, [number2], ...)`

**Examples:**
```excel
=SUM(A1:A10)              → Adds all values in A1 through A10
=SUM(A1,B1,C1)           → Adds three specific cells
=SUM(A1:A5,C1:C5)        → Adds two separate ranges
=SUM(100,200,300)        → Returns 600
```

**Real-World Use:**
- Total sales for the month
- Sum of expenses across categories
- Calculate grand totals

**Tips:**
- Ignores text and logical values
- Can handle up to 255 arguments
- Use SUM instead of + for large ranges

---

### SUMIF
**Adds cells that meet a specific condition**

**Syntax:** `=SUMIF(range, criteria, [sum_range])`

**Parameters:**
- `range`: Cells to evaluate
- `criteria`: Condition to match
- `sum_range`: (Optional) Cells to sum

**Examples:**
```excel
=SUMIF(A1:A10,">100")                  → Sum values greater than 100
=SUMIF(A1:A10,"Apple",B1:B10)          → Sum B values where A="Apple"
=SUMIF(C1:C10,">=50",D1:D10)           → Sum D where C is 50 or more
=SUMIF(A:A,"*North*",B:B)              → Sum B where A contains "North"
```

**Real-World Use:**
- Sum sales for a specific product
- Total expenses over a threshold
- Calculate regional totals

**Tips:**
- Use wildcards: * (any characters), ? (single character)
- Criteria can be number, text, or cell reference
- For multiple criteria, use SUMIFS

---

### SUMIFS
**Adds cells that meet multiple conditions**

**Syntax:** `=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Examples:**
```excel
=SUMIFS(D:D, A:A, "Apple", B:B, ">100")           → Sum where product="Apple" AND value>100
=SUMIFS(Sales, Region, "North", Month, "Jan")      → Sum sales for North region in January
=SUMIFS(C:C, A:A, ">=2024-01-01", A:A, "<2024-02-01")  → Sum for date range
```

**Real-World Use:**
- Sales by region and product
- Expenses by department and month
- Multi-criteria financial analysis

**Common Errors:**
- `#VALUE!` - Ranges different sizes
- Ensure all ranges are same dimension

---

### AVERAGE
**Calculates the arithmetic mean**

**Syntax:** `=AVERAGE(number1, [number2], ...)`

**Examples:**
```excel
=AVERAGE(A1:A10)          → Average of range
=AVERAGE(10,20,30)        → Returns 20
=AVERAGE(A:A)             → Average of entire column (ignores blanks)
```

**Real-World Use:**
- Average sales per day
- Mean test score
- Average response time

**Related Functions:**
- AVERAGEIF - Average with criteria
- AVERAGEIFS - Average with multiple criteria
- AVERAGEA - Includes text (as 0)

---

### COUNT
**Counts cells containing numbers**

**Syntax:** `=COUNT(value1, [value2], ...)`

**Examples:**
```excel
=COUNT(A1:A10)            → Count numeric cells
=COUNT(A:A)               → Count numbers in column
=COUNT(1,2,"text",TRUE)   → Returns 2 (only counts numbers)
```

**Related Functions:**
- COUNTA - Counts non-empty cells
- COUNTBLANK - Counts empty cells
- COUNTIF - Counts with criteria
- COUNTIFS - Counts with multiple criteria

---

### COUNTIF
**Counts cells meeting a condition**

**Syntax:** `=COUNTIF(range, criteria)`

**Examples:**
```excel
=COUNTIF(A1:A10,">50")              → Count values > 50
=COUNTIF(B:B,"Apple")               → Count cells = "Apple"
=COUNTIF(C:C,"<>0")                 → Count non-zero values
=COUNTIF(D:D,"*ing")                → Count ending with "ing"
```

**Real-World Use:**
- Count overdue invoices
- Number of A grades
- Items in stock

---

### COUNTIFS
**Counts cells meeting multiple conditions**

**Syntax:** `=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Examples:**
```excel
=COUNTIFS(A:A,"Apple",B:B,">100")              → Count where A="Apple" AND B>100
=COUNTIFS(Region,"North",Status,"Complete")     → Count completed North sales
```

---

## Aggregation Functions

### PRODUCT
**Multiplies all numbers**

**Syntax:** `=PRODUCT(number1, [number2], ...)`

**Examples:**
```excel
=PRODUCT(A1:A5)           → Multiply all values
=PRODUCT(2,3,4)           → Returns 24
=PRODUCT(A1:A3,10)        → Multiply range by 10
```

**Real-World Use:**
- Calculate compound growth
- Volume calculations (length × width × height)
- Probability calculations

---

### SUBTOTAL
**Returns a subtotal (ignores hidden rows)**

**Syntax:** `=SUBTOTAL(function_num, ref1, [ref2], ...)`

**Function Numbers:**
```
1 = AVERAGE      101 = AVERAGE (ignore hidden)
2 = COUNT        102 = COUNT
3 = COUNTA       103 = COUNTA
4 = MAX          104 = MAX
5 = MIN          105 = MIN
6 = PRODUCT      106 = PRODUCT
7 = STDEV        107 = STDEV
8 = STDEVP       108 = STDEVP
9 = SUM          109 = SUM
10 = VAR         110 = VAR
11 = VARP        111 = VARP
```

**Examples:**
```excel
=SUBTOTAL(9,A1:A10)       → Sum (ignores other SUBTOTAL)
=SUBTOTAL(109,A:A)        → Sum ignoring hidden rows
=SUBTOTAL(1,B2:B100)      → Average
```

**Real-World Use:**
- Totals in filtered data
- Nested subtotals
- Summary rows in tables

**Why Use It:**
- Works with AutoFilter
- Ignores hidden rows (100+ function codes)
- Prevents double-counting in nested subtotals

---

### AGGREGATE
**Like SUBTOTAL but more powerful**

**Syntax:** `=AGGREGATE(function_num, options, ref1, [ref2])`

**Function Numbers:** (Same as SUBTOTAL 1-11, plus more)

**Options:**
```
0 = Ignore nested SUBTOTAL/AGGREGATE
1 = Ignore hidden rows
2 = Ignore error values
3 = Ignore hidden rows and errors
4 = Ignore nothing
5 = Ignore hidden columns
6 = Ignore hidden rows and columns
7 = Ignore hidden rows, columns, and errors
```

**Examples:**
```excel
=AGGREGATE(9,6,A1:A100)              → Sum, ignore hidden rows
=AGGREGATE(14,6,A1:A100,3)           → 3rd largest, ignore hidden
=AGGREGATE(15,6,A1:A100,1)           → Smallest, ignore hidden
```

**Advantages over SUBTOTAL:**
- Can ignore errors
- Includes LARGE, SMALL, PERCENTILE functions
- More control with options

---

### SUMPRODUCT
**Multiplies corresponding ranges and sums the products**

**Syntax:** `=SUMPRODUCT(array1, [array2], ...)`

**Examples:**
```excel
=SUMPRODUCT(A1:A5,B1:B5)              → (A1*B1)+(A2*B2)+...+(A5*B5)
=SUMPRODUCT((A1:A10="Apple")*(B1:B10))  → Sum B where A="Apple"
=SUMPRODUCT(--(A1:A10>50),B1:B10)     → Sum B where A>50
```

**Real-World Use:**
- Calculate total cost (quantity × price)
- Weighted averages
- Conditional sums without SUMIFS

**Advanced Tricks:**
```excel
=SUMPRODUCT((Region="North")*(Product="Apple")*Sales)  → Multi-criteria sum
=SUMPRODUCT(1/COUNTIF(A1:A10,A1:A10))                  → Count unique values
```

**Tips:**
- Arrays must be same size
- Treats TRUE as 1, FALSE as 0
- More flexible than SUMIFS for complex criteria

---

## Rounding Functions

### ROUND
**Rounds to a specified number of digits**

**Syntax:** `=ROUND(number, num_digits)`

**Examples:**
```excel
=ROUND(2.567, 2)          → 2.57
=ROUND(2.567, 1)          → 2.6
=ROUND(2.567, 0)          → 3
=ROUND(1234.5, -1)        → 1230 (round to tens)
=ROUND(1234.5, -2)        → 1200 (round to hundreds)
```

**Real-World Use:**
- Financial calculations (2 decimals)
- Round to nearest dollar
- Scientific measurements

---

### ROUNDUP
**Always rounds up**

**Syntax:** `=ROUNDUP(number, num_digits)`

**Examples:**
```excel
=ROUNDUP(2.1, 0)          → 3
=ROUNDUP(2.567, 2)        → 2.57
=ROUNDUP(-2.1, 0)         → -3 (away from zero)
```

---

### ROUNDDOWN
**Always rounds down**

**Syntax:** `=ROUNDDOWN(number, num_digits)`

**Examples:**
```excel
=ROUNDDOWN(2.9, 0)        → 2
=ROUNDDOWN(2.567, 2)      → 2.56
=ROUNDDOWN(-2.9, 0)       → -2 (toward zero)
```

---

### MROUND
**Rounds to the nearest multiple**

**Syntax:** `=MROUND(number, multiple)`

**Examples:**
```excel
=MROUND(13, 5)            → 15 (nearest multiple of 5)
=MROUND(17, 5)            → 15
=MROUND(1.2, 0.5)         → 1.0
=MROUND(1.3, 0.5)         → 1.5
```

**Real-World Use:**
- Round to nearest $0.05 (nickel)
- Package quantities (boxes of 12)
- Time increments (15-minute intervals)

---

### CEILING
**Rounds up to nearest multiple**

**Syntax:** `=CEILING.MATH(number, [significance], [mode])`

**Examples:**
```excel
=CEILING.MATH(4.3)        → 5
=CEILING.MATH(4.3, 2)     → 6 (round up to multiple of 2)
=CEILING.MATH(-4.3)       → -4
=CEILING.MATH(-4.3,,1)    → -5 (away from zero)
```

**Real-World Use:**
- Minimum order quantities
- Round up to full units
- Price tiers

---

### FLOOR
**Rounds down to nearest multiple**

**Syntax:** `=FLOOR.MATH(number, [significance], [mode])`

**Examples:**
```excel
=FLOOR.MATH(4.9)          → 4
=FLOOR.MATH(4.9, 2)       → 4 (round down to multiple of 2)
=FLOOR.MATH(-4.3)         → -4
```

---

### INT
**Rounds down to nearest integer**

**Syntax:** `=INT(number)`

**Examples:**
```excel
=INT(8.9)                 → 8
=INT(-8.9)                → -9 (rounds down, not toward zero)
```

---

### TRUNC
**Truncates to an integer (removes decimal)**

**Syntax:** `=TRUNC(number, [num_digits])`

**Examples:**
```excel
=TRUNC(8.9)               → 8
=TRUNC(-8.9)              → -8 (removes decimal, toward zero)
=TRUNC(123.456, 2)        → 123.45
```

**INT vs TRUNC:**
- INT rounds down: `INT(-8.9) = -9`
- TRUNC toward zero: `TRUNC(-8.9) = -8`

---

## Trigonometric Functions

### SIN, COS, TAN
**Basic trigonometric functions**

**Syntax:** 
```excel
=SIN(angle_in_radians)
=COS(angle_in_radians)
=TAN(angle_in_radians)
```

**Examples:**
```excel
=SIN(RADIANS(30))         → 0.5
=COS(RADIANS(60))         → 0.5
=TAN(RADIANS(45))         → 1
=SIN(PI()/2)              → 1
```

**Real-World Use:**
- Engineering calculations
- Physics problems
- Circular motion

---

### ASIN, ACOS, ATAN
**Inverse trigonometric functions (returns radians)**

**Examples:**
```excel
=DEGREES(ASIN(0.5))       → 30 degrees
=DEGREES(ACOS(0.5))       → 60 degrees
=DEGREES(ATAN(1))         → 45 degrees
```

---

### ATAN2
**Returns arctangent of x and y coordinates**

**Syntax:** `=ATAN2(x_num, y_num)`

**Examples:**
```excel
=ATAN2(1, 1)              → Returns angle in radians
=DEGREES(ATAN2(1, 1))     → 45 degrees
```

**Use:** Calculate angle between two points

---

### RADIANS & DEGREES
**Convert between radians and degrees**

**Syntax:**
```excel
=RADIANS(degrees)
=DEGREES(radians)
```

**Examples:**
```excel
=RADIANS(180)             → 3.14159 (PI)
=DEGREES(PI())            → 180
```

---

### PI
**Returns the value of π**

**Syntax:** `=PI()`

**Examples:**
```excel
=PI()                     → 3.14159265358979
=2*PI()                   → 6.28318 (2π)
=PI()*5^2                 → 78.54 (area of circle, radius 5)
```

---

## Advanced Math

### POWER
**Returns a number raised to a power**

**Syntax:** `=POWER(number, power)`

**Examples:**
```excel
=POWER(2, 3)              → 8 (2³)
=POWER(5, 2)              → 25 (5²)
=POWER(4, 0.5)            → 2 (square root)
```

**Alternative:** Use ^ operator
```excel
=2^3                      → 8
=5^2                      → 25
```

---

### SQRT
**Returns the square root**

**Syntax:** `=SQRT(number)`

**Examples:**
```excel
=SQRT(16)                 → 4
=SQRT(2)                  → 1.414
=SQRT(A1^2 + B1^2)        → Pythagorean theorem
```

---

### EXP
**Returns e raised to a power**

**Syntax:** `=EXP(number)`

**Examples:**
```excel
=EXP(1)                   → 2.71828 (e)
=EXP(2)                   → 7.389
```

**Use:** Growth models, compound interest

---

### LN & LOG
**Natural logarithm and base-10 logarithm**

**Syntax:**
```excel
=LN(number)               → Natural log (base e)
=LOG(number, [base])      → Log with specified base
=LOG10(number)            → Log base 10
```

**Examples:**
```excel
=LN(2.71828)              → 1 (ln(e) = 1)
=LOG(100)                 → 2 (log₁₀(100))
=LOG(8, 2)                → 3 (log₂(8))
```

---

### ABS
**Returns absolute value**

**Syntax:** `=ABS(number)`

**Examples:**
```excel
=ABS(-5)                  → 5
=ABS(5)                   → 5
=ABS(A1-B1)               → Difference regardless of order
```

**Real-World Use:**
- Calculate variance
- Distance calculations
- Error margins

---

### SIGN
**Returns the sign of a number**

**Syntax:** `=SIGN(number)`

**Returns:**
- 1 if positive
- 0 if zero
- -1 if negative

**Examples:**
```excel
=SIGN(10)                 → 1
=SIGN(-10)                → -1
=SIGN(0)                  → 0
```

---

### MOD
**Returns the remainder after division**

**Syntax:** `=MOD(number, divisor)`

**Examples:**
```excel
=MOD(10, 3)               → 1 (10 ÷ 3 = 3 remainder 1)
=MOD(15, 4)               → 3
=MOD(A1, 2)               → 0 if even, 1 if odd
```

**Real-World Use:**
- Check if even/odd
- Cycle through values
- Stripe row colors: `=MOD(ROW(),2)=0`

---

### QUOTIENT
**Returns integer portion of division**

**Syntax:** `=QUOTIENT(numerator, denominator)`

**Examples:**
```excel
=QUOTIENT(10, 3)          → 3
=QUOTIENT(15, 4)          → 3
```

---

### GCD & LCM
**Greatest common divisor and least common multiple**

**Syntax:**
```excel
=GCD(number1, [number2], ...)
=LCM(number1, [number2], ...)
```

**Examples:**
```excel
=GCD(12, 18)              → 6
=LCM(12, 18)              → 36
```

---

### FACT & FACTDOUBLE
**Factorial functions**

**Syntax:**
```excel
=FACT(number)             → n! = n × (n-1) × ... × 1
=FACTDOUBLE(number)       → n!! (double factorial)
```

**Examples:**
```excel
=FACT(5)                  → 120 (5! = 5×4×3×2×1)
=FACTDOUBLE(6)            → 48 (6!! = 6×4×2)
```

---

### COMBIN & COMBINA
**Combinatorics functions**

**Syntax:**
```excel
=COMBIN(number, number_chosen)        → Combinations
=COMBINA(number, number_chosen)       → Combinations with repetition
```

**Examples:**
```excel
=COMBIN(5, 2)             → 10 (ways to choose 2 from 5)
=COMBINA(5, 2)            → 15 (with replacement)
```

---

### RAND & RANDBETWEEN
**Generate random numbers**

**Syntax:**
```excel
=RAND()                   → Random between 0 and 1
=RANDBETWEEN(bottom, top) → Random integer in range
```

**Examples:**
```excel
=RAND()                   → 0.234567 (changes on recalc)
=RAND()*100               → Random 0 to 100
=RANDBETWEEN(1, 10)       → Random integer 1-10
=RANDBETWEEN(1, 6)        → Simulate dice roll
```

**Real-World Use:**
- Simulations
- Random sampling
- Testing formulas

**Tips:**
- Recalculates every time worksheet changes
- Press F9 to force recalculation
- Use with volatile functions carefully (slows workbook)

---

### SUMX2MY2, SUMX2PY2, SUMXMY2
**Sum of products of differences**

**Syntax:**
```excel
=SUMX2MY2(array_x, array_y)  → Sum of (x² - y²)
=SUMX2PY2(array_x, array_y)  → Sum of (x² + y²)
=SUMXMY2(array_x, array_y)   → Sum of (x - y)²
```

**Use:** Statistical calculations

---

### MULTINOMIAL
**Returns the multinomial coefficient**

**Syntax:** `=MULTINOMIAL(number1, [number2], ...)`

**Example:**
```excel
=MULTINOMIAL(2, 3, 4)     → 1260
```

---

## Matrix Operations

### MMULT
**Matrix multiplication**

**Syntax:** `=MMULT(array1, array2)`

**Requirements:**
- Number of columns in array1 = number of rows in array2
- Enter as array formula (Ctrl+Shift+Enter in older Excel)

**Use:** Linear algebra, transformations

---

### MDETERM
**Matrix determinant**

**Syntax:** `=MDETERM(array)`

**Use:** Solve systems of equations

---

### MINVERSE
**Matrix inverse**

**Syntax:** `=MINVERSE(array)`

**Use:** Solve linear equations

---

## Quick Reference Table

| Function | Purpose | Example |
|----------|---------|---------|
| SUM | Add values | `=SUM(A1:A10)` |
| SUMIF | Conditional sum | `=SUMIF(A:A,">100",B:B)` |
| SUMIFS | Multi-criteria sum | `=SUMIFS(D:D,A:A,"North",B:B,"Jan")` |
| AVERAGE | Mean | `=AVERAGE(A1:A10)` |
| COUNT | Count numbers | `=COUNT(A1:A10)` |
| ROUND | Round number | `=ROUND(A1,2)` |
| SUMPRODUCT | Multiply & sum | `=SUMPRODUCT(A1:A5,B1:B5)` |
| ABS | Absolute value | `=ABS(-10)` |
| MOD | Remainder | `=MOD(10,3)` |
| RAND | Random number | `=RAND()` |

---

## Common Use Cases

### Calculate Total with Tax
```excel
=SUM(A1:A10) * 1.08          // Add 8% tax
=SUMPRODUCT(A1:A10, 1.08)    // Alternative
```

### Running Total
```excel
=SUM($A$1:A1)                // Copy down for cumulative sum
```

### Weighted Average
```excel
=SUMPRODUCT(Values, Weights) / SUM(Weights)
```

### Count Unique Values
```excel
=SUMPRODUCT(1/COUNTIF(A1:A10, A1:A10))
```

### Round to Nearest 5 Cents
```excel
=MROUND(A1, 0.05)
```

### Check if Even or Odd
```excel
=IF(MOD(A1,2)=0, "Even", "Odd")
```

---

## Tips & Best Practices

### Performance
- Use SUM instead of adding cells individually
- SUMIFS is faster than SUMPRODUCT for simple criteria
- Avoid volatile functions (RAND, NOW) in large workbooks

### Accuracy
- Use ROUND for display, keep full precision in calculations
- Be aware of floating-point errors in very large/small numbers
- Use MROUND for currency to avoid rounding errors

### Common Errors
- `#DIV/0!` - Division by zero (use IFERROR)
- `#VALUE!` - Wrong data type or array size mismatch
- `#NUM!` - Invalid numeric value

### Best Practices
- Name ranges for clarity: `=SUM(Sales_Q1)`
- Use structured references in tables: `=SUM(Table1[Sales])`
- Document complex formulas with comments
- Test with edge cases (zero, negative, very large)

---

**[⬆ Back to Main README](../../README.md)**
