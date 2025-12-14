# Math & Trigonometry Functions

> **80+ functions for mathematical calculations, rounding operations, and trigonometric computations**

## Quick Navigation

| I want to... | Use this function |
|--------------|-------------------|
| Add numbers | [SUM](./SUM.md) |
| Add with conditions | [SUMIF](./SUMIF.md), [SUMIFS](./SUMIFS.md) |
| Find average | [AVERAGE](./AVERAGE.md) |
| Count numbers | [COUNT](./COUNT.md) |
| Round a number | [ROUND](./ROUND.md), [ROUNDUP](./ROUNDUP.md), [ROUNDDOWN](./ROUNDDOWN.md) |
| Multiply and sum arrays | [SUMPRODUCT](./SUMPRODUCT.md) |
| Generate random numbers | [RAND](./RAND.md), [RANDBETWEEN](./RANDBETWEEN.md) |

---

## All Math & Trig Functions

### Basic Aggregation
| Function | Description | Example |
|----------|-------------|---------|
| [SUM](./SUM.md) | Adds all numbers | `=SUM(A1:A10)` → total |
| [SUMIF](./SUMIF.md) | Sum with one condition | `=SUMIF(A:A,"North",B:B)` |
| [SUMIFS](./SUMIFS.md) | Sum with multiple conditions | `=SUMIFS(C:C,A:A,"North",B:B,">100")` |
| [SUMPRODUCT](./SUMPRODUCT.md) | Multiply arrays and sum | `=SUMPRODUCT(A1:A5,B1:B5)` |
| [AVERAGE](./AVERAGE.md) | Arithmetic mean | `=AVERAGE(A1:A10)` |
| [PRODUCT](./PRODUCT.md) | Multiply all numbers | `=PRODUCT(A1:A5)` |
| [SUBTOTAL](./SUBTOTAL.md) | Aggregate ignoring filters | `=SUBTOTAL(9,A1:A10)` |
| [AGGREGATE](./AGGREGATE.md) | Flexible aggregation | `=AGGREGATE(9,6,A1:A10)` |

### Counting
| Function | Description | Example |
|----------|-------------|---------|
| [COUNT](./COUNT.md) | Count numeric cells | `=COUNT(A1:A100)` |
| [COUNTA](./COUNTA.md) | Count non-empty cells | `=COUNTA(A1:A100)` |
| [COUNTBLANK](./COUNTBLANK.md) | Count empty cells | `=COUNTBLANK(A1:A100)` |
| [COUNTIF](./COUNTIF.md) | Count with condition | `=COUNTIF(A:A,">100")` |
| [COUNTIFS](./COUNTIFS.md) | Count with multiple conditions | `=COUNTIFS(A:A,">100",B:B,"<200")` |

### Rounding Functions
| Function | Description | Example |
|----------|-------------|---------|
| [ROUND](./ROUND.md) | Round to digits | `=ROUND(3.567,2)` → 3.57 |
| [ROUNDUP](./ROUNDUP.md) | Always round up | `=ROUNDUP(3.21,1)` → 3.3 |
| [ROUNDDOWN](./ROUNDDOWN.md) | Always round down | `=ROUNDDOWN(3.89,1)` → 3.8 |
| [MROUND](./MROUND.md) | Round to multiple | `=MROUND(17,5)` → 15 |
| [CEILING](./CEILING.md) | Round up to multiple | `=CEILING(4.2,1)` → 5 |
| [CEILING.MATH](./CEILING.MATH.md) | Ceiling with options | `=CEILING.MATH(-4.2,1)` |
| [FLOOR](./FLOOR.md) | Round down to multiple | `=FLOOR(4.8,1)` → 4 |
| [FLOOR.MATH](./FLOOR.MATH.md) | Floor with options | `=FLOOR.MATH(-4.8,1)` |
| [INT](./INT.md) | Round to integer (down) | `=INT(4.9)` → 4 |
| [TRUNC](./TRUNC.md) | Truncate to digits | `=TRUNC(4.9)` → 4 |
| [EVEN](./EVEN.md) | Round up to even | `=EVEN(3)` → 4 |
| [ODD](./ODD.md) | Round up to odd | `=ODD(4)` → 5 |

### Basic Math Operations
| Function | Description | Example |
|----------|-------------|---------|
| [ABS](./ABS.md) | Absolute value | `=ABS(-5)` → 5 |
| [SIGN](./SIGN.md) | Sign of number | `=SIGN(-10)` → -1 |
| [MOD](./MOD.md) | Remainder after division | `=MOD(10,3)` → 1 |
| [QUOTIENT](./QUOTIENT.md) | Integer part of division | `=QUOTIENT(10,3)` → 3 |
| [POWER](./POWER.md) | Number raised to power | `=POWER(2,3)` → 8 |
| [SQRT](./SQRT.md) | Square root | `=SQRT(16)` → 4 |
| [SQRTPI](./SQRTPI.md) | Square root of π×n | `=SQRTPI(2)` |

### Exponential & Logarithmic
| Function | Description | Example |
|----------|-------------|---------|
| [EXP](./EXP.md) | e raised to power | `=EXP(1)` → 2.718... |
| [LN](./LN.md) | Natural logarithm | `=LN(2.718)` → 1 |
| [LOG](./LOG.md) | Logarithm (any base) | `=LOG(100,10)` → 2 |
| [LOG10](./LOG10.md) | Base-10 logarithm | `=LOG10(100)` → 2 |

### Random Numbers
| Function | Description | Example |
|----------|-------------|---------|
| [RAND](./RAND.md) | Random decimal 0-1 | `=RAND()` → 0.xxx |
| [RANDBETWEEN](./RANDBETWEEN.md) | Random integer in range | `=RANDBETWEEN(1,100)` |
| [RANDARRAY](./RANDARRAY.md) | Array of random numbers | `=RANDARRAY(5,5,1,100,TRUE)` |

### Trigonometric Functions
| Function | Description | Example |
|----------|-------------|---------|
| [PI](./PI.md) | Value of π | `=PI()` → 3.14159... |
| [SIN](./SIN.md) | Sine (radians) | `=SIN(PI()/2)` → 1 |
| [COS](./COS.md) | Cosine (radians) | `=COS(0)` → 1 |
| [TAN](./TAN.md) | Tangent (radians) | `=TAN(PI()/4)` → 1 |
| [ASIN](./ASIN.md) | Arcsine (returns radians) | `=ASIN(0.5)` |
| [ACOS](./ACOS.md) | Arccosine | `=ACOS(0.5)` |
| [ATAN](./ATAN.md) | Arctangent | `=ATAN(1)` |
| [ATAN2](./ATAN2.md) | Arctangent from x,y | `=ATAN2(1,1)` |
| [RADIANS](./RADIANS.md) | Convert degrees to radians | `=RADIANS(180)` → π |
| [DEGREES](./DEGREES.md) | Convert radians to degrees | `=DEGREES(PI())` → 180 |

### Matrix Functions
| Function | Description | Example |
|----------|-------------|---------|
| [MMULT](./MMULT.md) | Matrix multiplication | `=MMULT(A1:B2,D1:E2)` |
| [MDETERM](./MDETERM.md) | Matrix determinant | `=MDETERM(A1:C3)` |
| [MINVERSE](./MINVERSE.md) | Matrix inverse | `=MINVERSE(A1:C3)` |

### Combinatorics
| Function | Description | Example |
|----------|-------------|---------|
| [FACT](./FACT.md) | Factorial | `=FACT(5)` → 120 |
| [FACTDOUBLE](./FACTDOUBLE.md) | Double factorial | `=FACTDOUBLE(6)` → 48 |
| [COMBIN](./COMBIN.md) | Combinations | `=COMBIN(10,3)` → 120 |
| [COMBINA](./COMBINA.md) | Combinations with repetition | `=COMBINA(5,3)` |
| [PERMUT](./PERMUT.md) | Permutations | `=PERMUT(10,3)` → 720 |
| [GCD](./GCD.md) | Greatest common divisor | `=GCD(24,36)` → 12 |
| [LCM](./LCM.md) | Least common multiple | `=LCM(4,6)` → 12 |

### Advanced Math
| Function | Description | Example |
|----------|-------------|---------|
| [SERIESSUM](./SERIESSUM.md) | Sum of power series | Complex scientific |
| [MULTINOMIAL](./MULTINOMIAL.md) | Multinomial coefficient | `=MULTINOMIAL(2,3,4)` |
| [SUMSQ](./SUMSQ.md) | Sum of squares | `=SUMSQ(1,2,3)` → 14 |
| [SUMX2MY2](./SUMX2MY2.md) | Sum of x²-y² | Statistical use |
| [SUMX2PY2](./SUMX2PY2.md) | Sum of x²+y² | Statistical use |
| [SUMXMY2](./SUMXMY2.md) | Sum of (x-y)² | Statistical use |

---

## Common Formulas

### Running Total
```excel
=SUM($A$2:A2)
```
Copy down - creates cumulative sum.

### Weighted Average
```excel
=SUMPRODUCT(Scores, Weights)/SUM(Weights)
```

### Percentage of Total
```excel
=A2/SUM($A$2:$A$10)
```

### Count Unique Values
```excel
=SUMPRODUCT(1/COUNTIF(A2:A100,A2:A100))
```

### Check If Even or Odd
```excel
=IF(MOD(A1,2)=0, "Even", "Odd")
```

### Round to Nearest 5
```excel
=MROUND(A1, 5)
```

### Calculate Tax
```excel
=SUM(Subtotals) * TaxRate
```

---

## Tips & Best Practices

1. **Use SUM over + operator** for ranges: `=SUM(A1:A100)` is cleaner and faster than `=A1+A2+...+A100`

2. **SUMIFS > SUMPRODUCT** for simple conditions (better performance)

3. **AGGREGATE** ignores errors and hidden rows - use function code 9 for SUM, code 6 to ignore errors

4. **Avoid volatile functions** in large workbooks: RAND, RANDBETWEEN, and NOW recalculate constantly

5. **Name your ranges** for clearer formulas: `=SUM(MonthlySales)` beats `=SUM(B2:B13)`

---

[🏠 Back to Home](../../README.md) | [📖 All Functions A-Z](../../13-Quick-Reference/Functions-A-Z.md)
