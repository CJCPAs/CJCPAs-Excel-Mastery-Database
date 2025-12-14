# Logical Functions

> **15+ functions for conditional logic, decision-making, and error handling**

## Quick Navigation

| I want to... | Use this function |
|--------------|-------------------|
| Make a decision based on condition | [IF](./IF.md) |
| Check multiple conditions | [IFS](./IFS.md), [IF](./IF.md) with [AND](./AND.md)/[OR](./OR.md) |
| Handle errors gracefully | [IFERROR](./IFERROR.md), [IFNA](./IFNA.md) |
| Match a value to options | [SWITCH](./SWITCH.md) |
| Test if all conditions are true | [AND](./AND.md) |
| Test if any condition is true | [OR](./OR.md) |
| Create reusable formulas | [LET](./LET.md), [LAMBDA](./LAMBDA.md) |

---

## All Logical Functions

### Core Decision Functions
| Function | Description | Example |
|----------|-------------|---------|
| [IF](./IF.md) | Returns value based on condition | `=IF(A1>10,"High","Low")` |
| [IFS](./IFS.md) | Multiple IF conditions (no nesting) | `=IFS(A1>=90,"A",A1>=80,"B",TRUE,"C")` |
| [SWITCH](./SWITCH.md) | Match value against options | `=SWITCH(A1,1,"Jan",2,"Feb",3,"Mar")` |

### Error Handling
| Function | Description | Example |
|----------|-------------|---------|
| [IFERROR](./IFERROR.md) | Catch any error | `=IFERROR(A1/B1,"Error")` |
| [IFNA](./IFNA.md) | Catch only #N/A | `=IFNA(VLOOKUP(...),"Not found")` |

### Condition Testing
| Function | Description | Example |
|----------|-------------|---------|
| [AND](./AND.md) | TRUE if all conditions TRUE | `=AND(A1>10,B1<20)` |
| [OR](./OR.md) | TRUE if any condition TRUE | `=OR(A1="Yes",A1="Y")` |
| [NOT](./NOT.md) | Reverses TRUE/FALSE | `=NOT(A1="Complete")` |
| [XOR](./XOR.md) | TRUE if odd number TRUE | `=XOR(A1,B1,C1)` |

### Boolean Values
| Function | Description | Example |
|----------|-------------|---------|
| [TRUE](./TRUE.md) | Returns TRUE | `=TRUE()` |
| [FALSE](./FALSE.md) | Returns FALSE | `=FALSE()` |

### Advanced Functions (Excel 365)
| Function | Description | Example |
|----------|-------------|---------|
| [LET](./LET.md) | Define named calculations | `=LET(x,A1*2,y,x+5,y*2)` |
| [LAMBDA](./LAMBDA.md) | Create custom functions | `=LAMBDA(x,x^2)` |
| [MAP](./MAP.md) | Apply function to array | `=MAP(A1:A10,LAMBDA(x,x*2))` |
| [REDUCE](./REDUCE.md) | Reduce array to single value | `=REDUCE(0,A1:A5,LAMBDA(a,b,a+b))` |
| [SCAN](./SCAN.md) | Cumulative calculation | `=SCAN(0,A1:A5,LAMBDA(a,b,a+b))` |
| [MAKEARRAY](./MAKEARRAY.md) | Generate array with function | Complex |
| [BYROW](./BYROW.md) | Apply function to each row | `=BYROW(Data,LAMBDA(r,SUM(r)))` |
| [BYCOL](./BYCOL.md) | Apply function to each column | `=BYCOL(Data,LAMBDA(c,MAX(c)))` |

---

## Common Patterns

### Simple If-Then-Else
```excel
=IF(condition, value_if_true, value_if_false)
```

### Multiple Conditions (AND)
```excel
=IF(AND(A1>10, B1<20), "Yes", "No")
```

### Multiple Conditions (OR)
```excel
=IF(OR(A1="Red", A1="Blue"), "Primary", "Other")
```

### Nested IF (Avoid if Possible)
```excel
=IF(A1>=90, "A", IF(A1>=80, "B", IF(A1>=70, "C", "D")))
```

### Better: Use IFS (Excel 2019+)
```excel
=IFS(A1>=90, "A", A1>=80, "B", A1>=70, "C", TRUE, "D")
```

### Safe Division
```excel
=IFERROR(A1/B1, 0)
```

### Safe Lookup
```excel
=IFERROR(VLOOKUP(A1, Table, 2, FALSE), "Not found")
```

### LET for Readable Formulas
```excel
=LET(
    sales, SUM(B:B),
    costs, SUM(C:C),
    profit, sales - costs,
    margin, profit / sales,
    IF(margin > 0.2, "Good", "Review")
)
```

---

## Logic Truth Tables

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

### NOT Truth Table
| A | NOT(A) |
|---|--------|
| TRUE | FALSE |
| FALSE | TRUE |

---

## Tips & Best Practices

1. **Keep IFs simple:** If nesting more than 3 levels, consider IFS, SWITCH, or lookup table
2. **Use IFS/SWITCH** when available (Excel 2019+) for cleaner multi-condition logic
3. **Order matters in IFS:** First TRUE condition wins - put most specific first
4. **IFERROR hides problems:** Only use when you're sure you want to hide all errors
5. **LET improves readability:** Name intermediate calculations for complex formulas
6. **Test conditions separately:** Build complex AND/OR conditions in helper cells first

---

[🏠 Back to Home](../../README.md) | [📖 All Functions A-Z](../../13-Quick-Reference/Functions-A-Z.md)
