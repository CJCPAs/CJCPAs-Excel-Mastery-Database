# Dynamic Array Functions (Excel 365)

> **Filter, sort, and transform data with spilling arrays**

## Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| **FILTER** | Filter rows | `=FILTER(A:D, B:B="North")` |
| **SORT** | Sort data | `=SORT(A1:C10, 2, -1)` |
| **SORTBY** | Sort by other column | `=SORTBY(A:B, C:C, -1)` |
| **UNIQUE** | Unique values | `=UNIQUE(A2:A100)` |
| **SEQUENCE** | Number series | `=SEQUENCE(10,1,1,1)` |
| **RANDARRAY** | Random array | `=RANDARRAY(5,3,1,100)` |
| **XLOOKUP** | Modern lookup | `=XLOOKUP(val,lookup,return)` |
| **XMATCH** | Modern match | `=XMATCH(val,array)` |
| **CHOOSECOLS** | Select columns | `=CHOOSECOLS(A:D,1,3)` |
| **CHOOSEROWS** | Select rows | `=CHOOSEROWS(A1:D10,1,5)` |
| **TAKE** | First/last N | `=TAKE(array,5)` |
| **DROP** | Remove first/last N | `=DROP(array,2)` |
| **EXPAND** | Expand array | `=EXPAND(A1:B2,5,5)` |
| **VSTACK** | Stack vertically | `=VSTACK(A1:A5,B1:B3)` |
| **HSTACK** | Stack horizontally | `=HSTACK(A1:A5,B1:B5)` |
| **WRAPCOLS** | Wrap to columns | `=WRAPCOLS(A1:A12,4)` |
| **WRAPROWS** | Wrap to rows | `=WRAPROWS(A1:L1,4)` |
| **TOCOL** | Convert to column | `=TOCOL(A1:D3)` |
| **TOROW** | Convert to row | `=TOROW(A1:A10)` |
| **TEXTSPLIT** | Split text | `=TEXTSPLIT(A1,",")` |

## Common Patterns

### Top 5 by Value
```excel
=TAKE(SORT(A1:C100, 3, -1), 5)
```

### Unique Sorted List
```excel
=SORT(UNIQUE(A2:A100))
```

### Filter Multiple Criteria
```excel
=FILTER(Data, (Region="North")*(Sales>1000))
```

### Generate Date Series
```excel
=SEQUENCE(30,1,TODAY(),1)
```

---

[📚 Full Advanced Solutions](../../solutions/advanced/) | [🏠 Back to Home](../../README.md)
