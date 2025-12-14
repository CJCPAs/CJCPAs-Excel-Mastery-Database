# Information Functions

> **Check cell contents, types, and properties**

## Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| **ISBLANK** | Is cell empty? | `=ISBLANK(A1)` → TRUE/FALSE |
| **ISERROR** | Is any error? | `=ISERROR(A1)` |
| **ISERR** | Is error (not #N/A)? | `=ISERR(A1)` |
| **ISNA** | Is #N/A error? | `=ISNA(A1)` |
| **ISNUMBER** | Is number? | `=ISNUMBER(A1)` |
| **ISTEXT** | Is text? | `=ISTEXT(A1)` |
| **ISLOGICAL** | Is TRUE/FALSE? | `=ISLOGICAL(A1)` |
| **ISREF** | Is valid reference? | `=ISREF(A1)` |
| **ISFORMULA** | Contains formula? | `=ISFORMULA(A1)` |
| **ISODD** | Is odd number? | `=ISODD(A1)` |
| **ISEVEN** | Is even number? | `=ISEVEN(A1)` |
| **TYPE** | Value type code | `=TYPE(A1)` → 1=number |
| **ERROR.TYPE** | Error type code | `=ERROR.TYPE(A1)` |
| **NA** | Return #N/A | `=NA()` |
| **CELL** | Cell information | `=CELL("type",A1)` |
| **INFO** | System information | `=INFO("osversion")` |
| **N** | Convert to number | `=N(A1)` |
| **SHEET** | Sheet number | `=SHEET(A1)` |
| **SHEETS** | Number of sheets | `=SHEETS()` |

## Type Codes (TYPE function)
| Code | Meaning |
|------|---------|
| 1 | Number |
| 2 | Text |
| 4 | Logical |
| 16 | Error |
| 64 | Array |

## Common Solutions

### Check Before Calculate
```excel
=IF(ISNUMBER(A1), A1*2, 0)
```

### Handle Any Error
```excel
=IF(ISERROR(formula), "Error", formula)
```

### Validate Data Type
```excel
=IF(AND(ISNUMBER(A1), A1>0), "Valid", "Invalid")
```

---

[🏠 Back to Home](../../README.md)
