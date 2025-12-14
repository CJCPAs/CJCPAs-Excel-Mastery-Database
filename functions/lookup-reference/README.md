# Lookup & Reference Functions

> **30+ functions for finding, referencing, and retrieving data - the most powerful tools in Excel**

## Quick Navigation

| I want to... | Use this function |
|--------------|-------------------|
| Look up a value | [VLOOKUP](./VLOOKUP.md), [XLOOKUP](./XLOOKUP.md) |
| Look up with more flexibility | [INDEX](./INDEX.md) + [MATCH](./MATCH.md) |
| Find position of a value | [MATCH](./MATCH.md), [XMATCH](./XMATCH.md) |
| Filter data by criteria | [FILTER](./FILTER.md) |
| Sort data | [SORT](./SORT.md), [SORTBY](./SORTBY.md) |
| Get unique values | [UNIQUE](./UNIQUE.md) |
| Create dynamic reference | [OFFSET](./OFFSET.md), [INDIRECT](./INDIRECT.md) |

---

## Recommended Learning Path

### Beginner: Start Here
1. **[VLOOKUP](./VLOOKUP.md)** - The classic lookup function everyone should know
2. **[MATCH](./MATCH.md)** - Find position of values

### Intermediate: Level Up
3. **[INDEX](./INDEX.md) + [MATCH](./MATCH.md)** - The powerful combo that replaces VLOOKUP
4. **[XLOOKUP](./XLOOKUP.md)** - Modern lookup (if you have Excel 365/2021)

### Advanced: Dynamic Arrays
5. **[FILTER](./FILTER.md)** - Extract matching rows dynamically
6. **[SORT](./SORT.md) / [UNIQUE](./UNIQUE.md)** - Sort and deduplicate data

---

## All Lookup & Reference Functions

### Core Lookup Functions
| Function | Description | Example |
|----------|-------------|---------|
| [VLOOKUP](./VLOOKUP.md) | Vertical lookup (most common) | `=VLOOKUP("A1",Table,2,FALSE)` |
| [HLOOKUP](./HLOOKUP.md) | Horizontal lookup | `=HLOOKUP("Q1",A1:D5,3,FALSE)` |
| [XLOOKUP](./XLOOKUP.md) | Modern lookup (365/2021) | `=XLOOKUP(A1,IDs,Names)` |
| [LOOKUP](./LOOKUP.md) | Simple lookup (sorted data) | `=LOOKUP(75,A:A,B:B)` |
| [INDEX](./INDEX.md) | Return value at position | `=INDEX(A1:C10,5,2)` |
| [MATCH](./MATCH.md) | Find position of value | `=MATCH("X",A:A,0)` |
| [XMATCH](./XMATCH.md) | Modern MATCH (365/2021) | `=XMATCH("X",A:A)` |

### Dynamic Array Functions (Excel 365/2021)
| Function | Description | Example |
|----------|-------------|---------|
| [FILTER](./FILTER.md) | Filter rows by criteria | `=FILTER(A:C,B:B>100)` |
| [SORT](./SORT.md) | Sort data | `=SORT(A1:C10,2,-1)` |
| [SORTBY](./SORTBY.md) | Sort by another column | `=SORTBY(Names,Scores,-1)` |
| [UNIQUE](./UNIQUE.md) | Extract unique values | `=UNIQUE(A:A)` |
| [SEQUENCE](./SEQUENCE.md) | Generate number sequence | `=SEQUENCE(10,1,1,1)` |
| [RANDARRAY](./RANDARRAY.md) | Array of random numbers | `=RANDARRAY(5,5,1,100,TRUE)` |

### Advanced Array Functions (Excel 365/2021)
| Function | Description | Example |
|----------|-------------|---------|
| TAKE | Take first/last N rows | `=TAKE(Data,5)` |
| DROP | Drop first/last N rows | `=DROP(Data,1)` |
| CHOOSECOLS | Select specific columns | `=CHOOSECOLS(Data,1,3,5)` |
| CHOOSEROWS | Select specific rows | `=CHOOSEROWS(Data,1,2,5)` |
| VSTACK | Stack arrays vertically | `=VSTACK(A1:A5,C1:C5)` |
| HSTACK | Stack arrays horizontally | `=HSTACK(A1:A5,C1:C5)` |
| WRAPROWS | Wrap into rows | `=WRAPROWS(A1:L1,4)` |
| WRAPCOLS | Wrap into columns | `=WRAPCOLS(A1:A12,4)` |
| TOROW | Convert to single row | `=TOROW(A1:C3)` |
| TOCOL | Convert to single column | `=TOCOL(A1:C3)` |
| EXPAND | Expand array dimensions | `=EXPAND(A1:B2,5,5)` |

### Reference Functions
| Function | Description | Example |
|----------|-------------|---------|
| [OFFSET](./OFFSET.md) | Dynamic reference offset | `=OFFSET(A1,5,2)` |
| [INDIRECT](./INDIRECT.md) | Reference from text | `=INDIRECT("A"&B1)` |
| [ADDRESS](./ADDRESS.md) | Create cell address | `=ADDRESS(5,3)` → "$C$5" |
| [ROW](./ROW.md) | Row number | `=ROW()` |
| [COLUMN](./COLUMN.md) | Column number | `=COLUMN()` |
| [ROWS](./ROWS.md) | Count rows in range | `=ROWS(A1:A10)` → 10 |
| [COLUMNS](./COLUMNS.md) | Count columns in range | `=COLUMNS(A1:D1)` → 4 |
| [AREAS](./AREAS.md) | Count areas in reference | `=AREAS((A1:B2,D1:E2))` → 2 |

### Other Reference Functions
| Function | Description | Example |
|----------|-------------|---------|
| [CHOOSE](./CHOOSE.md) | Return nth item from list | `=CHOOSE(2,"A","B","C")` → "B" |
| [TRANSPOSE](./TRANSPOSE.md) | Flip rows and columns | `=TRANSPOSE(A1:C3)` |
| [FORMULATEXT](./FORMULATEXT.md) | Show formula as text | `=FORMULATEXT(A1)` |
| [HYPERLINK](./HYPERLINK.md) | Create clickable link | `=HYPERLINK("url","Click")` |
| [GETPIVOTDATA](./GETPIVOTDATA.md) | Get data from PivotTable | Auto-generated |

---

## Common Patterns

### Basic Lookup (VLOOKUP)
```excel
=VLOOKUP(lookup_value, table, column_num, FALSE)
```
Always use FALSE for exact match!

### Better Lookup (INDEX/MATCH)
```excel
=INDEX(return_column, MATCH(lookup_value, lookup_column, 0))
```
More flexible than VLOOKUP - can look left!

### Modern Lookup (XLOOKUP) - Excel 365/2021
```excel
=XLOOKUP(lookup_value, lookup_array, return_array, "Not found")
```
Built-in error handling!

### Two-Way Lookup
```excel
=INDEX(DataRange, MATCH(RowValue, RowHeaders, 0), MATCH(ColValue, ColHeaders, 0))
```

### Multiple Criteria Lookup
```excel
=INDEX(Return, MATCH(1, (Crit1=Range1)*(Crit2=Range2), 0))
```

### Filter and Sort
```excel
=SORT(FILTER(Data, Criteria), SortColumn, -1)
```

---

## Tips & Best Practices

1. **XLOOKUP > VLOOKUP** when available (Excel 365/2021)
2. **INDEX/MATCH > VLOOKUP** for flexibility and performance
3. **Use exact match (FALSE or 0)** unless you specifically need approximate
4. **Named ranges** make formulas readable: `=VLOOKUP(ID, Employees, 3, FALSE)`
5. **IFERROR for legacy** - wrap VLOOKUP in IFERROR for error handling
6. **Avoid OFFSET/INDIRECT** in large workbooks - they're volatile and slow

---

[🏠 Back to Home](../../README.md) | [📖 All Functions A-Z](../../13-Quick-Reference/Functions-A-Z.md)
