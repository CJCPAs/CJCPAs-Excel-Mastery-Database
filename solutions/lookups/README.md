# Lookup & Data Retrieval Solutions

> **Find and retrieve data from tables - the most common Excel task**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Look up a value | [VLOOKUP](#basic-lookup) or [XLOOKUP](#modern-lookup-xlookup) |
| Look up with multiple criteria | [Multiple Criteria Lookup](#lookup-with-multiple-criteria) |
| Look up to the LEFT | [INDEX/MATCH](#left-lookup-indexmatch) or XLOOKUP |
| Return multiple columns | [XLOOKUP](#return-multiple-columns) or FILTER |
| Find all matches | [FILTER](#return-all-matches) |
| Handle lookup errors | [Error Handling](#handle-lookup-errors) |
| Two-way lookup | [Row and Column Match](#two-way-lookup) |

---

## Basic Lookup

### The Challenge
You have a product ID and need to find its name, price, or other information from a table.

### Quick Answer
```excel
=VLOOKUP(lookup_value, table, column_number, FALSE)
```

### Full Example

**Lookup Table (A1:C5):**
| A | B | C |
|---|---|---|
| Product ID | Name | Price |
| P001 | Laptop | 999 |
| P002 | Mouse | 29 |
| P003 | Keyboard | 79 |
| P004 | Monitor | 349 |

**Looking up in E2:**
| E | F |
|---|---|
| Find Price for: | P002 |

**Formula:** `=VLOOKUP(F1, A2:C5, 3, FALSE)`

**Result:** `29`

**Explanation:**
1. Search for "P002" in first column (A)
2. When found, return value from column 3 (Price)
3. FALSE means exact match required

---

## Modern Lookup (XLOOKUP)

### The Challenge
Need a more powerful, flexible lookup than VLOOKUP.

### Quick Answer
```excel
=XLOOKUP(lookup_value, lookup_range, return_range, "Not found")
```

### Full Example

**Same table as above**

**Formula:** `=XLOOKUP("P002", A2:A5, C2:C5, "Product not found")`

**Result:** `29`

### Why XLOOKUP is Better

| Advantage | XLOOKUP | VLOOKUP |
|-----------|---------|---------|
| Look left | ✅ Yes | ❌ No |
| Error handling | ✅ Built-in | ❌ Need IFERROR |
| Return multiple columns | ✅ Yes | ❌ No |
| Column insertion safe | ✅ Yes | ❌ No |

**Note:** XLOOKUP requires Excel 365 or 2021

---

## Left Lookup (INDEX/MATCH)

### The Challenge
Your lookup column is to the RIGHT of the data you need (VLOOKUP can't do this).

### Quick Answer
```excel
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```

### Full Example

**Data (A1:C5):**
| A | B | C |
|---|---|---|
| Name | Department | Employee ID |
| Alice | Sales | E001 |
| Bob | Marketing | E002 |
| Carol | Finance | E003 |

**Need:** Find name for Employee ID "E002"

**Formula:** `=INDEX(A2:A4, MATCH("E002", C2:C4, 0))`

**Result:** `Bob`

**Explanation:**
1. MATCH finds "E002" at position 2 in C2:C4
2. INDEX returns the 2nd value from A2:A4

---

## Lookup with Multiple Criteria

### The Challenge
Find a value that matches TWO conditions (like product AND region).

### Quick Answer (XLOOKUP - Excel 365)
```excel
=XLOOKUP(1, (criteria1_range=value1)*(criteria2_range=value2), return_range)
```

### Quick Answer (INDEX/MATCH)
```excel
=INDEX(return_range, MATCH(1, (range1=crit1)*(range2=crit2), 0))
```

### Full Example

**Data (A1:D5):**
| A | B | C | D |
|---|---|---|---|
| Product | Region | Quarter | Sales |
| Widget | North | Q1 | 1000 |
| Widget | South | Q1 | 800 |
| Gadget | North | Q1 | 1200 |
| Widget | North | Q2 | 1100 |

**Need:** Sales for Widget in North region in Q1

**Formula (XLOOKUP):**
```excel
=XLOOKUP(1, (A2:A5="Widget")*(B2:B5="North")*(C2:C5="Q1"), D2:D5)
```

**Formula (INDEX/MATCH):**
```excel
=INDEX(D2:D5, MATCH(1, (A2:A5="Widget")*(B2:B5="North")*(C2:C5="Q1"), 0))
```

**Result:** `1000`

### Helper Column Alternative
Create a helper column that combines criteria:
```excel
Helper: =A2&B2&C2    // "WidgetNorthQ1"
Lookup: =VLOOKUP("WidgetNorthQ1", Helper:D, 2, FALSE)
```

---

## Return Multiple Columns

### The Challenge
Look up once but return several related values (like Name, Email, and Phone).

### Quick Answer (XLOOKUP - Excel 365)
```excel
=XLOOKUP(lookup_value, lookup_range, return_multiple_columns)
```

### Full Example

**Data (A1:D4):**
| A | B | C | D |
|---|---|---|---|
| ID | Name | Email | Phone |
| E001 | Alice | alice@co.com | 555-1234 |
| E002 | Bob | bob@co.com | 555-5678 |
| E003 | Carol | carol@co.com | 555-9012 |

**Formula:** `=XLOOKUP("E002", A2:A4, B2:D4)`

**Result (spills across 3 cells):**
| | | |
|---|---|---|
| Bob | bob@co.com | 555-5678 |

---

## Return All Matches

### The Challenge
Your lookup value appears multiple times and you need ALL matching rows.

### Quick Answer (FILTER - Excel 365)
```excel
=FILTER(data_range, criteria_range=value, "No matches")
```

### Full Example

**Data (A1:C6):**
| A | B | C |
|---|---|---|
| Customer | Product | Amount |
| ABC Inc | Widget | 500 |
| XYZ Corp | Gadget | 300 |
| ABC Inc | Gadget | 750 |
| XYZ Corp | Widget | 400 |
| ABC Inc | Widget | 600 |

**Need:** All purchases by "ABC Inc"

**Formula:** `=FILTER(A2:C6, A2:A6="ABC Inc", "No purchases found")`

**Result (spills down):**
| | | |
|---|---|---|
| ABC Inc | Widget | 500 |
| ABC Inc | Gadget | 750 |
| ABC Inc | Widget | 600 |

---

## Handle Lookup Errors

### The Challenge
VLOOKUP returns #N/A when the value isn't found - you want something cleaner.

### Quick Answer
```excel
=IFERROR(VLOOKUP(...), "Not found")      // Catches any error
=IFNA(VLOOKUP(...), "Not found")         // Catches only #N/A
```

### Full Example

**Formula without error handling:**
```excel
=VLOOKUP("P999", A2:C10, 2, FALSE)
```
**Result:** `#N/A`

**Formula with IFERROR:**
```excel
=IFERROR(VLOOKUP("P999", A2:C10, 2, FALSE), "Product not found")
```
**Result:** `Product not found`

**XLOOKUP has built-in handling:**
```excel
=XLOOKUP("P999", A2:A10, B2:B10, "Product not found")
```

---

## Two-Way Lookup

### The Challenge
Find a value at the intersection of a row and column (like a grade from a rubric).

### Quick Answer
```excel
=INDEX(data, MATCH(row_value, row_headers, 0), MATCH(col_value, col_headers, 0))
```

### Full Example

**Shipping Rate Table (A1:D4):**
| | Zone 1 | Zone 2 | Zone 3 |
|---|---|---|---|
| Small | 5.99 | 7.99 | 9.99 |
| Medium | 8.99 | 11.99 | 14.99 |
| Large | 12.99 | 16.99 | 21.99 |

**Need:** Rate for "Medium" package to "Zone 2"

**Formula:**
```excel
=INDEX(B2:D4, MATCH("Medium", A2:A4, 0), MATCH("Zone 2", B1:D1, 0))
```

**Result:** `11.99`

**Explanation:**
1. MATCH("Medium", A2:A4, 0) → 2 (row position)
2. MATCH("Zone 2", B1:D1, 0) → 2 (column position)
3. INDEX returns value at row 2, column 2 of B2:D4

---

## Approximate Match Lookup

### The Challenge
Find the rate for a value within a range (like tax brackets or shipping by weight).

### Quick Answer
```excel
=VLOOKUP(value, table, column, TRUE)    // TRUE = approximate
=XLOOKUP(value, range, return, , -1)    // -1 = next smaller
```

### Full Example - Tax Brackets

**Tax Table (A1:B5):**
| A | B |
|---|---|
| Income | Tax Rate |
| 0 | 10% |
| 10000 | 12% |
| 40000 | 22% |
| 85000 | 24% |

**Income to look up:** $50,000

**Formula:** `=VLOOKUP(50000, A2:B5, 2, TRUE)`

**Result:** `22%`

**Explanation:** Finds largest value ≤ 50000 (which is 40000) and returns its rate.

**Important:** Data MUST be sorted ascending for approximate match!

---

## Case-Sensitive Lookup

### The Challenge
VLOOKUP treats "ABC" and "abc" as the same - you need case-sensitive matching.

### Quick Answer
```excel
=INDEX(return_range, MATCH(TRUE, EXACT(lookup_value, lookup_range), 0))
```

### Full Example

**Data:**
| A | B |
|---|---|
| Code | Value |
| abc | 100 |
| ABC | 200 |
| Abc | 300 |

**Formula:** `=INDEX(B2:B4, MATCH(TRUE, EXACT("ABC", A2:A4), 0))`

**Result:** `200` (matches "ABC" exactly, not "abc" or "Abc")

---

## Related Solutions

- [Conditional Calculations](../conditional-calculations/README.md) - SUMIF, COUNTIF, AVERAGEIF
- [Error Handling](../error-handling/README.md) - Handle all lookup errors
- [Data Analysis](../data-analysis/README.md) - Analyze data after retrieval

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
