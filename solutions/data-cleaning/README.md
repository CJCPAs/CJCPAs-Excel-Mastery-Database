# Data Cleaning & Transformation Solutions

> **Clean, standardize, and transform your messy data into usable information**

## Quick Solutions

| Problem | Solution |
|---------|----------|
| Extra spaces | [TRIM](#remove-extra-spaces) |
| Inconsistent capitalization | [UPPER/LOWER/PROPER](#fix-capitalization) |
| Split names into first/last | [Text to Columns or formulas](#split-names) |
| Combine columns | [CONCAT/TEXTJOIN](#combine-columns) |
| Numbers stored as text | [VALUE or Paste Special](#convert-text-to-numbers) |
| Remove duplicates | [UNIQUE or Remove Duplicates](#remove-duplicates) |
| Clean imported data | [TRIM + CLEAN](#clean-imported-data) |

---

## Remove Extra Spaces

### The Challenge
Data imported from other systems often has extra spaces - leading spaces, trailing spaces, or multiple spaces between words.

### Quick Answer
```excel
=TRIM(A2)
```

### Full Example

**Starting Data:**
| A |
|---|
| "  John   Smith  " |
| "Alice    Jones" |
| "Bob Brown" |

**Formula:** `=TRIM(A2)`

**Result:**
| Original | Cleaned |
|----------|---------|
| "  John   Smith  " | "John Smith" |
| "Alice    Jones" | "Alice Jones" |
| "Bob Brown" | "Bob Brown" |

### What TRIM Does
- Removes leading spaces
- Removes trailing spaces
- Reduces multiple spaces to single space
- Does NOT remove line breaks (use CLEAN for that)

### Apply to Entire Column
1. Add formula in helper column
2. Copy entire column
3. Paste Values over original column (Ctrl+Shift+V)
4. Delete helper column

---

## Fix Capitalization

### The Challenge
Data entry inconsistency: some names are ALL CAPS, some lowercase, some mixed.

### Quick Answer
```excel
=PROPER(A2)           // Title Case: "john smith" → "John Smith"
=UPPER(A2)            // ALL CAPS: "john smith" → "JOHN SMITH"
=LOWER(A2)            // all lower: "JOHN SMITH" → "john smith"
```

### Full Example

**Starting Data:**
| A |
|---|
| JOHN SMITH |
| alice jones |
| bOB BrOwN |

**Formula:** `=PROPER(A2)`

**Result:**
| Original | Proper Case |
|----------|-------------|
| JOHN SMITH | John Smith |
| alice jones | Alice Jones |
| bOB BrOwN | Bob Brown |

### Limitation - PROPER Doesn't Handle
- McDonald → Mcdonald (not McDonald)
- O'Brien → O'brien (not O'Brien)

**Fix:** Use SUBSTITUTE for special cases:
```excel
=SUBSTITUTE(PROPER(A2), "Mcdonald", "McDonald")
```

---

## Split Names

### The Challenge
Full names in one cell need to be separated into First Name and Last Name columns.

### Quick Answer - For "First Last" Format
```excel
First Name: =LEFT(A2, FIND(" ", A2)-1)
Last Name:  =MID(A2, FIND(" ", A2)+1, 100)
```

### Full Example

**Starting Data:**
| A |
|---|
| John Smith |
| Alice Jones |
| Bob Brown Jr |

**First Name Formula:** `=IFERROR(LEFT(A2, FIND(" ", A2)-1), A2)`

**Last Name Formula:** `=IFERROR(MID(A2, FIND(" ", A2)+1, 100), "")`

**Result:**
| Full Name | First | Last |
|-----------|-------|------|
| John Smith | John | Smith |
| Alice Jones | Alice | Jones |
| Bob Brown Jr | Bob | Brown Jr |

### Excel 365 - Use TEXTSPLIT
```excel
=TEXTSPLIT(A2, " ")
```
This spills first name and last name into separate columns automatically.

### Alternative: Text to Columns
1. Select the column
2. Data → Text to Columns
3. Choose "Delimited"
4. Check "Space"
5. Finish

---

## Combine Columns

### The Challenge
You have first name and last name in separate columns and need them combined.

### Quick Answer
```excel
=A2 & " " & B2                    // Simple concatenation
=TEXTJOIN(" ", TRUE, A2:C2)       // Join multiple with delimiter
```

### Full Example

**Starting Data:**
| A | B | C |
|---|---|---|
| First | Last | Title |
| John | Smith | Mr. |
| Alice | Jones | Ms. |

**Simple Join:** `=A2 & " " & B2`
**Result:** `John Smith`

**With Title:** `=C2 & " " & A2 & " " & B2`
**Result:** `Mr. John Smith`

**TEXTJOIN (ignores blanks):** `=TEXTJOIN(" ", TRUE, C2, A2, B2)`
**Result:** `Mr. John Smith` (works even if Title is blank)

### Create Email from Name
```excel
=LOWER(A2) & "." & LOWER(B2) & "@company.com"
```
**Result:** `john.smith@company.com`

---

## Convert Text to Numbers

### The Challenge
Numbers imported from other sources are stored as text (green triangle in corner), so formulas don't work correctly.

### Quick Answer
```excel
=VALUE(A2)              // Convert text to number
=A2*1                   // Multiply by 1 (quick trick)
=A2+0                   // Add 0 (same effect)
=--A2                   // Double negative (same effect)
```

### Method 1: VALUE Function
**Formula:** `=VALUE(A2)`

**Starting Data:** `"1234"` (text)
**Result:** `1234` (number)

### Method 2: Paste Special Multiply
1. Type `1` in an empty cell
2. Copy that cell
3. Select the "text numbers"
4. Paste Special → Multiply
5. Delete the 1

### Method 3: Text to Columns
1. Select the column
2. Data → Text to Columns
3. Click Finish immediately (no changes needed)

This forces Excel to re-interpret the values.

---

## Remove Duplicates

### The Challenge
You have a list with duplicate entries and need unique values only.

### Quick Answer - Formula (Excel 365)
```excel
=UNIQUE(A2:A100)
```

### Quick Answer - Feature (All Versions)
1. Select your data
2. Data → Remove Duplicates
3. Choose columns to check
4. OK

### Full Example

**Starting Data:**
| A |
|---|
| Apple |
| Banana |
| Apple |
| Cherry |
| Banana |

**Formula:** `=UNIQUE(A2:A6)`

**Result (spills down):**
| A |
|---|
| Apple |
| Banana |
| Cherry |

### Sort Unique Values
```excel
=SORT(UNIQUE(A2:A100))
```

### Count Unique Values
```excel
=ROWS(UNIQUE(A2:A100))
```
Or for older Excel:
```excel
=SUMPRODUCT(1/COUNTIF(A2:A100, A2:A100))
```

---

## Clean Imported Data

### The Challenge
Data from external sources (web, databases, CSV files) often contains hidden characters, weird spaces, and formatting issues.

### Quick Answer
```excel
=TRIM(CLEAN(A2))
```

### The Ultimate Data Cleaning Formula
```excel
=TRIM(CLEAN(SUBSTITUTE(A2, CHAR(160), " ")))
```

### What Each Function Does
- **CLEAN:** Removes non-printable characters (codes 0-31)
- **TRIM:** Removes extra spaces
- **SUBSTITUTE(A2, CHAR(160), " "):** Replaces non-breaking spaces with regular spaces

### Full Data Cleaning Workflow

**Step 1:** Check for hidden characters
```excel
=LEN(A2)    // If longer than visible text, there are hidden chars
```

**Step 2:** Apply cleaning formula
```excel
=TRIM(CLEAN(SUBSTITUTE(A2, CHAR(160), " ")))
```

**Step 3:** Convert to proper format
```excel
=PROPER(TRIM(CLEAN(A2)))
```

**Step 4:** Replace original data
1. Copy the cleaned column
2. Paste as Values over original
3. Delete helper column

---

## Handle Blank Cells

### The Challenge
Blank cells cause errors in calculations or lookups.

### Quick Answer
```excel
=IF(A2="", "No Data", A2)           // Replace blanks with text
=IF(ISBLANK(A2), 0, A2)             // Replace blanks with 0
```

### Count Blank Cells
```excel
=COUNTBLANK(A2:A100)
```

### Fill Blanks with Value Above
1. Select range with blanks
2. Go To Special (Ctrl+G → Special)
3. Select "Blanks"
4. Type `=` and the cell above (e.g., `=A2`)
5. Ctrl+Enter to fill all

---

## Remove Special Characters

### The Challenge
Phone numbers with parentheses, dashes, or spaces need to be cleaned.

### Quick Answer
```excel
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2, "(", ""), ")", ""), "-", "")
```

### Full Example - Clean Phone Numbers

**Starting Data:**
| A |
|---|
| (555) 123-4567 |
| 555.234.5678 |
| 555-345-6789 |

**Formula:**
```excel
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2, "(", ""), ")", ""), "-", ""), " ", "")
```

**Result:**
| Original | Cleaned |
|----------|---------|
| (555) 123-4567 | 5551234567 |
| 555.234.5678 | 555.234.5678 |
| 555-345-6789 | 5553456789 |

### Extract Only Numbers (Advanced)
For complex cases, use Power Query or VBA.

---

## Standardize Date Formats

### The Challenge
Dates in various formats (text) need to be converted to proper Excel dates.

### Quick Answer
```excel
=DATEVALUE(A2)          // For text dates like "12/14/2024"
```

### Handle Different Formats

**US Format (MM/DD/YYYY):**
```excel
=DATEVALUE(A2)
```

**European Format (DD/MM/YYYY):**
```excel
=DATE(RIGHT(A2,4), MID(A2,4,2), LEFT(A2,2))
```

**Text format like "December 14, 2024":**
```excel
=DATEVALUE(A2)
```

### Display Date Consistently
After converting to real dates, format all cells the same way:
1. Select all date cells
2. Ctrl+1 (Format Cells)
3. Choose Date
4. Pick your preferred format

---

## Related Solutions

- [Text Manipulation](../text-manipulation/README.md) - More text transformations
- [Error Handling](../error-handling/README.md) - Deal with errors during cleaning
- [Lookups](../lookups/README.md) - Match cleaned data to other tables

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
