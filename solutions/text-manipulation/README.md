# Text Manipulation Solutions

> **Extract, combine, clean, and transform text data**

## Quick Solutions

| I want to... | Solution |
|--------------|----------|
| Split first/last name | [Split Names](#split-names) |
| Combine cells with separator | [Concatenate Text](#combine-text) |
| Extract part of text | [Extract Substrings](#extract-text) |
| Clean up messy data | [Clean Text](#clean-text) |
| Change case | [Change Case](#change-text-case) |
| Find and replace | [Find/Replace in Formulas](#find-and-replace) |
| Extract numbers from text | [Extract Numbers](#extract-numbers) |
| Extract email/domain | [Extract Email Parts](#extract-email-parts) |
| Pad numbers with zeros | [Leading Zeros](#add-leading-zeros) |
| Format phone numbers | [Format Phones](#format-phone-numbers) |

---

## Split Names

### First Name from Full Name
```excel
=LEFT(A1, FIND(" ", A1)-1)
```

### Last Name from Full Name
```excel
=RIGHT(A1, LEN(A1)-FIND(" ", A1))
```

### Full Example - Split "John Smith"

| A | B (First) | C (Last) |
|---|-----------|----------|
| John Smith | John | Smith |

**First Name Formula:** `=LEFT(A1, FIND(" ",A1)-1)`
**Last Name Formula:** `=RIGHT(A1, LEN(A1)-FIND(" ",A1))`

### Handle Middle Names - Get Last Word
```excel
=TRIM(RIGHT(SUBSTITUTE(A1," ",REPT(" ",100)),100))
```

### Three-Part Split (First, Middle, Last)
```excel
First:  =LEFT(A1, FIND(" ",A1)-1)
Middle: =MID(A1, FIND(" ",A1)+1, FIND(" ",A1,FIND(" ",A1)+1)-FIND(" ",A1)-1)
Last:   =TRIM(RIGHT(SUBSTITUTE(A1," ",REPT(" ",100)),100))
```

### Using Flash Fill (Fastest Method)
1. Type desired result in first cell (e.g., "John")
2. Press Ctrl+E
3. Excel fills the rest automatically

### Excel 365: TEXTSPLIT
```excel
=TEXTSPLIT(A1, " ")
```
Returns array: `{"John", "Smith"}` (spills to columns)

---

## Combine Text

### Basic Concatenation
```excel
=A1 & " " & B1                          → "John Smith"
=CONCAT(A1, " ", B1)                    → "John Smith"
=A1 & ", " & B1                         → "Smith, John"
```

### TEXTJOIN with Delimiter (Excel 2019+)
```excel
=TEXTJOIN(", ", TRUE, A1:D1)
```
- First argument: delimiter
- Second argument: TRUE = ignore blanks
- Third argument: range to join

### Full Example - Build Address

| A | B | C | D |
|---|---|---|---|
| 123 Main St | Suite 100 | New York | NY |

**Formula:** `=TEXTJOIN(", ", TRUE, A1:D1)`
**Result:** `123 Main St, Suite 100, New York, NY`

### Combine with Formatting
```excel
="Name: " & A1 & " | Total: " & TEXT(B1, "$#,##0.00")
```
**Result:** `Name: John | Total: $1,234.56`

### Multi-Line Text in Cell
```excel
=A1 & CHAR(10) & B1 & CHAR(10) & C1
```
(Enable Wrap Text for cell)

---

## Extract Text

### LEFT - First N Characters
```excel
=LEFT(A1, 3)        → First 3 characters
=LEFT(A1, 5)        → First 5 characters
```

### RIGHT - Last N Characters
```excel
=RIGHT(A1, 4)       → Last 4 characters
=RIGHT(A1, 2)       → Last 2 characters
```

### MID - Characters from Middle
```excel
=MID(A1, start, length)
=MID("ABCDEFG", 3, 2)   → "CD"
```

### Extract Before/After Delimiter

**Before first space:**
```excel
=LEFT(A1, FIND(" ", A1)-1)
```

**After first space:**
```excel
=MID(A1, FIND(" ", A1)+1, 100)
```

**Between delimiters:**
```excel
=MID(A1, FIND("-",A1)+1, FIND("-",A1,FIND("-",A1)+1)-FIND("-",A1)-1)
```

### Extract Nth Word
```excel
=TRIM(MID(SUBSTITUTE(A1," ",REPT(" ",100)), (N-1)*100+1, 100))
```
Where N = word number (1, 2, 3...)

---

## Clean Text

### Remove Extra Spaces
```excel
=TRIM(A1)           → Removes leading, trailing, and extra internal spaces
```

### Remove Non-Printable Characters
```excel
=CLEAN(A1)          → Removes characters 0-31 (non-printable)
```

### Full Clean
```excel
=TRIM(CLEAN(A1))    → Both in one
```

### Remove Specific Characters
```excel
=SUBSTITUTE(A1, "-", "")        → Remove dashes
=SUBSTITUTE(A1, " ", "")        → Remove all spaces
=SUBSTITUTE(A1, ",", "")        → Remove commas
```

### Remove Multiple Characters
```excel
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A1,"-",""),"(",""),")","")
```

### Remove Line Breaks
```excel
=SUBSTITUTE(SUBSTITUTE(A1, CHAR(10), " "), CHAR(13), " ")
```

### Clean Phone Numbers (Keep Only Digits)
```excel
=TEXTJOIN("",TRUE,IF(ISNUMBER(MID(A1,ROW(INDIRECT("1:"&LEN(A1))),1)+0),MID(A1,ROW(INDIRECT("1:"&LEN(A1))),1),""))
```
Or simpler with multiple SUBSTITUTE for common characters.

---

## Change Text Case

### UPPER - All Capitals
```excel
=UPPER(A1)          → "JOHN SMITH"
```

### LOWER - All Lowercase
```excel
=LOWER(A1)          → "john smith"
```

### PROPER - Title Case
```excel
=PROPER(A1)         → "John Smith"
```

### Sentence Case (First letter capital)
```excel
=UPPER(LEFT(A1,1)) & LOWER(MID(A1,2,LEN(A1)))
```

### Handle Exceptions (Mc, Mac, O')
```excel
=IF(LEFT(A1,2)="Mc",
    "Mc"&UPPER(MID(A1,3,1))&LOWER(MID(A1,4,100)),
    PROPER(A1))
```

---

## Find and Replace

### SUBSTITUTE - Replace Text
```excel
=SUBSTITUTE(A1, "old", "new")
=SUBSTITUTE(A1, "Mr.", "Mr")           → Remove period
=SUBSTITUTE(A1, "-", "/")               → Dash to slash
```

### Replace Nth Occurrence
```excel
=SUBSTITUTE(A1, "old", "new", 2)        → Replace only 2nd occurrence
```

### REPLACE - By Position
```excel
=REPLACE(A1, start, num_chars, new_text)
=REPLACE("ABCDEF", 3, 2, "XYZ")         → "ABXYZEF"
```

### Case-Sensitive Search
```excel
=FIND("Text", A1)       → Case-sensitive, returns position or #VALUE!
=SEARCH("text", A1)     → Case-insensitive
```

### Check if Text Contains
```excel
=ISNUMBER(SEARCH("find", A1))           → TRUE if found
=IF(ISNUMBER(SEARCH("LLC", A1)), "Corporation", "Individual")
```

---

## Extract Numbers

### Extract Numbers Only (Excel 365)
```excel
=TEXTJOIN("",TRUE,IF(ISNUMBER(--MID(A1,SEQUENCE(LEN(A1)),1)),MID(A1,SEQUENCE(LEN(A1)),1),""))
```

### Extract First Number from Text
```excel
=LOOKUP(9.99E+307,--("0"&MID(A1,MIN(SEARCH({0,1,2,3,4,5,6,7,8,9},A1&"0123456789")),ROW(INDIRECT("1:"&LEN(A1))))))
```

### Simpler: Extract Known Pattern
If format is consistent like "ABC-123":
```excel
=VALUE(RIGHT(A1, 3))        → 123
=--MID(A1, 5, 3)            → 123
```

### Extract Currency Amount
```excel
=VALUE(SUBSTITUTE(SUBSTITUTE(A1,"$",""),",",""))
```
Converts "$1,234.56" to 1234.56

---

## Extract Email Parts

### Full Email: john.smith@company.com

### Username (Before @)
```excel
=LEFT(A1, FIND("@", A1)-1)
```
**Result:** `john.smith`

### Domain (After @)
```excel
=MID(A1, FIND("@", A1)+1, 100)
```
**Result:** `company.com`

### Domain Name Only (Without .com)
```excel
=MID(A1, FIND("@",A1)+1, FIND(".",A1,FIND("@",A1))-FIND("@",A1)-1)
```
**Result:** `company`

### Extension Only (.com, .org, etc.)
```excel
=RIGHT(A1, LEN(A1)-FIND(".",A1,FIND("@",A1)))
```
**Result:** `com`

---

## Add Leading Zeros

### Fixed Width with TEXT
```excel
=TEXT(A1, "00000")          → "00123" (5 digits)
=TEXT(A1, "0000000000")     → "0000000123" (10 digits)
```

### Using REPT and RIGHT
```excel
=RIGHT("0000" & A1, 5)      → Last 5 chars, padded with zeros
```

### Preserve as Text (for IDs)
```excel
=TEXT(A1, "@")              → Keeps leading zeros
```

### Remove Leading Zeros
```excel
=VALUE(A1)                  → Converts to number
=A1*1                       → Quick conversion
=A1+0                       → Quick conversion
```

---

## Format Phone Numbers

### Convert 1234567890 to (123) 456-7890
```excel
="(" & LEFT(A1,3) & ") " & MID(A1,4,3) & "-" & RIGHT(A1,4)
```

### With TEXT Function
```excel
=TEXT(A1, "(###) ###-####")
```
Note: May not work if stored as number > 9999999999

### Strip Formatting and Reformat
```excel
Step 1 - Remove all formatting:
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A1,"-",""),"(",""),")","")," ","")

Step 2 - Reformat clean number
```

### Handle International (+1)
```excel
=IF(LEN(A1)=11, "+"&LEFT(A1,1)&" ("&MID(A1,2,3)&") "&MID(A1,5,3)&"-"&RIGHT(A1,4),
   "("&LEFT(A1,3)&") "&MID(A1,4,3)&"-"&RIGHT(A1,4))
```

---

## Advanced Text Operations

### Reverse Text
```excel
=TEXTJOIN("",1,MID(A1,LEN(A1)-ROW(INDIRECT("1:"&LEN(A1)))+1,1))
```

### Remove Duplicate Words
```excel
(Requires helper columns or complex array formula - use Power Query for best results)
```

### Extract Initials
```excel
=LEFT(A1,1) & MID(A1, FIND(" ",A1)+1, 1)
```
"John Smith" → "JS"

### Count Specific Character
```excel
=LEN(A1)-LEN(SUBSTITUTE(A1,"a",""))     → Count of "a" in cell
```

### Count Words
```excel
=LEN(TRIM(A1))-LEN(SUBSTITUTE(TRIM(A1)," ",""))+1
```

---

## Text Function Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| LEFT | First N chars | `=LEFT("Hello",2)` → "He" |
| RIGHT | Last N chars | `=RIGHT("Hello",2)` → "lo" |
| MID | Middle chars | `=MID("Hello",2,3)` → "ell" |
| LEN | Length | `=LEN("Hello")` → 5 |
| FIND | Position (case-sens) | `=FIND("l","Hello")` → 3 |
| SEARCH | Position (case-insens) | `=SEARCH("L","Hello")` → 3 |
| SUBSTITUTE | Replace text | `=SUBSTITUTE("Hi","i","o")` → "Ho" |
| REPLACE | Replace by position | `=REPLACE("ABC",2,1,"X")` → "AXC" |
| TRIM | Remove extra spaces | `=TRIM("  Hi  ")` → "Hi" |
| CLEAN | Remove non-printable | `=CLEAN(A1)` |
| UPPER | Uppercase | `=UPPER("hi")` → "HI" |
| LOWER | Lowercase | `=LOWER("HI")` → "hi" |
| PROPER | Title case | `=PROPER("hi")` → "Hi" |
| TEXT | Format number as text | `=TEXT(123,"0000")` → "0123" |
| VALUE | Text to number | `=VALUE("123")` → 123 |
| CONCAT | Join text | `=CONCAT("A","B")` → "AB" |
| TEXTJOIN | Join with delimiter | `=TEXTJOIN(",",1,A1:A3)` |
| REPT | Repeat text | `=REPT("*",5)` → "*****" |
| EXACT | Case-sensitive compare | `=EXACT("Hi","hi")` → FALSE |

---

## Related Solutions

- [Data Cleaning](../data-cleaning/README.md) - Clean imported data
- [Lookups](../lookups/README.md) - Find data based on text
- [Error Handling](../error-handling/README.md) - Handle text formula errors

---

[🏠 Back to Home](../../README.md) | [🎯 All Solutions](../README.md)
