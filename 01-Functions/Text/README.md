# 📝 Text Functions

> **30+ functions for text manipulation, formatting, and data cleaning**

## 📋 Table of Contents

- [Text Extraction](#text-extraction)
- [Text Combination](#text-combination)
- [Text Transformation](#text-transformation)
- [Text Search & Replace](#text-search--replace)
- [Text Conversion](#text-conversion)
- [Text Information](#text-information)
- [Advanced Text Functions](#advanced-text-functions)

---

## Text Extraction

### LEFT
**Extracts characters from the left side of text**

**Syntax:** `=LEFT(text, [num_chars])`

**Parameters:**
- `text`: Text string
- `num_chars`: (Optional) Number of characters to extract (default: 1)

**Examples:**
```excel
=LEFT("Excel", 3)                       → "Exc"
=LEFT(A1, 5)                            → First 5 characters
=LEFT("Product-123", 7)                 → "Product"
=LEFT(A1)                               → First character
```

**Real-World Uses:**
- Extract first name from full name
- Get area code from phone number
- Extract product code prefix
- Get first word

**Practical Example - Extract First Name:**
```excel
=LEFT(A1, FIND(" ", A1)-1)              // "John Smith" → "John"
```

---

### RIGHT
**Extracts characters from the right side of text**

**Syntax:** `=RIGHT(text, [num_chars])`

**Examples:**
```excel
=RIGHT("Excel", 3)                      → "cel"
=RIGHT(A1, 4)                           → Last 4 characters
=RIGHT("Invoice-2024", 4)               → "2024"
```

**Real-World Uses:**
- Extract file extension
- Get last 4 digits of credit card
- Extract year from date string
- Get last name

**Practical Example - Extract Extension:**
```excel
=RIGHT(A1, LEN(A1)-FIND(".", A1))       // "report.xlsx" → "xlsx"
```

---

### MID
**Extracts characters from the middle of text**

**Syntax:** `=MID(text, start_num, num_chars)`

**Parameters:**
- `text`: Text string
- `start_num`: Starting position (1 = first character)
- `num_chars`: Number of characters to extract

**Examples:**
```excel
=MID("Excel 2024", 7, 4)                → "2024"
=MID(A1, 5, 3)                          → 3 chars starting at position 5
=MID("ABC-DEF-GHI", 5, 3)               → "DEF"
```

**Real-World Uses:**
- Extract middle name
- Get month from date string
- Extract code from middle of ID
- Parse structured data

**Practical Example - Extract Area Code:**
```excel
=MID("(555) 123-4567", 2, 3)            → "555"
```

---

### TEXTSPLIT
**Splits text into multiple columns/rows (Excel 365)**

**Syntax:** `=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])`

**Examples:**
```excel
=TEXTSPLIT("John,Smith,35", ",")        → Splits into 3 columns
=TEXTSPLIT(A1, " ")                     → Split by space
=TEXTSPLIT("A|B|C", "|")                → Split by pipe
```

**Real-World Uses:**
- Split names
- Parse CSV data
- Separate addresses
- Break up concatenated values

---

## Text Combination

### CONCAT
**Combines text from multiple cells (Excel 2019+)**

**Syntax:** `=CONCAT(text1, [text2], ...)`

**Examples:**
```excel
=CONCAT(A1, " ", B1)                    → Combine with space
=CONCAT("Total: ", C1)                  → Add label
=CONCAT(A1:A5)                          → Join range
```

**CONCAT vs CONCATENATE:**
- CONCAT: Modern, accepts ranges
- CONCATENATE: Legacy, individual arguments only

---

### TEXTJOIN
**Combines text with a delimiter (Excel 2019+)**

**Syntax:** `=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`

**Parameters:**
- `delimiter`: Character(s) to insert between items
- `ignore_empty`: TRUE to skip empty cells
- `text1, text2...`: Text to combine

**Examples:**
```excel
=TEXTJOIN(", ", TRUE, A1:A5)            → "Apple, Orange, Banana"
=TEXTJOIN("-", TRUE, A1, B1, C1)        → "123-ABC-XYZ"
=TEXTJOIN(" ", TRUE, A1:A10)            → Join with spaces
=TEXTJOIN(CHAR(10), TRUE, A1:A5)        → Join with line breaks
```

**Real-World Uses:**
- Create comma-separated lists
- Build full addresses
- Combine names
- Create email lists

**Advanced Example - Create Full Address:**
```excel
=TEXTJOIN(", ", TRUE, Street, City, State, ZIP)
```

---

### CONCATENATE
**Combines text (legacy function)**

**Syntax:** `=CONCATENATE(text1, [text2], ...)`

**Examples:**
```excel
=CONCATENATE(A1, " ", B1)               → "John Smith"
=CONCATENATE("Total: $", C1)            → "Total: $100"
```

**Note:** Use CONCAT or & operator instead in modern Excel

**Alternative - & Operator:**
```excel
=A1 & " " & B1                          → Same as CONCATENATE
="Total: $" & C1                        → Simpler syntax
```

---

## Text Transformation

### UPPER
**Converts text to uppercase**

**Syntax:** `=UPPER(text)`

**Examples:**
```excel
=UPPER("excel")                         → "EXCEL"
=UPPER(A1)                              → Convert cell to uppercase
=UPPER("john smith")                    → "JOHN SMITH"
```

**Real-World Uses:**
- Standardize data entry
- Create acronyms
- Format headers
- Database matching (case-insensitive)

---

### LOWER
**Converts text to lowercase**

**Syntax:** `=LOWER(text)`

**Examples:**
```excel
=LOWER("EXCEL")                         → "excel"
=LOWER(A1)                              → Convert to lowercase
=LOWER("John Smith")                    → "john smith"
```

**Real-World Uses:**
- Email addresses
- URLs
- Standardize input
- Create usernames

---

### PROPER
**Converts text to proper case (Title Case)**

**Syntax:** `=PROPER(text)`

**Examples:**
```excel
=PROPER("john smith")                   → "John Smith"
=PROPER("EXCEL MASTERY")                → "Excel Mastery"
=PROPER(A1)                             → Convert to title case
```

**Real-World Uses:**
- Format names correctly
- Clean data entry
- Create titles
- Standardize addresses

**Limitation:**
```excel
=PROPER("mcdonald")                     → "Mcdonald" (not "McDonald")
```

---

### TRIM
**Removes extra spaces from text**

**Syntax:** `=TRIM(text)`

**Examples:**
```excel
=TRIM("  Excel  ")                      → "Excel"
=TRIM("Multiple   spaces")              → "Multiple spaces" (single space)
=TRIM(A1)                               → Clean up cell
```

**What it does:**
- Removes leading spaces
- Removes trailing spaces
- Reduces multiple spaces to single space
- Does NOT remove line breaks

**Real-World Uses:**
- Clean imported data
- Fix data entry errors
- Prepare for lookups
- Standardize text

**Best Practice:**
```excel
=TRIM(UPPER(A1))                        // Clean and standardize
```

---

### CLEAN
**Removes non-printable characters**

**Syntax:** `=CLEAN(text)`

**Examples:**
```excel
=CLEAN(A1)                              → Remove hidden characters
```

**Real-World Uses:**
- Clean data from web/databases
- Remove line breaks (CHAR(10))
- Fix imported data
- Prepare for export

**TRIM vs CLEAN:**
- TRIM: Removes extra spaces
- CLEAN: Removes non-printable characters
- Often use together: `=TRIM(CLEAN(A1))`

---

### SUBSTITUTE
**Replaces specific text with new text**

**Syntax:** `=SUBSTITUTE(text, old_text, new_text, [instance_num])`

**Parameters:**
- `text`: Original text
- `old_text`: Text to replace
- `new_text`: Replacement text
- `instance_num`: (Optional) Which occurrence to replace

**Examples:**
```excel
=SUBSTITUTE("Excel Excel", "Excel", "Word")              → "Word Word"
=SUBSTITUTE("Excel Excel", "Excel", "Word", 1)           → "Word Excel" (first only)
=SUBSTITUTE(A1, " ", "")                                 → Remove all spaces
=SUBSTITUTE(A1, "-", "/")                                → Replace hyphens with slashes
=SUBSTITUTE(A1, CHAR(10), ", ")                          → Replace line breaks
```

**Real-World Uses:**
- Fix formatting
- Replace abbreviations
- Clean phone numbers
- Convert date formats

**Case-Sensitive:**
```excel
=SUBSTITUTE("Excel excel", "excel", "WORD")              → "Excel WORD" (case matters)
```

---

### REPLACE
**Replaces text at a specific position**

**Syntax:** `=REPLACE(old_text, start_num, num_chars, new_text)`

**Examples:**
```excel
=REPLACE("Excel 2019", 7, 4, "2024")    → "Excel 2024"
=REPLACE(A1, 1, 3, "***")               → Replace first 3 chars with ***
```

**REPLACE vs SUBSTITUTE:**
- REPLACE: Based on position
- SUBSTITUTE: Based on content

---

## Text Search & Replace

### FIND
**Finds position of text (case-sensitive)**

**Syntax:** `=FIND(find_text, within_text, [start_num])`

**Examples:**
```excel
=FIND("x", "Excel")                     → 2
=FIND(" ", "John Smith")                → 5 (position of first space)
=FIND("@", "user@email.com")            → 5
=FIND(".", "file.txt")                  → 5
```

**Returns:** Position number (1-based) or #VALUE! if not found

**Real-World Uses:**
- Find delimiter positions
- Parse email addresses
- Split text at specific character
- Validate format

**With Other Functions:**
```excel
=LEFT(A1, FIND("@", A1)-1)              // Extract username from email
=MID(A1, FIND("@", A1)+1, LEN(A1))      // Extract domain from email
```

---

### SEARCH
**Finds position of text (case-insensitive)**

**Syntax:** `=SEARCH(find_text, within_text, [start_num])`

**Examples:**
```excel
=SEARCH("excel", "Microsoft Excel")     → 11 (case-insensitive)
=SEARCH("x", "Excel")                   → 2
```

**Supports Wildcards:**
```excel
=SEARCH("E*l", "Excel")                 → 1 (* = any characters)
=SEARCH("E?cel", "Excel")               → 1 (? = single character)
```

**FIND vs SEARCH:**
| Feature | FIND | SEARCH |
|---------|------|--------|
| Case-sensitive | Yes | No |
| Wildcards | No | Yes |
| Speed | Faster | Slower |

---

### LEN
**Returns the length of text**

**Syntax:** `=LEN(text)`

**Examples:**
```excel
=LEN("Excel")                           → 5
=LEN(A1)                                → Count characters
=LEN("   ")                             → 3 (includes spaces)
```

**Real-World Uses:**
- Validate input length
- Check password requirements
- Count characters in tweets
- Data validation

**Practical Examples:**
```excel
=IF(LEN(A1)>50, "Too Long", "OK")       // Validate length
=IF(LEN(A1)=0, "Empty", "Has Value")    // Check if empty
=LEN(A1)-LEN(SUBSTITUTE(A1," ",""))+1   // Count words
```

---

## Text Conversion

### TEXT
**Converts number to text with formatting**

**Syntax:** `=TEXT(value, format_text)`

**Examples:**
```excel
=TEXT(1234.5, "$#,##0.00")              → "$1,234.50"
=TEXT(0.15, "0%")                       → "15%"
=TEXT(TODAY(), "MM/DD/YYYY")            → "12/14/2025"
=TEXT(TODAY(), "MMMM DD, YYYY")         → "December 14, 2025"
=TEXT(A1, "0000")                       → "0042" (pad with zeros)
```

**Common Format Codes:**

**Numbers:**
```excel
=TEXT(1234, "0")                        → "1234"
=TEXT(1234, "0.00")                     → "1234.00"
=TEXT(1234, "#,##0")                    → "1,234"
=TEXT(1234, "$#,##0.00")                → "$1,234.00"
```

**Dates:**
```excel
=TEXT(date, "MM/DD/YYYY")               → "12/14/2025"
=TEXT(date, "DD-MMM-YYYY")              → "14-Dec-2025"
=TEXT(date, "MMMM D, YYYY")             → "December 14, 2025"
=TEXT(date, "DDD")                      → "Sat"
=TEXT(date, "DDDD")                     → "Saturday"
```

**Times:**
```excel
=TEXT(time, "HH:MM:SS")                 → "14:30:00"
=TEXT(time, "HH:MM AM/PM")              → "02:30 PM"
```

**Real-World Uses:**
- Format invoice numbers
- Create custom date displays
- Combine numbers with text
- Export formatting

**Important Note:**
Result is TEXT, not a number. Can't use in calculations.

---

### VALUE
**Converts text to number**

**Syntax:** `=VALUE(text)`

**Examples:**
```excel
=VALUE("123")                           → 123
=VALUE("$1,234.50")                     → 1234.5
=VALUE("15%")                           → 0.15
=VALUE("12/14/2025")                    → 45639 (date serial)
```

**Real-World Uses:**
- Convert imported text numbers
- Parse formatted strings
- Fix "numbers stored as text"
- Data cleaning

**Error Handling:**
```excel
=IFERROR(VALUE(A1), 0)                  // Return 0 if can't convert
```

---

### NUMBERVALUE
**Converts text to number with custom decimal/grouping**

**Syntax:** `=NUMBERVALUE(text, [decimal_separator], [group_separator])`

**Examples:**
```excel
=NUMBERVALUE("1.234,56", ",", ".")      → 1234.56 (European format)
=NUMBERVALUE("1 234,56", ",", " ")      → 1234.56
```

**Use:** International number formats

---

### DOLLAR
**Converts number to text in currency format**

**Syntax:** `=DOLLAR(number, [decimals])`

**Examples:**
```excel
=DOLLAR(1234.567, 2)                    → "$1,234.57"
=DOLLAR(1234.567)                       → "$1,234.57"
=DOLLAR(1234.567, 0)                    → "$1,235"
```

**Note:** Result is TEXT. Use TEXT() for more flexibility.

---

### CHAR
**Returns character from number (ASCII/Unicode)**

**Syntax:** `=CHAR(number)`

**Examples:**
```excel
=CHAR(65)                               → "A"
=CHAR(10)                               → Line break
=CHAR(13)                               → Carriage return
=CHAR(9)                                → Tab
```

**Common Uses:**
```excel
="Line 1" & CHAR(10) & "Line 2"         // Multi-line cell
=TEXTJOIN(CHAR(10), TRUE, A1:A5)        // Join with line breaks
```

---

### CODE
**Returns numeric code for first character**

**Syntax:** `=CODE(text)`

**Examples:**
```excel
=CODE("A")                              → 65
=CODE("Excel")                          → 69 (E)
=CODE("1")                              → 49
```

---

### UNICHAR & UNICODE
**Unicode character and code (Excel 2013+)**

**Syntax:**
```excel
=UNICHAR(number)                        → Returns character
=UNICODE(text)                          → Returns code
```

**Examples:**
```excel
=UNICHAR(9733)                          → "★" (star)
=UNICHAR(128512)                        → "😀" (emoji)
=UNICODE("★")                           → 9733
```

---

## Text Information

### EXACT
**Case-sensitive text comparison**

**Syntax:** `=EXACT(text1, text2)`

**Examples:**
```excel
=EXACT("Excel", "Excel")                → TRUE
=EXACT("Excel", "excel")                → FALSE
=EXACT(A1, B1)                          → Compare cells
```

**Real-World Uses:**
- Case-sensitive validation
- Password matching
- Quality control
- Data verification

**Note:** Regular = comparison is case-insensitive

---

### ISTEXT
**Checks if value is text**

**Syntax:** `=ISTEXT(value)`

**Examples:**
```excel
=ISTEXT("Excel")                        → TRUE
=ISTEXT(123)                            → FALSE
=ISTEXT(A1)                             → Check cell type
```

---

## Advanced Text Functions

### TEXTBEFORE & TEXTAFTER
**Extract text before/after delimiter (Excel 365)**

**Syntax:**
```excel
=TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
=TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
```

**Examples:**
```excel
=TEXTBEFORE("John.Smith@email.com", "@")     → "John.Smith"
=TEXTAFTER("John.Smith@email.com", "@")      → "email.com"
=TEXTBEFORE(A1, " ", 2)                      → Text before 2nd space
```

---

### REPT
**Repeats text a specified number of times**

**Syntax:** `=REPT(text, number_times)`

**Examples:**
```excel
=REPT("*", 5)                           → "*****"
=REPT("-", 10)                          → "----------"
=REPT(A1, 3)                            → Repeat cell value 3 times
```

**Real-World Uses:**
- Create visual bars in cells
- Format separators
- Pad strings

**Visual Bar Chart:**
```excel
=REPT("█", A1/10)                       // Bar chart in cell
=REPT("▓", INT(B1*10))                  // Rating display
```

---

### T
**Returns text or empty string**

**Syntax:** `=T(value)`

**Examples:**
```excel
=T("Excel")                             → "Excel"
=T(123)                                 → "" (empty)
=T(TRUE)                                → "" (empty)
```

**Use:** Rarely needed in modern Excel

---

## Practical Examples & Patterns

### Extract Email Username
```excel
=LEFT(A1, FIND("@", A1)-1)
// "user@email.com" → "user"
```

### Extract Email Domain
```excel
=MID(A1, FIND("@", A1)+1, LEN(A1))
// "user@email.com" → "email.com"
```

### Extract First Name
```excel
=LEFT(A1, FIND(" ", A1)-1)
// "John Smith" → "John"
```

### Extract Last Name
```excel
=RIGHT(A1, LEN(A1)-FIND(" ", A1))
// "John Smith" → "Smith"
```

### Count Words
```excel
=LEN(TRIM(A1))-LEN(SUBSTITUTE(A1," ",""))+1
```

### Remove Non-Numeric Characters
```excel
=SUMPRODUCT(MID(0&A1,LARGE(INDEX(ISNUMBER(--MID(A1,ROW($1:$25),1))*ROW($1:$25),0),ROW($1:$25))+1,1)*10^ROW($1:$25)/10)
// Simpler in Excel 365 with TEXTJOIN + IF
```

### Create Initials
```excel
=LEFT(A1,1) & LEFT(MID(A1,FIND(" ",A1)+1,LEN(A1)),1)
// "John Smith" → "JS"
```

### Reverse Text
```excel
=TEXTJOIN("",TRUE,MID(A1,LEN(A1)-ROW(INDIRECT("1:"&LEN(A1)))+1,1))
```

### Title Case with Exceptions
```excel
=PROPER(LOWER(A1))
// Better than just PROPER for all caps
```

### Clean Phone Number
```excel
=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A1,"-",""),"(",""),")","")
// "(555) 123-4567" → "5551234567"
```

### Format Phone Number
```excel
="(" & LEFT(A1,3) & ") " & MID(A1,4,3) & "-" & RIGHT(A1,4)
// "5551234567" → "(555) 123-4567"
```

---

## Quick Reference

| Function | Purpose | Example |
|----------|---------|---------|
| LEFT | Extract from left | `=LEFT(A1,5)` |
| RIGHT | Extract from right | `=RIGHT(A1,3)` |
| MID | Extract from middle | `=MID(A1,5,3)` |
| TEXTJOIN | Join with delimiter | `=TEXTJOIN(", ",TRUE,A1:A5)` |
| CONCAT | Combine text | `=CONCAT(A1," ",B1)` |
| UPPER | To uppercase | `=UPPER(A1)` |
| LOWER | To lowercase | `=LOWER(A1)` |
| PROPER | To title case | `=PROPER(A1)` |
| TRIM | Remove extra spaces | `=TRIM(A1)` |
| SUBSTITUTE | Replace text | `=SUBSTITUTE(A1,"old","new")` |
| FIND | Find position (case) | `=FIND("@",A1)` |
| SEARCH | Find position (no case) | `=SEARCH("text",A1)` |
| LEN | Text length | `=LEN(A1)` |
| TEXT | Format as text | `=TEXT(A1,"$#,##0.00")` |
| VALUE | Convert to number | `=VALUE(A1)` |

---

## Best Practices

### Data Cleaning
```excel
=TRIM(CLEAN(PROPER(A1)))                // Ultimate clean
```

### Error Handling
```excel
=IFERROR(FIND("@",A1), 0)               // Return 0 if not found
```

### Performance
- Use CONCAT instead of CONCATENATE
- Use TEXTJOIN for multiple items
- Avoid volatile functions in large datasets

### Validation
```excel
=AND(LEN(A1)>=8, ISNUMBER(FIND("@",A1)))  // Email validation
```

---

**[⬆ Back to Main README](../../README.md)**
