# Formulas Tab - Complete Reference

> **Function library, defined names, formula auditing, and calculation options**

**Ribbon Access:** Press **Alt+M** to activate Formulas tab

---

## Tab Overview

The Formulas tab provides access to all Excel functions and formula tools:

| Group | Purpose |
|-------|---------|
| [Function Library](#function-library-group) | Insert functions by category |
| [Defined Names](#defined-names-group) | Named ranges and constants |
| [Formula Auditing](#formula-auditing-group) | Debug and trace formulas |
| [Calculation](#calculation-group) | Calculation mode settings |

---

## Function Library Group

### Insert Function
| Item | Description | Shortcut |
|------|-------------|----------|
| **Insert Function** | Opens function wizard | Shift+F3 |

#### Insert Function Dialog
| Feature | Description |
|---------|-------------|
| **Search for a function** | Type description to find function |
| **Select a category** | Filter by function type |
| **Select a function** | Choose from list |
| **Help on this function** | Opens help for selected function |

### AutoSum
| Item | Description | Shortcut |
|------|-------------|----------|
| **AutoSum** | Quick aggregate functions | Alt+= |

#### AutoSum Dropdown
| Function | Description |
|----------|-------------|
| **Sum** | =SUM(range) |
| **Average** | =AVERAGE(range) |
| **Count Numbers** | =COUNT(range) |
| **Max** | =MAX(range) |
| **Min** | =MIN(range) |
| **More Functions...** | Opens Insert Function dialog |

### Recently Used
| Item | Description | Shortcut |
|------|-------------|----------|
| **Recently Used** | Your recent functions | - |

### Financial
| Item | Description | Shortcut |
|------|-------------|----------|
| **Financial** | Financial functions | - |

#### Financial Functions Include
| Category | Functions |
|----------|-----------|
| **Loans** | PMT, PPMT, IPMT, NPER, RATE, PV, FV |
| **Investments** | NPV, XNPV, IRR, XIRR, MIRR |
| **Depreciation** | SLN, DB, DDB, SYD, VDB |
| **Bonds** | PRICE, YIELD, DURATION, MDURATION |
| **Treasury** | TBILLEQ, TBILLPRICE, TBILLYIELD |
| **Coupons** | COUPDAYBS, COUPDAYS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD |
| **Interest** | ACCRINT, ACCRINTM, INTRATE, RECEIVED |
| **Conversion** | DOLLARDE, DOLLARFR, NOMINAL, EFFECT |

### Logical
| Item | Description | Shortcut |
|------|-------------|----------|
| **Logical** | Decision functions | - |

#### Logical Functions Include
| Function | Purpose |
|----------|---------|
| **IF** | Conditional test |
| **IFS** | Multiple conditions (Excel 2019+) |
| **IFERROR** | Trap errors |
| **IFNA** | Trap #N/A only |
| **AND** | All conditions true |
| **OR** | Any condition true |
| **XOR** | Exclusive or |
| **NOT** | Reverse logic |
| **SWITCH** | Multiple value matching |
| **TRUE** | Logical true |
| **FALSE** | Logical false |
| **LET** | Define variables (Excel 365) |
| **LAMBDA** | Create custom functions (Excel 365) |

### Text
| Item | Description | Shortcut |
|------|-------------|----------|
| **Text** | Text manipulation | - |

#### Text Functions Include
| Category | Functions |
|----------|-----------|
| **Extract** | LEFT, RIGHT, MID |
| **Combine** | CONCAT, CONCATENATE, TEXTJOIN |
| **Transform** | UPPER, LOWER, PROPER, TRIM, CLEAN |
| **Find** | FIND, SEARCH, LEN |
| **Replace** | REPLACE, SUBSTITUTE |
| **Format** | TEXT, VALUE, FIXED, DOLLAR |
| **Compare** | EXACT |
| **Repeat** | REPT |
| **Character** | CHAR, CODE, UNICHAR, UNICODE |

### Date & Time
| Item | Description | Shortcut |
|------|-------------|----------|
| **Date & Time** | Date/time functions | - |

#### Date & Time Functions Include
| Category | Functions |
|----------|-----------|
| **Current** | TODAY, NOW |
| **Create** | DATE, TIME, DATEVALUE, TIMEVALUE |
| **Extract** | YEAR, MONTH, DAY, HOUR, MINUTE, SECOND |
| **Calculate** | DATEDIF, DAYS, NETWORKDAYS, WORKDAY |
| **Month End** | EOMONTH, EDATE |
| **Week** | WEEKDAY, WEEKNUM, ISOWEEKNUM |
| **Year** | YEARFRAC |

### Lookup & Reference
| Item | Description | Shortcut |
|------|-------------|----------|
| **Lookup & Reference** | Data retrieval | - |

#### Lookup & Reference Functions Include
| Category | Functions |
|----------|-----------|
| **Lookup** | VLOOKUP, HLOOKUP, XLOOKUP, LOOKUP |
| **Index/Match** | INDEX, MATCH, XMATCH |
| **Reference** | ROW, ROWS, COLUMN, COLUMNS |
| **Indirect** | INDIRECT, ADDRESS, OFFSET |
| **Hyperlink** | HYPERLINK |
| **Array (365)** | FILTER, SORT, SORTBY, UNIQUE, SEQUENCE |
| **Areas** | AREAS, CHOOSE |
| **Transpose** | TRANSPOSE |

### Math & Trig
| Item | Description | Shortcut |
|------|-------------|----------|
| **Math & Trig** | Mathematical functions | - |

#### Math & Trig Functions Include
| Category | Functions |
|----------|-----------|
| **Basic** | SUM, PRODUCT, ABS, SIGN |
| **Conditional** | SUMIF, SUMIFS, SUMPRODUCT |
| **Rounding** | ROUND, ROUNDUP, ROUNDDOWN, CEILING, FLOOR, TRUNC, INT |
| **Modulo** | MOD, QUOTIENT |
| **Power** | POWER, SQRT, EXP, LN, LOG, LOG10 |
| **Random** | RAND, RANDBETWEEN, RANDARRAY |
| **Trig** | SIN, COS, TAN, ASIN, ACOS, ATAN, RADIANS, DEGREES |
| **Matrix** | MMULT, MINVERSE, MDETERM |
| **Subtotal** | SUBTOTAL, AGGREGATE |
| **Sequences** | SEQUENCE (365), ROMAN, ARABIC |

### More Functions
| Item | Description | Shortcut |
|------|-------------|----------|
| **More Functions** | Additional categories | - |

#### More Functions Categories
| Category | Description |
|----------|-------------|
| **Statistical** | AVERAGE, MEDIAN, MODE, STDEV, VAR, CORREL, PERCENTILE, RANK, COUNTIF, COUNTIFS, AVERAGEIF, MAXIFS, MINIFS |
| **Engineering** | CONVERT, BIN2DEC, DEC2BIN, COMPLEX, DELTA, GESTEP |
| **Cube** | CUBEVALUE, CUBEMEMBER, CUBESET, CUBERANKEDMEMBER |
| **Information** | ISBLANK, ISERROR, ISNUMBER, ISTEXT, NA, TYPE, CELL, INFO |
| **Compatibility** | Legacy functions from older Excel |
| **Web** | WEBSERVICE, FILTERXML, ENCODEURL |

---

## Defined Names Group

### Name Manager
| Item | Description | Shortcut |
|------|-------------|----------|
| **Name Manager** | Manage all named ranges | Ctrl+F3 |

#### Name Manager Dialog
| Column | Description |
|--------|-------------|
| **Name** | Name of range/formula |
| **Value** | Current value |
| **Refers To** | Definition/formula |
| **Scope** | Workbook or specific sheet |
| **Comment** | Description |

#### Name Manager Buttons
| Button | Description |
|--------|-------------|
| **New...** | Create new name |
| **Edit...** | Modify selected name |
| **Delete** | Remove selected name |
| **Filter** | Show names by type |

#### Filter Options
| Filter | Description |
|--------|-------------|
| **Clear Filter** | Show all |
| **Names Scoped to Worksheet** | Sheet-level names |
| **Names Scoped to Workbook** | Global names |
| **Names with Errors** | Invalid definitions |
| **Names without Errors** | Valid definitions |
| **Defined Names** | Named ranges |
| **Table Names** | Table references |

### Define Name
| Item | Description | Shortcut |
|------|-------------|----------|
| **Define Name** | Create new named range | - |

#### Define Name Options
| Option | Description |
|--------|-------------|
| **Define Name...** | Opens New Name dialog |
| **Apply Names...** | Replace references with names |

#### New Name Dialog
| Field | Description |
|-------|-------------|
| **Name** | Name (no spaces) |
| **Scope** | Workbook or Sheet |
| **Comment** | Description |
| **Refers to** | Cell reference or formula |

### Use in Formula
| Item | Description | Shortcut |
|------|-------------|----------|
| **Use in Formula** | Insert existing names | - |

#### Use in Formula Options
| Option | Description |
|--------|-------------|
| **Paste Names...** | Opens Paste Name dialog |
| **[Named Ranges]** | List of all names |

### Create from Selection
| Item | Description | Shortcut |
|------|-------------|----------|
| **Create from Selection** | Auto-create names from data | Ctrl+Shift+F3 |

#### Create Names From
| Option | Description |
|--------|-------------|
| **Top row** | Use first row as names |
| **Left column** | Use first column as names |
| **Bottom row** | Use last row as names |
| **Right column** | Use last column as names |

---

## Formula Auditing Group

### Trace Precedents
| Item | Description | Shortcut |
|------|-------------|----------|
| **Trace Precedents** | Show cells that feed into active cell | - |

#### Precedent Arrows
| Arrow Type | Description |
|------------|-------------|
| **Blue** | Within same sheet |
| **Black dashed** | From another sheet |
| **Red** | Error cell |

### Trace Dependents
| Item | Description | Shortcut |
|------|-------------|----------|
| **Trace Dependents** | Show cells that use active cell | - |

### Remove Arrows
| Item | Description | Shortcut |
|------|-------------|----------|
| **Remove Arrows** | Clear tracing arrows | - |

#### Remove Arrows Options
| Option | Description |
|--------|-------------|
| **Remove Arrows** | All arrows |
| **Remove Precedent Arrows** | Input arrows only |
| **Remove Dependent Arrows** | Output arrows only |

### Show Formulas
| Item | Description | Shortcut |
|------|-------------|----------|
| **Show Formulas** | Toggle formula view | Ctrl+` |

### Error Checking
| Item | Description | Shortcut |
|------|-------------|----------|
| **Error Checking** | Find and fix errors | - |

#### Error Checking Options
| Option | Description |
|--------|-------------|
| **Error Checking...** | Check entire sheet |
| **Trace Error** | Show precedents of error |
| **Circular References** | Find circular refs |

#### Error Types Checked
| Error | Description |
|-------|-------------|
| **Cells containing formulas that result in an error** | #VALUE!, #REF!, etc. |
| **Inconsistent calculated column formula in tables** | Broken table formulas |
| **Cells containing years represented as 2 digits** | Y2K issues |
| **Numbers formatted as text or preceded by an apostrophe** | Text-numbers |
| **Formulas inconsistent with other formulas in the region** | Unusual patterns |
| **Formulas which omit cells in a region** | Possible range errors |
| **Unlocked cells containing formulas** | Security check |
| **Formulas referring to empty cells** | Blank references |
| **Data entered in a table is invalid** | Validation errors |

### Evaluate Formula
| Item | Description | Shortcut |
|------|-------------|----------|
| **Evaluate Formula** | Step through calculation | - |

#### Evaluate Dialog
| Button | Description |
|--------|-------------|
| **Evaluate** | Calculate next step |
| **Step In** | Enter nested function |
| **Step Out** | Exit nested function |
| **Restart** | Start over |

### Watch Window
| Item | Description | Shortcut |
|------|-------------|----------|
| **Watch Window** | Monitor cell values | - |

#### Watch Window Columns
| Column | Description |
|--------|-------------|
| **Book** | Workbook name |
| **Sheet** | Sheet name |
| **Name** | Named range (if any) |
| **Cell** | Cell address |
| **Value** | Current value |
| **Formula** | Cell formula |

---

## Calculation Group

### Calculation Options
| Item | Description | Shortcut |
|------|-------------|----------|
| **Calculation Options** | When to calculate | - |

#### Calculation Modes
| Mode | Description |
|------|-------------|
| **Automatic** | Calculate after each change |
| **Automatic Except for Data Tables** | Skip What-If tables |
| **Manual** | Only when requested |

### Calculate Now
| Item | Description | Shortcut |
|------|-------------|----------|
| **Calculate Now** | Recalculate all workbooks | F9 |

### Calculate Sheet
| Item | Description | Shortcut |
|------|-------------|----------|
| **Calculate Sheet** | Recalculate current sheet | Shift+F9 |

---

## Complete Alt Shortcuts Reference

### Function Library Group (Alt+M)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Insert Function | Alt+M, F | Shift+F3 |
| AutoSum | Alt+M, U, S | Alt+= |
| AutoSum Average | Alt+M, U, A | - |
| AutoSum Count | Alt+M, U, C | - |
| AutoSum Max | Alt+M, U, M | - |
| AutoSum Min | Alt+M, U, I | - |
| Recently Used | Alt+M, E | - |
| Financial Functions | Alt+M, I | - |
| Logical Functions | Alt+M, L | - |
| Text Functions | Alt+M, T | - |
| Date & Time Functions | Alt+M, D | - |
| Lookup & Reference | Alt+M, O | - |
| Math & Trig Functions | Alt+M, G | - |
| More Functions | Alt+M, Q | - |

### Defined Names Group (Alt+M)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Name Manager | Alt+M, M | Ctrl+F3 |
| Define Name | Alt+M, M, N | - |
| Use in Formula | Alt+M, U | - |
| Create from Selection | Alt+M, C | Ctrl+Shift+F3 |

### Formula Auditing Group (Alt+M)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Trace Precedents | Alt+M, P | - |
| Trace Dependents | Alt+M, D | - |
| Remove Arrows | Alt+M, A, A | - |
| Remove Precedent Arrows | Alt+M, A, P | - |
| Remove Dependent Arrows | Alt+M, A, D | - |
| Show Formulas | Alt+M, H | Ctrl+` |
| Error Checking | Alt+M, K | - |
| Evaluate Formula | Alt+M, V | - |
| Watch Window | Alt+M, W | - |

### Calculation Group (Alt+M)
| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Calculation Options | Alt+M, X | - |
| Calculate Now | Alt+M, X, N | F9 |
| Calculate Sheet | Alt+M, X, S | Shift+F9 |

---

## Keyboard Shortcuts Summary

| Action | Alt Shortcut | Other |
|--------|--------------|-------|
| Insert Function dialog | Alt+M, F | Shift+F3 |
| AutoSum | Alt+M, U, S | Alt+= |
| Name Manager | Alt+M, M | Ctrl+F3 |
| Create names from selection | Alt+M, C | Ctrl+Shift+F3 |
| Paste name in formula | - | F3 |
| Show Formulas toggle | Alt+M, H | Ctrl+` |
| Calculate all workbooks | Alt+M, X, N | F9 |
| Calculate active sheet | Alt+M, X, S | Shift+F9 |
| Calculate all (forced) | - | Ctrl+Alt+F9 |
| Recalculate all (rebuild) | - | Ctrl+Shift+Alt+F9 |
| Trace Precedents | Alt+M, P | - |
| Trace Dependents | Alt+M, D | - |
| Evaluate formula | Alt+M, V | - |

---

## Formula Entry Tips

| Tip | Description |
|-----|-------------|
| **Tab** | Accept AutoComplete suggestion |
| **Ctrl+Shift+A** | Insert argument names |
| **F4** | Toggle absolute/relative reference |
| **F2** | Edit cell formula |
| **Esc** | Cancel formula entry |
| **Ctrl+Shift+Enter** | Array formula (legacy) |
| **Ctrl+`** | Show/hide formulas |

---

[🎗️ Back to Ribbon Reference](./README.md) | [🏠 Back to Home](../README.md)
