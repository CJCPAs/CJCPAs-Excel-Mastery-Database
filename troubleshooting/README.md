# Excel Troubleshooting Guide

> **Solutions to the most common Excel problems**

---

## Formula Problems

### Formula Shows as Text (Not Calculating)

**Symptoms:**
- Formula displays as text instead of result
- No calculation happens when you enter formula

**Causes & Fixes:**

| Cause | Fix |
|-------|-----|
| Cell formatted as Text | Format as General, re-enter formula |
| Space before = | Delete space, formula should start with = |
| Show Formulas mode on | Press Ctrl+` or Formulas → Show Formulas |
| Apostrophe before = | Delete ' before the = sign |

**Quick Fix:**
1. Select the cell
2. Press Ctrl+1 (Format Cells)
3. Choose "General"
4. Press F2 then Enter

---

### Formula Returns Wrong Result

**Debugging Steps:**

1. **Check cell references**
   - Are absolute references ($) where needed?
   - Did references shift when copying?

2. **Evaluate formula step-by-step**
   - Formulas → Evaluate Formula
   - Or select part of formula, press F9 to see its value

3. **Check data types**
   - Numbers stored as text won't calculate correctly
   - Look for green triangles in cells

4. **Check for hidden characters**
   - Use `=LEN(A1)` to check actual length
   - Use `=TRIM(CLEAN(A1))` to remove hidden chars

5. **Test with simple data**
   - Replace complex references with simple values
   - Isolate which part fails

---

### Circular Reference Error

**Message:** "There is a circular reference..."

**Cause:** Formula refers to itself, directly or indirectly

**How to Find It:**
1. Formulas → Error Checking → Circular References
2. Shows which cells have circular references

**Common Causes:**
- Running total that includes its own cell
- Formula referencing entire column including itself
- Accidentally including formula cell in SUM range

**Fix:**
- Adjust range to exclude the formula cell
- Use a separate column for running totals
- Use absolute references to limit range

---

### #N/A Error

**Cause:** Lookup function couldn't find the value

**Fixes:**
| Issue | Solution |
|-------|----------|
| Typo in lookup value | Check spelling |
| Extra spaces | Use TRIM on lookup value and data |
| Numbers vs text mismatch | Ensure same type (VALUE, TEXT) |
| Value doesn't exist | Wrap in IFERROR or IFNA |

**Example Fix:**
```excel
=IFNA(VLOOKUP(TRIM(A2), Table, 2, FALSE), "Not found")
```

---

### #DIV/0! Error

**Cause:** Division by zero or empty cell

**Fixes:**
```excel
=IF(B2=0, "", A2/B2)              // Check before dividing
=IFERROR(A2/B2, 0)                // Catch after
=IF(B2<>0, A2/B2, "No data")      // With message
```

---

### #VALUE! Error

**Cause:** Wrong type of argument

**Common Causes:**
- Text in a numeric calculation
- Array size mismatch in SUMPRODUCT
- Non-numeric data in range

**Fixes:**
1. Check all cells contain expected data types
2. Use `=ISNUMBER(A1)` or `=ISTEXT(A1)` to verify
3. Convert text to numbers: `=VALUE(A1)` or multiply by 1

---

### #REF! Error

**Cause:** Invalid cell reference (usually deleted cells)

**Common Causes:**
- Deleted a row/column that formula referenced
- Pasted over referenced cells
- Moved cells breaking the reference

**Fix:**
- Undo (Ctrl+Z) if just happened
- Find & Replace to fix references
- Rebuild formula if necessary

---

## Data Problems

### Numbers Stored as Text

**Symptoms:**
- Green triangle in cell corner
- SUM returns 0 or ignores some cells
- VLOOKUP can't find matching numbers

**Quick Fixes:**

**Method 1: Error Option**
1. Select cells with green triangles
2. Click the error icon
3. Choose "Convert to Number"

**Method 2: Paste Special**
1. Enter 1 in an empty cell
2. Copy that cell
3. Select problem cells
4. Paste Special → Multiply

**Method 3: Text to Columns**
1. Select the column
2. Data → Text to Columns
3. Click Finish immediately

---

### Dates Not Recognized

**Symptoms:**
- Dates align left (text) instead of right (numbers)
- Date functions return errors
- Sorting doesn't work correctly

**Fixes:**

**If format is MM/DD/YYYY (or your locale):**
```excel
=DATEVALUE(A1)
```

**If format is non-standard:**
```excel
=DATE(RIGHT(A1,4), MID(A1,4,2), LEFT(A1,2))  // For DD/MM/YYYY
```

**Then apply date format** to the result cells

---

### Duplicate Data

**To Find Duplicates:**
1. Select range
2. Conditional Formatting → Highlight Cell Rules → Duplicate Values

**To Remove Duplicates:**
1. Select range
2. Data → Remove Duplicates
3. Choose columns to check

**To Get Unique List (Excel 365):**
```excel
=UNIQUE(A2:A100)
```

---

## Performance Problems

### Slow Calculation

**Quick Fixes:**

1. **Switch to Manual Calculation**
   - Formulas → Calculation Options → Manual
   - Press F9 when ready to calculate

2. **Reduce volatile functions**
   - NOW(), OFFSET(), INDIRECT(), RAND()
   - Replace with non-volatile alternatives

3. **Limit range references**
   - Use A2:A1000 instead of A:A
   - Use Tables for auto-sizing

4. **Simplify formulas**
   - Break into helper columns
   - Reduce nested functions

5. **Remove conditional formatting**
   - Especially formulas-based rules on large ranges

---

### File Won't Open

**Try These:**

1. **Open in Safe Mode**
   - Hold Ctrl while starting Excel
   - Try opening file

2. **Repair the file**
   - File → Open
   - Select file, click arrow next to Open
   - Choose "Open and Repair"

3. **Try different program**
   - Open in Google Sheets (web)
   - Open in LibreOffice

4. **Extract data**
   - Change .xlsx to .zip
   - Open and extract files
   - xl/worksheets/ contains XML data

---

### Excel Keeps Crashing

**Fixes:**

1. **Update Excel** - Check for Office updates

2. **Disable add-ins**
   - File → Options → Add-ins
   - Manage COM Add-ins → Go
   - Uncheck all, restart, re-enable one by one

3. **Repair Office**
   - Control Panel → Programs → Microsoft Office
   - Change → Repair

4. **Start in Safe Mode**
   - Hold Ctrl while starting Excel

---

## Printing Problems

### Print Preview Doesn't Match Screen

**Fixes:**

1. **Set print area**
   - Select exact range
   - Page Layout → Print Area → Set Print Area

2. **Adjust page breaks**
   - View → Page Break Preview
   - Drag blue lines to adjust

3. **Scale to fit**
   - Page Layout → Scale to Fit
   - Set width to 1 page

---

### Headers Cut Off

**Fixes:**

1. **Repeat rows at top**
   - Page Layout → Print Titles
   - Rows to repeat at top: select header row

2. **Widen columns**
   - Double-click column border to auto-fit

3. **Wrap text**
   - Home → Wrap Text
   - Or Format Cells → Alignment → Wrap text

---

## Common Symptoms Quick Reference

| Symptom | Likely Cause | First Fix |
|---------|--------------|-----------|
| Formula shows as text | Cell formatted as text | Change to General, re-enter |
| SUM returns 0 | Numbers as text | Paste Special → Multiply by 1 |
| #N/A in lookup | Value not found | TRIM both sides, check data types |
| Green triangles | Numbers as text or inconsistency | Click error icon |
| Slow calculations | Volatile functions | Change to Manual calc |
| File very large | Pictures, embedded objects | Check for hidden images |
| Can't delete rows | Merged cells or protected | Unmerge, unprotect |
| Formulas not updating | Manual calculation mode | Press F9 or change to Auto |

---

## Getting Help

### Built-in Help
- Press F1 for Excel Help
- Right-click formula → Show Help

### Audit Tools
- Formulas → Trace Precedents (what feeds into this)
- Formulas → Trace Dependents (what uses this)
- Formulas → Evaluate Formula (step through)

### Error Checking
- Formulas → Error Checking
- Reviews all errors in worksheet

---

[🏠 Back to Home](../README.md)
