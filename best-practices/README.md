# Excel Best Practices

> **Write better formulas, organize better workbooks, and work more efficiently**

---

## Formula Best Practices

### 1. Use Named Ranges
**Instead of:** `=VLOOKUP(A2, Sheet2!$A$2:$D$500, 3, FALSE)`

**Use:** `=VLOOKUP(A2, ProductTable, 3, FALSE)`

**How to create:**
1. Select your data range
2. Click in the Name Box (left of formula bar)
3. Type a name (no spaces, start with letter)
4. Press Enter

**Benefits:**
- Formulas are easier to read
- Less likely to break when sheets change
- Self-documenting

---

### 2. Use Tables (Ctrl+T)
Convert your data to an Excel Table for automatic benefits:

| Feature | Benefit |
|---------|---------|
| Structured references | `=SUM(Sales[Amount])` instead of `=SUM(B:B)` |
| Auto-expanding | New rows automatically included in formulas |
| Filter/sort built-in | No need to set up separately |
| Consistent formatting | Alternating rows, headers |

**How to create:**
1. Click any cell in your data
2. Press Ctrl+T
3. Ensure "My table has headers" is checked
4. OK

---

### 3. Avoid Entire Column References When Possible
**Instead of:** `=SUMIF(A:A, "North", B:B)`

**Use:** `=SUMIF(A2:A1000, "North", B2:B1000)`

**Why:**
- Faster calculation
- Less memory usage
- Prevents accidental inclusion of header

**Exception:** Okay for small datasets or when data grows unpredictably

---

### 4. Use Absolute References Correctly
| Reference | When It Changes | Use Case |
|-----------|----------------|----------|
| A1 | Both row and column | Rarely needed |
| $A1 | Only row | Lookup range columns |
| A$1 | Only column | Lookup range rows |
| $A$1 | Never | Constants, lookup tables |

**Toggle with F4** while editing a reference

---

### 5. Break Complex Formulas Into Steps
**Instead of:**
```excel
=IF(AND(VLOOKUP(A2,Data,3,0)>100,MONTH(B2)=12),VLOOKUP(A2,Data,4,0)*1.1,VLOOKUP(A2,Data,4,0))
```

**Use LET (Excel 365):**
```excel
=LET(
    quantity, VLOOKUP(A2, Data, 3, 0),
    price, VLOOKUP(A2, Data, 4, 0),
    is_december, MONTH(B2) = 12,
    is_bulk, quantity > 100,
    IF(AND(is_december, is_bulk), price * 1.1, price)
)
```

**Or use helper columns** in older Excel versions.

---

### 6. Always Use Exact Match in Lookups
**Use FALSE (or 0) for exact match:**
```excel
=VLOOKUP(A2, Table, 2, FALSE)    ✅ Correct
=VLOOKUP(A2, Table, 2)           ❌ Dangerous (defaults to approximate)
```

**Only use TRUE** when you specifically need approximate matching (like tax brackets).

---

### 7. Comment Your Complex Formulas
Add a note explaining what the formula does:
1. Right-click the cell
2. Insert Comment (or New Note in newer Excel)
3. Explain the logic

Or use a cell above/beside for documentation.

---

## Workbook Organization

### 1. Consistent Sheet Naming
**Good Names:**
- `Data_Input`
- `Calculations`
- `Dashboard`
- `Reference_Tables`

**Avoid:**
- `Sheet1`, `Sheet2` (meaningless)
- `John's Data` (spaces and special chars cause problems)
- Very long names (hard to reference)

---

### 2. Separate Data from Presentation
| Sheet | Contents |
|-------|----------|
| Raw Data | Original imported data, untouched |
| Processed | Cleaned, formatted data |
| Calculations | Helper columns, intermediate results |
| Output/Dashboard | What users see |

**Benefits:**
- Easy to refresh data without breaking reports
- Clear audit trail
- Easier troubleshooting

---

### 3. Document Your Workbook
Create a "README" or "Instructions" sheet containing:
- Purpose of the workbook
- Data sources and refresh dates
- How to use it
- Who to contact for help
- Change log

---

### 4. Protect What Matters
| Protection | What It Does | When to Use |
|------------|-------------|-------------|
| Lock cells | Prevent editing | Formulas, headers |
| Hide formulas | Formula bar shows nothing | Sensitive calculations |
| Protect sheet | Lock structure | Prevent accidental changes |
| Protect workbook | Lock sheet order | Prevent sheet deletion |

**How to protect cells:**
1. Select cells to ALLOW editing
2. Format Cells → Protection → Uncheck "Locked"
3. Review → Protect Sheet

---

## Performance Best Practices

### 1. Avoid Volatile Functions
These recalculate every time ANYTHING changes:

| Volatile | Alternative |
|----------|-------------|
| NOW() | TODAY() (less volatile) |
| OFFSET() | INDEX() |
| INDIRECT() | Direct references or Tables |
| RAND() | Use once, paste as values |

---

### 2. Use Efficient Functions
| Slow | Fast | Notes |
|------|------|-------|
| SUMPRODUCT for simple sums | SUMIFS | SUMIFS optimized for criteria |
| VLOOKUP with entire columns | XLOOKUP or limited ranges | XLOOKUP stops at first match |
| Array formulas everywhere | Dynamic arrays | Modern Excel handles better |
| Nested IFs | IFS or SWITCH | Cleaner and can be faster |

---

### 3. Minimize External Links
- External links slow opening
- Broken links cause errors
- Consider copying data instead of linking

---

### 4. Turn Off Automatic Calculation When Needed
For large workbooks during data entry:
1. Formulas → Calculation Options → Manual
2. Press F9 when you want to calculate
3. Remember to turn back to Automatic!

---

## Data Entry Best Practices

### 1. Use Data Validation
Prevent bad data at entry:
- **Lists** - Dropdown of allowed values
- **Numbers** - Min/max ranges
- **Dates** - Valid date ranges
- **Custom** - Formula-based rules

**How:**
1. Select cells
2. Data → Data Validation
3. Set criteria

---

### 2. Consistent Data Types
| Column | Should Contain | Enforce With |
|--------|---------------|--------------|
| Dates | Only dates | Data validation |
| Currency | Numbers, not text | Cell format + validation |
| Categories | From fixed list | Dropdown list |
| IDs | Consistent format | Custom validation |

---

### 3. No Merged Cells in Data Areas
Merged cells break:
- Sorting
- Filtering
- Formulas referencing the range
- Copy/paste operations

**Use Center Across Selection** instead for visual centering without merging.

---

## Keyboard Efficiency

### Essential Shortcuts
| Action | Windows | Mac |
|--------|---------|-----|
| Navigate to edge of data | Ctrl+Arrow | ⌘+Arrow |
| Select to edge | Ctrl+Shift+Arrow | ⌘+Shift+Arrow |
| Toggle absolute reference | F4 | ⌘+T |
| Edit cell | F2 | Control+U |
| AutoSum | Alt+= | ⌘+Shift+T |
| Fill down | Ctrl+D | ⌘+D |
| Create table | Ctrl+T | ⌘+T |
| Format cells | Ctrl+1 | ⌘+1 |

---

## Version Control

### 1. Date Your Backups
`Budget_2024-01-15.xlsx` instead of `Budget_v2.xlsx`

### 2. Track Changes
- Use built-in Track Changes (Review tab)
- Or keep a changelog sheet

### 3. Save Regularly
- Enable AutoSave for OneDrive/SharePoint files
- Save versions at milestones

---

## Error Prevention Checklist

Before sharing a workbook:

- [ ] Test all formulas with edge cases
- [ ] Check that named ranges are still valid
- [ ] Verify links to external workbooks
- [ ] Remove personal/sensitive data
- [ ] Protect cells containing formulas
- [ ] Add data validation where needed
- [ ] Include documentation/instructions
- [ ] Save a backup copy

---

[🏠 Back to Home](../README.md)
