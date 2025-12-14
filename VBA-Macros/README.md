# VBA Macros for Public Accounting

## There's a VBA for That!

> **Ready-to-copy-paste VBA code for accountants, auditors, and finance professionals**

---

## Quick Navigation

| Category | What's Inside |
|----------|---------------|
| [**Data Import & Cleanup**](./data-import-cleanup/) | Import files, remove duplicates, clean data, parse text |
| [**Journal Entries**](./journal-entries/) | Create JEs, validate debits/credits, post entries, reverse entries |
| [**Reconciliations**](./reconciliations/) | Bank recons, account matching, variance analysis, tick marks |
| [**Financial Statements**](./financial-statements/) | Generate BS/IS/CF, comparative statements, ratios |
| [**Workpapers**](./workpapers/) | Create PBC lists, index workpapers, add tickmarks, sign-offs |
| [**Reporting**](./reporting/) | Generate reports, export to PDF, email automation |
| [**Audit Procedures**](./audit-procedures/) | Sample selection, confirmations, testing procedures |
| [**Formatting & Utilities**](./formatting-utilities/) | Format cells, protect sheets, navigation, utilities |

---

## How to Use These Macros

### Step 1: Open VBA Editor
Press **Alt+F11** to open the Visual Basic Editor

### Step 2: Insert a Module
1. In the VBA Editor, go to **Insert → Module**
2. A new module window appears

### Step 3: Copy & Paste
1. Copy the VBA code you need from this repository
2. Paste it into the module window
3. Press **F5** to run (or assign to a button)

### Step 4: Save as Macro-Enabled
Save your workbook as **.xlsm** (Excel Macro-Enabled Workbook)

---

## Assigning Macros to Buttons

### Quick Access Toolbar
1. **File → Options → Quick Access Toolbar**
2. Choose **Macros** from dropdown
3. Select your macro → **Add**
4. Click **OK**

### Ribbon Button
1. **File → Options → Customize Ribbon**
2. Create a **New Tab** or **New Group**
3. Choose **Macros** from dropdown
4. Add your macro to the group

### Shape/Button on Worksheet
1. **Insert → Shapes** (choose any shape)
2. Draw shape on worksheet
3. Right-click → **Assign Macro**
4. Select your macro → **OK**

---

## Macro Security Settings

### Recommended Setting
**File → Options → Trust Center → Trust Center Settings → Macro Settings**

Choose: **"Disable all macros with notification"**

This prompts you to enable macros when opening files with macros.

### Trusted Locations
Add folders you trust to avoid prompts:
1. **Trust Center → Trusted Locations**
2. **Add new location...**
3. Browse to your folder
4. Check "Subfolders of this location are also trusted"

---

## VBA Best Practices for Accountants

### 1. Comment Your Code
```vba
' Purpose: Creates a journal entry template
' Author: Your Name
' Date: December 2024
' Usage: Run from JE worksheet
Sub CreateJournalEntry()
    ' Code here
End Sub
```

### 2. Use Meaningful Names
```vba
' Good:
Dim wsTrialBalance As Worksheet
Dim rngDebitColumn As Range
Dim dblTotalDebits As Double

' Bad:
Dim ws1 As Worksheet
Dim r As Range
Dim x As Double
```

### 3. Error Handling
```vba
Sub SafeMacro()
    On Error GoTo ErrorHandler

    ' Your code here

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

### 4. Screen Updating (Speed)
```vba
Sub FastMacro()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Your code here (runs faster)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

---

## Common VBA Objects for Accounting

| Object | What It Does | Example |
|--------|--------------|---------|
| `Workbook` | The Excel file | `ThisWorkbook`, `Workbooks("Budget.xlsx")` |
| `Worksheet` | A single sheet | `Sheets("Trial Balance")`, `ActiveSheet` |
| `Range` | Cell(s) | `Range("A1")`, `Range("A1:D100")` |
| `Cells` | Row/column reference | `Cells(1, 1)` = A1, `Cells(5, 3)` = C5 |
| `Selection` | Currently selected | `Selection.Copy`, `Selection.Clear` |

### Quick Reference: Common Operations

```vba
' Get last row with data in column A
LastRow = Cells(Rows.Count, "A").End(xlUp).Row

' Get last column with data in row 1
LastCol = Cells(1, Columns.Count).End(xlToLeft).Column

' Loop through each row
For i = 2 To LastRow
    ' Process row i
Next i

' Loop through each cell in range
For Each cell In Range("A1:A100")
    ' Process cell
Next cell

' Copy range to another location
Range("A1:D10").Copy Destination:=Range("F1")

' Clear contents (keep formatting)
Range("A1:D10").ClearContents

' Delete rows where column A = "Delete"
For i = LastRow To 2 Step -1
    If Cells(i, "A").Value = "Delete" Then
        Rows(i).Delete
    End If
Next i
```

---

## Personal Macro Workbook (PERSONAL.XLSB)

Store macros you use across ALL workbooks in PERSONAL.XLSB:

### Create PERSONAL.XLSB
1. Press **Alt+F11** (VBA Editor)
2. Go to **View → Project Explorer**
3. If PERSONAL.XLSB doesn't exist:
   - Record a dummy macro (Developer → Record Macro)
   - Choose "Personal Macro Workbook" as storage
   - Stop recording
4. PERSONAL.XLSB is created and loads automatically with Excel

### Location
- **Windows:** `C:\Users\[Username]\AppData\Roaming\Microsoft\Excel\XLSTART`
- **Mac:** `/Users/[Username]/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Excel`

---

## Keyboard Shortcuts for VBA

| Action | Shortcut |
|--------|----------|
| Open VBA Editor | **Alt+F11** |
| Run Macro Dialog | **Alt+F8** |
| Step through code | **F8** |
| Run current procedure | **F5** |
| Toggle breakpoint | **F9** |
| Immediate window | **Ctrl+G** |
| Object browser | **F2** |
| Find | **Ctrl+F** |
| Comment lines | Select + type `'` |

---

## Debugging Tips

### Use the Immediate Window (Ctrl+G)
```vba
' Print values to Immediate Window
Debug.Print "Total: " & dblTotal
Debug.Print "Current row: " & i
Debug.Print Range("A1").Value
```

### Set Breakpoints (F9)
Click in the gray margin to set a breakpoint. Code stops there when running.

### Step Through Code (F8)
Execute one line at a time to see exactly what happens.

### Watch Variables
Add variables to the Watch window to monitor their values during execution.

---

## Category Quick Links

### By Task
| I Need To... | Go To |
|--------------|-------|
| Import data from another file | [Data Import & Cleanup](./data-import-cleanup/) |
| Clean up messy data | [Data Import & Cleanup](./data-import-cleanup/) |
| Create journal entries | [Journal Entries](./journal-entries/) |
| Match transactions | [Reconciliations](./reconciliations/) |
| Generate financial reports | [Financial Statements](./financial-statements/) |
| Format my workpapers | [Workpapers](./workpapers/) |
| Create PDF reports | [Reporting](./reporting/) |
| Select audit samples | [Audit Procedures](./audit-procedures/) |
| Format cells quickly | [Formatting & Utilities](./formatting-utilities/) |

### By Complexity
| Level | Categories |
|-------|------------|
| **Beginner** | [Formatting & Utilities](./formatting-utilities/), [Workpapers](./workpapers/) |
| **Intermediate** | [Data Import](./data-import-cleanup/), [Journal Entries](./journal-entries/), [Reporting](./reporting/) |
| **Advanced** | [Reconciliations](./reconciliations/), [Financial Statements](./financial-statements/), [Audit Procedures](./audit-procedures/) |

---

## Sample: Your First Macro

Try this simple macro to format a trial balance:

```vba
Sub FormatTrialBalance()
    '================================================
    ' Format Trial Balance - Quick Formatting Macro
    '================================================

    ' Turn off screen updating for speed
    Application.ScreenUpdating = False

    ' Select all data
    Cells.Select

    ' AutoFit columns
    Cells.EntireColumn.AutoFit

    ' Format header row
    Rows(1).Font.Bold = True
    Rows(1).Interior.Color = RGB(0, 51, 102)  ' Dark blue
    Rows(1).Font.Color = RGB(255, 255, 255)   ' White text

    ' Format numbers as accounting
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row

    ' Assuming amounts in columns C and D
    Range("C2:D" & LastRow).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    ' Add borders
    Range("A1:D" & LastRow).Borders.LineStyle = xlContinuous

    ' Return to cell A1
    Range("A1").Select

    ' Turn screen updating back on
    Application.ScreenUpdating = True

    MsgBox "Trial Balance formatted!", vbInformation

End Sub
```

---

[🏠 Back to Home](../README.md) | [📚 Function Reference](../functions/) | [🎯 Solutions Library](../solutions/)
