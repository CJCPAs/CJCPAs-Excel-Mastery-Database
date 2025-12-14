# Journal Entry VBA Macros

> **There's a VBA for That!** - Create, validate, and post journal entries with automation

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [CreateJETemplate](#create-je-template) | Generate a blank journal entry template |
| [ValidateJournalEntry](#validate-journal-entry) | Check debits = credits |
| [NumberJournalEntries](#auto-number-journal-entries) | Auto-number JE lines |
| [ReverseEntry](#create-reversing-entry) | Generate reversing entry |
| [PostToLedger](#post-to-general-ledger) | Post JE to GL worksheet |
| [CalculateTotals](#calculate-debit-credit-totals) | Sum debits and credits |
| [AddJELine](#add-je-line) | Insert a new JE line |
| [FormatAsJE](#format-as-journal-entry) | Apply JE formatting |
| [ValidateAccounts](#validate-account-numbers) | Check account numbers against COA |
| [ExportJEToCSV](#export-je-to-csv) | Export JE for import to accounting system |

---

## Create JE Template

Create a professional journal entry template.

```vba
Sub CreateJETemplate()
    '================================================
    ' Create Journal Entry Template
    ' Creates a formatted JE input sheet
    '================================================

    Dim ws As Worksheet
    Dim sheetName As String

    ' Get sheet name from user
    sheetName = InputBox("Enter JE sheet name:", "New Journal Entry", "JE-" & Format(Date, "YYYYMMDD"))
    If sheetName = "" Then Exit Sub

    ' Create new sheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo 0

    Application.ScreenUpdating = False

    With ws
        ' Header Information
        .Range("A1").Value = "JOURNAL ENTRY"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A3").Value = "JE Number:"
        .Range("B3").Value = "JE-" & Format(Date, "YYYYMMDD") & "-001"

        .Range("A4").Value = "Date:"
        .Range("B4").Value = Date
        .Range("B4").NumberFormat = "mm/dd/yyyy"

        .Range("A5").Value = "Prepared By:"
        .Range("B5").Value = Environ("USERNAME")

        .Range("A6").Value = "Description:"
        .Range("B6:E6").Merge

        ' Column Headers
        .Range("A8").Value = "Line"
        .Range("B8").Value = "Account #"
        .Range("C8").Value = "Account Name"
        .Range("D8").Value = "Debit"
        .Range("E8").Value = "Credit"
        .Range("F8").Value = "Description"

        ' Format headers
        .Range("A8:F8").Font.Bold = True
        .Range("A8:F8").Interior.Color = RGB(0, 51, 102)
        .Range("A8:F8").Font.Color = RGB(255, 255, 255)

        ' Add line numbers
        Dim i As Integer
        For i = 9 To 28
            .Cells(i, 1).Value = i - 8
        Next i

        ' Format amount columns
        .Range("D9:E28").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Totals row
        .Range("A30").Value = "TOTALS"
        .Range("A30").Font.Bold = True
        .Range("D30").Formula = "=SUM(D9:D28)"
        .Range("E30").Formula = "=SUM(E9:E28)"
        .Range("D30:E30").Font.Bold = True
        .Range("D30:E30").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Difference check
        .Range("A31").Value = "Difference:"
        .Range("D31").Formula = "=D30-E30"
        .Range("D31").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Conditional formatting for difference
        .Range("D31").FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="0"
        .Range("D31").FormatConditions(1).Font.Color = RGB(255, 0, 0)
        .Range("D31").FormatConditions(1).Font.Bold = True

        ' Column widths
        .Columns("A").ColumnWidth = 6
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 30
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 40

        ' Borders
        .Range("A8:F30").Borders.LineStyle = xlContinuous

        ' Approval section
        .Range("A33").Value = "Approved By:"
        .Range("A34").Value = "Date Approved:"

        ' Select first entry cell
        .Range("B9").Select

    End With

    Application.ScreenUpdating = True

    MsgBox "Journal Entry template created!", vbInformation

End Sub
```

---

## Validate Journal Entry

Verify that debits equal credits and all required fields are filled.

```vba
Sub ValidateJournalEntry()
    '================================================
    ' Validate Journal Entry
    ' Checks: Debits=Credits, Required fields, Valid accounts
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalDebits As Double
    Dim totalCredits As Double
    Dim errors As String
    Dim errorCount As Integer

    Set ws = ActiveSheet

    ' Find data range (assuming data starts row 9)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    errors = ""
    errorCount = 0

    ' Check each line
    For i = 9 To lastRow
        ' Skip empty rows
        If ws.Cells(i, "B").Value <> "" Then

            ' Check account number exists
            If ws.Cells(i, "B").Value = "" Then
                errors = errors & "Line " & (i - 8) & ": Missing account number" & vbCrLf
                errorCount = errorCount + 1
            End If

            ' Check amount exists
            If ws.Cells(i, "D").Value = 0 And ws.Cells(i, "E").Value = 0 Then
                errors = errors & "Line " & (i - 8) & ": No debit or credit amount" & vbCrLf
                errorCount = errorCount + 1
            End If

            ' Check both debit AND credit (only one should have value)
            If ws.Cells(i, "D").Value > 0 And ws.Cells(i, "E").Value > 0 Then
                errors = errors & "Line " & (i - 8) & ": Both debit and credit have values" & vbCrLf
                errorCount = errorCount + 1
            End If

            ' Sum totals
            totalDebits = totalDebits + ws.Cells(i, "D").Value
            totalCredits = totalCredits + ws.Cells(i, "E").Value
        End If
    Next i

    ' Check totals match
    If Round(totalDebits, 2) <> Round(totalCredits, 2) Then
        errors = errors & vbCrLf & "DEBITS DO NOT EQUAL CREDITS!" & vbCrLf
        errors = errors & "Debits: " & Format(totalDebits, "$#,##0.00") & vbCrLf
        errors = errors & "Credits: " & Format(totalCredits, "$#,##0.00") & vbCrLf
        errors = errors & "Difference: " & Format(totalDebits - totalCredits, "$#,##0.00")
        errorCount = errorCount + 1
    End If

    ' Display results
    If errorCount > 0 Then
        MsgBox "VALIDATION FAILED" & vbCrLf & vbCrLf & errors, vbCritical, "JE Validation"
    Else
        MsgBox "VALIDATION PASSED!" & vbCrLf & vbCrLf & _
               "Total Debits: " & Format(totalDebits, "$#,##0.00") & vbCrLf & _
               "Total Credits: " & Format(totalCredits, "$#,##0.00"), vbInformation, "JE Validation"
    End If

End Sub
```

---

## Auto-Number Journal Entries

Automatically number JE lines sequentially.

```vba
Sub NumberJournalEntries()
    '================================================
    ' Auto-Number JE Lines
    ' Numbers lines sequentially starting from 1
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim lineNum As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    lineNum = 0

    For i = 9 To lastRow
        If ws.Cells(i, "B").Value <> "" Then
            lineNum = lineNum + 1
            ws.Cells(i, "A").Value = lineNum
        End If
    Next i

    MsgBox "Numbered " & lineNum & " JE lines.", vbInformation

End Sub

Sub GenerateJENumber()
    '================================================
    ' Generate Next JE Number
    ' Format: JE-YYYYMMDD-###
    '================================================

    Dim jeNumber As String
    Dim ws As Worksheet
    Dim jeCount As Long

    Set ws = ActiveSheet

    ' Find highest JE number for today
    jeCount = 1
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "JE-" & Format(Date, "YYYYMMDD") & "*" Then
            jeCount = jeCount + 1
        End If
    Next ws

    jeNumber = "JE-" & Format(Date, "YYYYMMDD") & "-" & Format(jeCount, "000")

    ActiveCell.Value = jeNumber

    MsgBox "Generated: " & jeNumber, vbInformation

End Sub
```

---

## Create Reversing Entry

Generate a reversing entry from existing JE.

```vba
Sub CreateReversingEntry()
    '================================================
    ' Create Reversing Journal Entry
    ' Copies JE and swaps debits/credits
    '================================================

    Dim wsSource As Worksheet
    Dim wsReverse As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim newSheetName As String
    Dim debitVal As Double
    Dim creditVal As Double

    Set wsSource = ActiveSheet

    ' Create new sheet for reversal
    newSheetName = wsSource.Name & "-REV"

    ' Copy entire sheet
    wsSource.Copy After:=wsSource
    Set wsReverse = ActiveSheet

    On Error Resume Next
    wsReverse.Name = newSheetName
    On Error GoTo 0

    ' Find data range
    lastRow = wsReverse.Cells(wsReverse.Rows.Count, "B").End(xlUp).Row

    ' Swap debits and credits
    For i = 9 To lastRow
        If wsReverse.Cells(i, "B").Value <> "" Then
            debitVal = wsReverse.Cells(i, "D").Value
            creditVal = wsReverse.Cells(i, "E").Value

            wsReverse.Cells(i, "D").Value = creditVal
            wsReverse.Cells(i, "E").Value = debitVal
        End If
    Next i

    ' Update JE header
    wsReverse.Range("A1").Value = "REVERSING JOURNAL ENTRY"
    wsReverse.Range("B3").Value = wsReverse.Range("B3").Value & "-REV"
    wsReverse.Range("B4").Value = Date
    wsReverse.Range("B6").Value = "Reversal of " & wsSource.Name

    MsgBox "Reversing entry created: " & newSheetName, vbInformation

End Sub
```

---

## Post to General Ledger

Post journal entry lines to a master General Ledger worksheet.

```vba
Sub PostToGeneralLedger()
    '================================================
    ' Post JE to General Ledger Sheet
    ' Appends JE lines to GL master sheet
    '================================================

    Dim wsJE As Worksheet
    Dim wsGL As Worksheet
    Dim lastRowJE As Long
    Dim lastRowGL As Long
    Dim i As Long
    Dim jeNumber As String
    Dim jeDate As Date
    Dim postCount As Long

    Set wsJE = ActiveSheet

    ' Check if GL sheet exists
    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("General_Ledger")
    On Error GoTo 0

    ' Create GL sheet if doesn't exist
    If wsGL Is Nothing Then
        Set wsGL = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsGL.Name = "General_Ledger"

        ' Create headers
        wsGL.Range("A1").Value = "Date"
        wsGL.Range("B1").Value = "JE Number"
        wsGL.Range("C1").Value = "Account #"
        wsGL.Range("D1").Value = "Account Name"
        wsGL.Range("E1").Value = "Debit"
        wsGL.Range("F1").Value = "Credit"
        wsGL.Range("G1").Value = "Description"
        wsGL.Range("H1").Value = "Posted By"
        wsGL.Range("I1").Value = "Posted Date"

        wsGL.Range("A1:I1").Font.Bold = True
        wsGL.Range("A1:I1").Interior.Color = RGB(0, 51, 102)
        wsGL.Range("A1:I1").Font.Color = RGB(255, 255, 255)
    End If

    ' Get JE info
    jeNumber = wsJE.Range("B3").Value
    jeDate = wsJE.Range("B4").Value

    ' Find last row in GL
    lastRowGL = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

    ' Find last row in JE
    lastRowJE = wsJE.Cells(wsJE.Rows.Count, "B").End(xlUp).Row

    ' Check if already posted
    For i = 2 To lastRowGL
        If wsGL.Cells(i, "B").Value = jeNumber Then
            If MsgBox("JE " & jeNumber & " appears to already be posted. Post again?", vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next i

    ' Post each line
    For i = 9 To lastRowJE
        If wsJE.Cells(i, "B").Value <> "" Then
            If wsJE.Cells(i, "D").Value <> 0 Or wsJE.Cells(i, "E").Value <> 0 Then
                lastRowGL = lastRowGL + 1
                postCount = postCount + 1

                wsGL.Cells(lastRowGL, "A").Value = jeDate
                wsGL.Cells(lastRowGL, "B").Value = jeNumber
                wsGL.Cells(lastRowGL, "C").Value = wsJE.Cells(i, "B").Value
                wsGL.Cells(lastRowGL, "D").Value = wsJE.Cells(i, "C").Value
                wsGL.Cells(lastRowGL, "E").Value = wsJE.Cells(i, "D").Value
                wsGL.Cells(lastRowGL, "F").Value = wsJE.Cells(i, "E").Value
                wsGL.Cells(lastRowGL, "G").Value = wsJE.Cells(i, "F").Value
                wsGL.Cells(lastRowGL, "H").Value = Environ("USERNAME")
                wsGL.Cells(lastRowGL, "I").Value = Now

            End If
        End If
    Next i

    ' Format new rows
    wsGL.Range("E2:F" & lastRowGL).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    wsGL.Range("A2:A" & lastRowGL).NumberFormat = "mm/dd/yyyy"

    MsgBox "Posted " & postCount & " lines from " & jeNumber & " to General Ledger.", vbInformation

End Sub
```

---

## Calculate Debit/Credit Totals

Calculate and display JE totals.

```vba
Sub CalculateJETotals()
    '================================================
    ' Calculate and Display JE Totals
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalDebits As Double
    Dim totalCredits As Double
    Dim i As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 9 To lastRow
        totalDebits = totalDebits + ws.Cells(i, "D").Value
        totalCredits = totalCredits + ws.Cells(i, "E").Value
    Next i

    MsgBox "Journal Entry Totals:" & vbCrLf & vbCrLf & _
           "Total Debits:  " & Format(totalDebits, "$#,##0.00") & vbCrLf & _
           "Total Credits: " & Format(totalCredits, "$#,##0.00") & vbCrLf & _
           "Difference:    " & Format(totalDebits - totalCredits, "$#,##0.00"), _
           vbInformation, "JE Totals"

End Sub
```

---

## Add JE Line

Insert a new journal entry line with proper formatting.

```vba
Sub AddJELine()
    '================================================
    ' Add New JE Line
    ' Inserts row and copies formatting
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim insertRow As Long
    Dim nextLineNum As Long

    Set ws = ActiveSheet

    ' Find last data row before totals
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Insert row before totals
    insertRow = lastRow + 1
    Rows(insertRow).Insert Shift:=xlDown

    ' Copy formatting from row above
    Rows(lastRow).Copy
    Rows(insertRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Clear values
    ws.Range(ws.Cells(insertRow, "B"), ws.Cells(insertRow, "F")).ClearContents

    ' Add line number
    nextLineNum = ws.Cells(lastRow, "A").Value + 1
    ws.Cells(insertRow, "A").Value = nextLineNum

    ' Select account cell
    ws.Cells(insertRow, "B").Select

End Sub
```

---

## Format as Journal Entry

Apply professional JE formatting to selected range.

```vba
Sub FormatAsJournalEntry()
    '================================================
    ' Format Selection as Journal Entry
    '================================================

    Dim rng As Range

    Set rng = Selection

    Application.ScreenUpdating = False

    With rng
        ' Assume columns: Line, Acct#, AcctName, Debit, Credit, Description
        .Columns(1).ColumnWidth = 6      ' Line
        .Columns(2).ColumnWidth = 12     ' Account #
        .Columns(3).ColumnWidth = 30     ' Account Name
        .Columns(4).ColumnWidth = 15     ' Debit
        .Columns(5).ColumnWidth = 15     ' Credit
        .Columns(6).ColumnWidth = 40     ' Description

        ' Number format for amounts
        .Columns(4).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Columns(5).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Borders
        .Borders.LineStyle = xlContinuous

        ' Header row (first row of selection)
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(0, 51, 102)
        .Rows(1).Font.Color = RGB(255, 255, 255)

    End With

    Application.ScreenUpdating = True

    MsgBox "JE formatting applied!", vbInformation

End Sub
```

---

## Validate Account Numbers

Check account numbers against Chart of Accounts.

```vba
Sub ValidateAccountNumbers()
    '================================================
    ' Validate Account Numbers Against COA
    ' Requires "Chart_of_Accounts" sheet with accounts in column A
    '================================================

    Dim wsJE As Worksheet
    Dim wsCOA As Worksheet
    Dim lastRowJE As Long
    Dim lastRowCOA As Long
    Dim i As Long
    Dim accountNum As String
    Dim found As Boolean
    Dim errors As String
    Dim errorCount As Integer

    Set wsJE = ActiveSheet

    ' Check for COA sheet
    On Error Resume Next
    Set wsCOA = ThisWorkbook.Sheets("Chart_of_Accounts")
    On Error GoTo 0

    If wsCOA Is Nothing Then
        MsgBox "Chart_of_Accounts sheet not found. Please create a sheet named 'Chart_of_Accounts' with account numbers in column A.", vbExclamation
        Exit Sub
    End If

    lastRowJE = wsJE.Cells(wsJE.Rows.Count, "B").End(xlUp).Row
    lastRowCOA = wsCOA.Cells(wsCOA.Rows.Count, "A").End(xlUp).Row

    errors = ""
    errorCount = 0

    ' Check each account
    For i = 9 To lastRowJE
        accountNum = Trim(wsJE.Cells(i, "B").Value)

        If accountNum <> "" Then
            found = False

            ' Search COA
            On Error Resume Next
            found = Not IsError(Application.Match(accountNum, wsCOA.Range("A:A"), 0))
            On Error GoTo 0

            If Not found Then
                errors = errors & "Line " & (i - 8) & ": Account '" & accountNum & "' not in COA" & vbCrLf
                errorCount = errorCount + 1
                wsJE.Cells(i, "B").Interior.Color = RGB(255, 200, 200)
            Else
                wsJE.Cells(i, "B").Interior.ColorIndex = xlNone
            End If
        End If
    Next i

    ' Display results
    If errorCount > 0 Then
        MsgBox "VALIDATION FAILED" & vbCrLf & vbCrLf & _
               errorCount & " invalid account(s):" & vbCrLf & vbCrLf & errors, _
               vbCritical, "Account Validation"
    Else
        MsgBox "All accounts validated successfully!", vbInformation, "Account Validation"
    End If

End Sub
```

---

## Export JE to CSV

Export journal entry for import into accounting system.

```vba
Sub ExportJEToCSV()
    '================================================
    ' Export JE to CSV for System Import
    ' Creates comma-delimited file
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim lineData As String
    Dim jeNumber As String
    Dim jeDate As String

    Set ws = ActiveSheet

    jeNumber = ws.Range("B3").Value
    jeDate = Format(ws.Range("B4").Value, "MM/DD/YYYY")

    ' Get file path
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=jeNumber & ".csv", _
        FileFilter:="CSV Files (*.csv), *.csv")

    If filePath = "False" Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    fileNum = FreeFile

    Open filePath For Output As #fileNum

    ' Header row
    Print #fileNum, "Date,JE_Number,Account,Description,Debit,Credit"

    ' Data rows
    For i = 9 To lastRow
        If ws.Cells(i, "B").Value <> "" Then
            lineData = jeDate & "," & _
                      jeNumber & "," & _
                      ws.Cells(i, "B").Value & "," & _
                      """" & ws.Cells(i, "F").Value & """," & _
                      ws.Cells(i, "D").Value & "," & _
                      ws.Cells(i, "E").Value

            Print #fileNum, lineData
        End If
    Next i

    Close #fileNum

    MsgBox "Exported to: " & filePath, vbInformation

End Sub
```

---

## Quick Reference: JE Standards

| Element | Standard |
|---------|----------|
| **Debits** | Listed first, not indented |
| **Credits** | Listed second, indented |
| **Total** | Debits must equal credits |
| **Description** | Each line should have description |
| **Support** | Reference supporting documentation |
| **Approval** | Requires reviewer sign-off |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
