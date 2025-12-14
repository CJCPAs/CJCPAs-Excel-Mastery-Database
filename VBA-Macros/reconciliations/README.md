# Reconciliation VBA Macros

> **There's a VBA for That!** - Match transactions, find variances, complete reconciliations faster

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [CreateReconTemplate](#create-reconciliation-template) | Generate reconciliation worksheet |
| [MatchTransactions](#match-transactions) | Auto-match by amount or reference |
| [MatchByMultipleCriteria](#match-by-multiple-criteria) | Match on amount AND date |
| [HighlightUnmatched](#highlight-unmatched-items) | Color-code unmatched items |
| [CalculateVariance](#calculate-variance) | Compute and display differences |
| [AgeOutstandingItems](#age-outstanding-items) | Categorize by age buckets |
| [BankRecTemplate](#bank-reconciliation-template) | Bank rec specific template |
| [ReconcileToZero](#reconcile-to-zero-check) | Verify recon balances to zero |
| [AddTickmarks](#add-reconciliation-tickmarks) | Insert standard tick marks |
| [GenerateReconReport](#generate-reconciliation-report) | Summary report of recon status |

---

## Create Reconciliation Template

Generate a standard account reconciliation template.

```vba
Sub CreateReconTemplate()
    '================================================
    ' Create Account Reconciliation Template
    '================================================

    Dim ws As Worksheet
    Dim sheetName As String

    sheetName = InputBox("Enter account name/number:", "New Reconciliation", "Acct-" & Format(Date, "YYYYMM"))
    If sheetName = "" Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    On Error Resume Next
    ws.Name = Left(sheetName, 31)
    On Error GoTo 0

    Application.ScreenUpdating = False

    With ws
        ' Header
        .Range("A1").Value = "ACCOUNT RECONCILIATION"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        ' Account Info
        .Range("A3").Value = "Account:"
        .Range("B3").Value = sheetName
        .Range("A4").Value = "Period End:"
        .Range("B4").Value = Application.WorksheetFunction.EoMonth(Date, 0)
        .Range("B4").NumberFormat = "mmmm d, yyyy"
        .Range("A5").Value = "Prepared By:"
        .Range("B5").Value = Environ("USERNAME")
        .Range("A6").Value = "Date Prepared:"
        .Range("B6").Value = Date

        ' GL Balance section
        .Range("A8").Value = "PER GENERAL LEDGER"
        .Range("A8").Font.Bold = True
        .Range("A9").Value = "GL Balance:"
        .Range("B9").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Sub-ledger section
        .Range("A11").Value = "PER SUB-LEDGER / SUPPORT"
        .Range("A11").Font.Bold = True
        .Range("A12").Value = "Beginning Balance:"
        .Range("A13").Value = "Add: Additions"
        .Range("A14").Value = "Less: Reductions"
        .Range("A15").Value = "Ending Balance:"
        .Range("A15").Font.Bold = True
        .Range("B12:B15").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("B15").Formula = "=B12+B13-B14"

        ' Reconciling Items
        .Range("A17").Value = "RECONCILING ITEMS"
        .Range("A17").Font.Bold = True

        .Range("A18").Value = "Date"
        .Range("B18").Value = "Reference"
        .Range("C18").Value = "Description"
        .Range("D18").Value = "Amount"
        .Range("E18").Value = "Status"

        .Range("A18:E18").Font.Bold = True
        .Range("A18:E18").Interior.Color = RGB(0, 51, 102)
        .Range("A18:E18").Font.Color = RGB(255, 255, 255)

        ' Reconciling items data rows
        .Range("D19:D38").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Totals
        .Range("A40").Value = "Total Reconciling Items:"
        .Range("A40").Font.Bold = True
        .Range("D40").Formula = "=SUM(D19:D38)"
        .Range("D40").Font.Bold = True

        ' Difference
        .Range("A42").Value = "Reconciled Balance:"
        .Range("B42").Formula = "=B15+D40"
        .Range("B42").Font.Bold = True
        .Range("B42").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        .Range("A43").Value = "DIFFERENCE:"
        .Range("B43").Formula = "=B9-B42"
        .Range("B43").Font.Bold = True
        .Range("B43").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Conditional format for difference (should be zero)
        .Range("B43").FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="0"
        .Range("B43").FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        .Range("B43").FormatConditions(1).Font.Color = RGB(255, 255, 255)

        ' Column widths
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 40
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 12

        ' Signature lines
        .Range("A46").Value = "Reviewed By:"
        .Range("A47").Value = "Review Date:"

    End With

    Application.ScreenUpdating = True

    MsgBox "Reconciliation template created!", vbInformation

End Sub
```

---

## Match Transactions

Automatically match transactions between two columns based on amount.

```vba
Sub MatchTransactions()
    '================================================
    ' Match Transactions by Amount
    ' Compares two columns and marks matches
    '================================================

    Dim ws As Worksheet
    Dim rngSource As Range
    Dim rngCompare As Range
    Dim cellSource As Range
    Dim cellCompare As Range
    Dim matchCount As Long
    Dim sourceCol As Integer
    Dim compareCol As Integer
    Dim statusCol As Integer

    Set ws = ActiveSheet

    ' Get source column
    sourceCol = Application.InputBox("Enter SOURCE column number (e.g., 1 for A):", "Match Transactions", 1, Type:=1)
    If sourceCol = 0 Then Exit Sub

    ' Get compare column
    compareCol = Application.InputBox("Enter COMPARE column number:", "Match Transactions", 2, Type:=1)
    If compareCol = 0 Then Exit Sub

    ' Status column
    statusCol = Application.InputBox("Enter STATUS column number:", "Match Transactions", 3, Type:=1)
    If statusCol = 0 Then Exit Sub

    Application.ScreenUpdating = False

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, sourceCol).End(xlUp).Row

    ' Loop through source column
    For Each cellSource In ws.Range(ws.Cells(2, sourceCol), ws.Cells(lastRow, sourceCol))
        If cellSource.Value <> "" And ws.Cells(cellSource.Row, statusCol).Value = "" Then

            ' Search compare column for match
            For Each cellCompare In ws.Range(ws.Cells(2, compareCol), ws.Cells(lastRow, compareCol))
                If cellCompare.Value <> "" And ws.Cells(cellCompare.Row, statusCol).Value = "" Then

                    ' Check if amounts match (with rounding tolerance)
                    If Round(cellSource.Value, 2) = Round(cellCompare.Value, 2) Then
                        ' Mark both as matched
                        ws.Cells(cellSource.Row, statusCol).Value = "Matched"
                        ws.Cells(cellCompare.Row, statusCol).Value = "Matched"
                        matchCount = matchCount + 1

                        ' Highlight in green
                        ws.Cells(cellSource.Row, sourceCol).Interior.Color = RGB(198, 239, 206)
                        ws.Cells(cellCompare.Row, compareCol).Interior.Color = RGB(198, 239, 206)

                        Exit For
                    End If

                End If
            Next cellCompare

        End If
    Next cellSource

    Application.ScreenUpdating = True

    MsgBox "Matched " & matchCount & " transactions.", vbInformation

End Sub
```

---

## Match by Multiple Criteria

Match transactions using amount AND date (or other criteria).

```vba
Sub MatchByMultipleCriteria()
    '================================================
    ' Match by Amount AND Date
    ' More precise matching using multiple fields
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim matchCount As Long

    ' Assumes structure:
    ' Column A: Date (Source)
    ' Column B: Amount (Source)
    ' Column C: Status (Source)
    ' Column D: Date (Compare)
    ' Column E: Amount (Compare)
    ' Column F: Status (Compare)

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        If ws.Cells(i, "C").Value = "" And ws.Cells(i, "B").Value <> "" Then

            For j = 2 To lastRow
                If ws.Cells(j, "F").Value = "" And ws.Cells(j, "E").Value <> "" Then

                    ' Match on Amount AND Date
                    If Round(ws.Cells(i, "B").Value, 2) = Round(ws.Cells(j, "E").Value, 2) And _
                       ws.Cells(i, "A").Value = ws.Cells(j, "D").Value Then

                        ws.Cells(i, "C").Value = "Matched-" & j
                        ws.Cells(j, "F").Value = "Matched-" & i

                        ws.Range(ws.Cells(i, "A"), ws.Cells(i, "B")).Interior.Color = RGB(198, 239, 206)
                        ws.Range(ws.Cells(j, "D"), ws.Cells(j, "E")).Interior.Color = RGB(198, 239, 206)

                        matchCount = matchCount + 1
                        Exit For
                    End If

                End If
            Next j

        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Matched " & matchCount & " transactions.", vbInformation

End Sub
```

---

## Highlight Unmatched Items

Color-code all unmatched (outstanding) items.

```vba
Sub HighlightUnmatchedItems()
    '================================================
    ' Highlight Unmatched Items
    ' Colors unmatched rows in yellow/red
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim statusCol As Integer
    Dim i As Long
    Dim unmatchedCount As Long

    Set ws = ActiveSheet

    statusCol = Application.InputBox("Enter STATUS column number:", "Highlight Unmatched", 3, Type:=1)
    If statusCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        If ws.Cells(i, statusCol).Value = "" Or _
           LCase(ws.Cells(i, statusCol).Value) = "unmatched" Or _
           LCase(ws.Cells(i, statusCol).Value) = "open" Then

            ws.Rows(i).Interior.Color = RGB(255, 235, 156)  ' Yellow
            unmatchedCount = unmatchedCount + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox unmatchedCount & " unmatched items highlighted.", vbInformation

End Sub

Sub ClearHighlights()
    '================================================
    ' Clear All Highlighting
    '================================================

    ActiveSheet.Cells.Interior.ColorIndex = xlNone
    MsgBox "Highlights cleared.", vbInformation

End Sub
```

---

## Calculate Variance

Calculate and display the difference between two values.

```vba
Sub CalculateVariance()
    '================================================
    ' Calculate Variance Between Two Balances
    '================================================

    Dim glBalance As Double
    Dim subBalance As Double
    Dim variance As Double
    Dim variancePercent As Double

    glBalance = Application.InputBox("Enter GL Balance:", "Variance Calculation", Type:=1)
    subBalance = Application.InputBox("Enter Sub-ledger/Support Balance:", "Variance Calculation", Type:=1)

    variance = glBalance - subBalance

    If glBalance <> 0 Then
        variancePercent = (variance / glBalance) * 100
    Else
        variancePercent = 0
    End If

    MsgBox "VARIANCE ANALYSIS" & vbCrLf & vbCrLf & _
           "GL Balance:      " & Format(glBalance, "$#,##0.00") & vbCrLf & _
           "Support Balance: " & Format(subBalance, "$#,##0.00") & vbCrLf & _
           String(35, "-") & vbCrLf & _
           "Variance:        " & Format(variance, "$#,##0.00") & vbCrLf & _
           "Variance %:      " & Format(variancePercent, "0.00") & "%", _
           IIf(Abs(variance) > 0.01, vbExclamation, vbInformation), "Variance"

End Sub

Sub InsertVarianceFormula()
    '================================================
    ' Insert Variance Formula in Selected Cell
    '================================================

    Dim cell1 As String
    Dim cell2 As String

    cell1 = Application.InputBox("Enter first cell reference (GL):", "Variance Formula", "B9")
    cell2 = Application.InputBox("Enter second cell reference (Support):", "Variance Formula", "B15")

    ActiveCell.Formula = "=" & cell1 & "-" & cell2

    ' Format and conditional formatting
    ActiveCell.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    ActiveCell.FormatConditions.Delete
    ActiveCell.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="0"
    ActiveCell.FormatConditions(1).Interior.Color = RGB(255, 199, 206)

End Sub
```

---

## Age Outstanding Items

Categorize outstanding items by age buckets.

```vba
Sub AgeOutstandingItems()
    '================================================
    ' Age Outstanding Items
    ' Categorizes items into aging buckets
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dateCol As Integer
    Dim ageCol As Integer
    Dim i As Long
    Dim itemDate As Date
    Dim itemAge As Long
    Dim asOfDate As Date

    Set ws = ActiveSheet

    dateCol = Application.InputBox("Enter DATE column number:", "Age Items", 1, Type:=1)
    If dateCol = 0 Then Exit Sub

    ageCol = Application.InputBox("Enter column for AGE BUCKET:", "Age Items", Type:=1)
    If ageCol = 0 Then Exit Sub

    asOfDate = Application.InputBox("Enter as-of date (mm/dd/yyyy):", "Age Items", Date, Type:=1)
    If asOfDate = 0 Then asOfDate = Date

    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row

    ' Add header
    ws.Cells(1, ageCol).Value = "Age Bucket"
    ws.Cells(1, ageCol).Font.Bold = True

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        If IsDate(ws.Cells(i, dateCol).Value) Then
            itemDate = ws.Cells(i, dateCol).Value
            itemAge = asOfDate - itemDate

            Select Case itemAge
                Case Is <= 30
                    ws.Cells(i, ageCol).Value = "0-30 Days"
                    ws.Cells(i, ageCol).Interior.Color = RGB(198, 239, 206)  ' Green
                Case 31 To 60
                    ws.Cells(i, ageCol).Value = "31-60 Days"
                    ws.Cells(i, ageCol).Interior.Color = RGB(255, 235, 156)  ' Yellow
                Case 61 To 90
                    ws.Cells(i, ageCol).Value = "61-90 Days"
                    ws.Cells(i, ageCol).Interior.Color = RGB(255, 199, 206)  ' Light Red
                Case Is > 90
                    ws.Cells(i, ageCol).Value = "Over 90 Days"
                    ws.Cells(i, ageCol).Interior.Color = RGB(255, 0, 0)  ' Red
                    ws.Cells(i, ageCol).Font.Color = RGB(255, 255, 255)
            End Select
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Aging complete!", vbInformation

End Sub
```

---

## Bank Reconciliation Template

Create a specific bank reconciliation template.

```vba
Sub BankRecTemplate()
    '================================================
    ' Create Bank Reconciliation Template
    '================================================

    Dim ws As Worksheet
    Dim acctName As String

    acctName = InputBox("Enter bank account name:", "Bank Reconciliation", "Operating Account")
    If acctName = "" Then Exit Sub

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    On Error Resume Next
    ws.Name = "BankRec-" & Format(Date, "YYYYMM")
    On Error GoTo 0

    Application.ScreenUpdating = False

    With ws
        ' Title
        .Range("A1").Value = "BANK RECONCILIATION"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = acctName
        .Range("A3").Value = "Period Ending: " & Format(Application.WorksheetFunction.EoMonth(Date, 0), "mmmm d, yyyy")

        ' Balance Per Bank Statement
        .Range("A5").Value = "BALANCE PER BANK STATEMENT"
        .Range("A5").Font.Bold = True
        .Range("A5").Interior.Color = RGB(0, 51, 102)
        .Range("A5").Font.Color = RGB(255, 255, 255)

        .Range("A6").Value = "Ending Balance per Bank:"
        .Range("C6").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Deposits in Transit
        .Range("A8").Value = "ADD: Deposits in Transit"
        .Range("A8").Font.Bold = True
        .Range("A9").Value = "Date"
        .Range("B9").Value = "Reference"
        .Range("C9").Value = "Amount"
        .Range("A9:C9").Font.Bold = True
        .Range("C10:C19").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("A20").Value = "Total Deposits in Transit:"
        .Range("C20").Formula = "=SUM(C10:C19)"
        .Range("C20").Font.Bold = True

        ' Outstanding Checks
        .Range("A22").Value = "LESS: Outstanding Checks"
        .Range("A22").Font.Bold = True
        .Range("A23").Value = "Check #"
        .Range("B23").Value = "Date"
        .Range("C23").Value = "Amount"
        .Range("A23:C23").Font.Bold = True
        .Range("C24:C43").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("A44").Value = "Total Outstanding Checks:"
        .Range("C44").Formula = "=SUM(C24:C43)"
        .Range("C44").Font.Bold = True

        ' Adjusted Bank Balance
        .Range("A46").Value = "ADJUSTED BANK BALANCE"
        .Range("A46").Font.Bold = True
        .Range("C46").Formula = "=C6+C20-C44"
        .Range("C46").Font.Bold = True
        .Range("C46").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("A46:C46").Interior.Color = RGB(221, 235, 247)

        ' Balance Per Books
        .Range("E5").Value = "BALANCE PER BOOKS"
        .Range("E5").Font.Bold = True
        .Range("E5").Interior.Color = RGB(0, 51, 102)
        .Range("E5").Font.Color = RGB(255, 255, 255)

        .Range("E6").Value = "GL Balance:"
        .Range("G6").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Adjustments
        .Range("E8").Value = "ADD: Interest Income"
        .Range("E9").Value = "ADD: Other Credits"
        .Range("E10").Value = "LESS: Bank Fees"
        .Range("E11").Value = "LESS: NSF Checks"
        .Range("E12").Value = "LESS: Other Debits"
        .Range("G8:G12").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Adjusted Book Balance
        .Range("E14").Value = "ADJUSTED BOOK BALANCE"
        .Range("E14").Font.Bold = True
        .Range("G14").Formula = "=G6+G8+G9-G10-G11-G12"
        .Range("G14").Font.Bold = True
        .Range("G14").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("E14:G14").Interior.Color = RGB(221, 235, 247)

        ' Difference
        .Range("E16").Value = "DIFFERENCE:"
        .Range("E16").Font.Bold = True
        .Range("G16").Formula = "=C46-G14"
        .Range("G16").Font.Bold = True
        .Range("G16").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

        ' Conditional formatting for difference
        .Range("G16").FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="0"
        .Range("G16").FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        .Range("G16").FormatConditions(1).Font.Color = RGB(255, 255, 255)

        ' Column widths
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 3
        .Columns("E").ColumnWidth = 25
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15

        ' Signatures
        .Range("A48").Value = "Prepared By:"
        .Range("A49").Value = "Date:"
        .Range("E48").Value = "Reviewed By:"
        .Range("E49").Value = "Date:"

    End With

    Application.ScreenUpdating = True

    MsgBox "Bank Reconciliation template created!", vbInformation

End Sub
```

---

## Reconcile to Zero Check

Verify that the reconciliation balances to zero.

```vba
Sub ReconcileToZeroCheck()
    '================================================
    ' Check if Reconciliation Balances to Zero
    '================================================

    Dim ws As Worksheet
    Dim diffCell As Range
    Dim diffValue As Double

    Set ws = ActiveSheet

    ' Ask user to select the difference cell
    On Error Resume Next
    Set diffCell = Application.InputBox("Select the DIFFERENCE cell:", "Zero Check", Type:=8)
    On Error GoTo 0

    If diffCell Is Nothing Then Exit Sub

    diffValue = diffCell.Value

    If Round(diffValue, 2) = 0 Then
        MsgBox "RECONCILIATION COMPLETE!" & vbCrLf & vbCrLf & _
               "Difference: $0.00" & vbCrLf & vbCrLf & _
               "The reconciliation balances.", vbInformation, "Reconciled"

        diffCell.Interior.Color = RGB(198, 239, 206)  ' Green
        diffCell.Font.Bold = True
    Else
        MsgBox "RECONCILIATION DOES NOT BALANCE" & vbCrLf & vbCrLf & _
               "Difference: " & Format(diffValue, "$#,##0.00") & vbCrLf & vbCrLf & _
               "Please review for errors.", vbCritical, "Out of Balance"

        diffCell.Interior.Color = RGB(255, 199, 206)  ' Red
        diffCell.Font.Bold = True
    End If

End Sub
```

---

## Add Reconciliation Tickmarks

Insert standard audit/reconciliation tick marks.

```vba
Sub AddReconciliationTickmarks()
    '================================================
    ' Add Tickmark Legend and Symbols
    '================================================

    Dim ws As Worksheet
    Dim startRow As Long

    Set ws = ActiveSheet

    ' Find a good spot for legend
    startRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row + 3

    With ws
        .Range("I" & startRow).Value = "TICKMARK LEGEND"
        .Range("I" & startRow).Font.Bold = True
        .Range("I" & startRow).Interior.Color = RGB(0, 51, 102)
        .Range("I" & startRow).Font.Color = RGB(255, 255, 255)

        .Range("I" & startRow + 1).Value = ChrW(10003)  ' Checkmark
        .Range("J" & startRow + 1).Value = "Agreed to source document"

        .Range("I" & startRow + 2).Value = "F"
        .Range("J" & startRow + 2).Value = "Footed / Recalculated"

        .Range("I" & startRow + 3).Value = "T"
        .Range("J" & startRow + 3).Value = "Traced to supporting detail"

        .Range("I" & startRow + 4).Value = "TB"
        .Range("J" & startRow + 4).Value = "Agreed to Trial Balance"

        .Range("I" & startRow + 5).Value = "BS"
        .Range("J" & startRow + 5).Value = "Agreed to Bank Statement"

        .Range("I" & startRow + 6).Value = "PY"
        .Range("J" & startRow + 6).Value = "Agreed to Prior Year"

        .Range("I" & startRow + 7).Value = "C"
        .Range("J" & startRow + 7).Value = "Confirmed"

        .Range("I" & startRow + 8).Value = "*"
        .Range("J" & startRow + 8).Value = "See comment"

        .Columns("I").ColumnWidth = 5
        .Columns("J").ColumnWidth = 30

    End With

    MsgBox "Tickmark legend added!", vbInformation

End Sub

Sub InsertTickmark()
    '================================================
    ' Insert Tickmark in Selected Cell
    '================================================

    Dim tickmark As String

    tickmark = InputBox("Enter tickmark:" & vbCrLf & vbCrLf & _
                        ChrW(10003) & " = Agreed" & vbCrLf & _
                        "F = Footed" & vbCrLf & _
                        "T = Traced" & vbCrLf & _
                        "TB = Trial Balance" & vbCrLf & _
                        "C = Confirmed" & vbCrLf & _
                        "* = See comment", _
                        "Insert Tickmark", ChrW(10003))

    If tickmark <> "" Then
        Selection.Value = tickmark
        Selection.HorizontalAlignment = xlCenter
        Selection.Font.Color = RGB(0, 112, 192)  ' Blue
        Selection.Font.Bold = True
    End If

End Sub
```

---

## Generate Reconciliation Report

Create summary report of reconciliation status.

```vba
Sub GenerateReconciliationReport()
    '================================================
    ' Generate Reconciliation Summary Report
    '================================================

    Dim ws As Worksheet
    Dim wsReport As Worksheet
    Dim reportRow As Long
    Dim reconSheets As Collection
    Dim sheetName As Variant

    ' Create or clear report sheet
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Recon_Summary")
    On Error GoTo 0

    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsReport.Name = "Recon_Summary"
    Else
        wsReport.Cells.Clear
    End If

    Application.ScreenUpdating = False

    With wsReport
        .Range("A1").Value = "RECONCILIATION STATUS SUMMARY"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Generated: " & Now

        .Range("A4").Value = "Account"
        .Range("B4").Value = "Period"
        .Range("C4").Value = "GL Balance"
        .Range("D4").Value = "Support Balance"
        .Range("E4").Value = "Difference"
        .Range("F4").Value = "Status"
        .Range("A4:F4").Font.Bold = True
        .Range("A4:F4").Interior.Color = RGB(0, 51, 102)
        .Range("A4:F4").Font.Color = RGB(255, 255, 255)

        reportRow = 5

        ' Loop through all sheets looking for reconciliations
        For Each ws In ThisWorkbook.Worksheets
            If Left(ws.Name, 4) = "Acct" Or _
               Left(ws.Name, 7) = "BankRec" Or _
               InStr(1, ws.Name, "Recon", vbTextCompare) > 0 Then

                .Cells(reportRow, 1).Value = ws.Name

                ' Try to find key values (adjust cell references as needed)
                On Error Resume Next
                .Cells(reportRow, 2).Value = ws.Range("B4").Value
                .Cells(reportRow, 3).Value = ws.Range("B9").Value
                .Cells(reportRow, 4).Value = ws.Range("B15").Value
                .Cells(reportRow, 5).Formula = "=C" & reportRow & "-D" & reportRow
                On Error GoTo 0

                ' Status based on difference
                If Abs(.Cells(reportRow, 5).Value) < 0.01 Then
                    .Cells(reportRow, 6).Value = "Reconciled"
                    .Cells(reportRow, 6).Interior.Color = RGB(198, 239, 206)
                Else
                    .Cells(reportRow, 6).Value = "OPEN"
                    .Cells(reportRow, 6).Interior.Color = RGB(255, 199, 206)
                End If

                reportRow = reportRow + 1
            End If
        Next ws

        ' Format columns
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 15
        .Columns("C:E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 12
        .Range("C5:E" & reportRow - 1).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    Application.ScreenUpdating = True

    wsReport.Activate
    MsgBox "Reconciliation summary generated!", vbInformation

End Sub
```

---

## Quick Tips for Reconciliations

| Tip | Description |
|-----|-------------|
| **Always balance to zero** | Difference should always be $0.00 |
| **Document exceptions** | Note reasons for reconciling items |
| **Follow up on old items** | Items over 90 days need investigation |
| **Standard tickmarks** | Use consistent symbols |
| **Date your work** | Include preparer and review dates |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
