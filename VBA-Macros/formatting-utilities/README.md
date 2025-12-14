# Formatting & Utilities VBA Macros

> **There's a VBA for That!** - Format cells, protect sheets, navigate, and general utilities

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [FormatAsAccounting](#format-as-accounting) | Apply accounting number format |
| [FormatAsPercentage](#format-as-percentage) | Apply percentage format |
| [FormatHeader](#format-header-row) | Style header row professionally |
| [AlternatingRows](#alternating-row-colors) | Add zebra striping |
| [AutoFitAll](#autofit-all-columns) | Auto-size all columns |
| [RemoveGridlines](#remove-gridlines) | Hide gridlines on sheet |
| [ProtectAllSheets](#protect-all-sheets) | Password protect all sheets |
| [UnhideAllSheets](#unhide-all-sheets) | Show all hidden sheets |
| [GoToLastCell](#go-to-last-cell) | Navigate to end of data |
| [HighlightNegatives](#highlight-negative-values) | Color negative numbers |
| [CleanWorkbook](#clean-workbook) | Remove empty sheets, fix issues |

---

## Format as Accounting

Apply accounting number format to selection.

```vba
Sub FormatAsAccounting()
    '================================================
    ' Format Selection as Accounting (with $ and alignment)
    '================================================

    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

End Sub

Sub FormatAsAccountingNoDecimals()
    '================================================
    ' Format as Accounting - No Decimals
    '================================================

    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

End Sub

Sub FormatAsNumber()
    '================================================
    ' Format as Number with Commas
    '================================================

    Selection.NumberFormat = "#,##0.00"

End Sub

Sub FormatAsNumberNoDecimals()
    '================================================
    ' Format as Number - No Decimals
    '================================================

    Selection.NumberFormat = "#,##0"

End Sub
```

---

## Format as Percentage

Apply percentage formatting.

```vba
Sub FormatAsPercentage()
    '================================================
    ' Format Selection as Percentage
    '================================================

    Selection.NumberFormat = "0.0%"

End Sub

Sub FormatAsPercentageNoDecimals()
    '================================================
    ' Format as Percentage - No Decimals
    '================================================

    Selection.NumberFormat = "0%"

End Sub

Sub FormatAsPercentageTwoDecimals()
    '================================================
    ' Format as Percentage - Two Decimals
    '================================================

    Selection.NumberFormat = "0.00%"

End Sub
```

---

## Format Header Row

Style the header row professionally.

```vba
Sub FormatHeaderRow()
    '================================================
    ' Format Selection as Professional Header
    '================================================

    With Selection
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)  ' Dark blue
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With

End Sub

Sub FormatHeaderGreen()
    '================================================
    ' Format Header - Green Theme
    '================================================

    With Selection
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 128, 0)  ' Green
        .HorizontalAlignment = xlCenter
    End With

End Sub

Sub FormatHeaderGray()
    '================================================
    ' Format Header - Gray Theme
    '================================================

    With Selection
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)  ' Light gray
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

End Sub

Sub FormatSubtotalRow()
    '================================================
    ' Format Selection as Subtotal Row
    '================================================

    With Selection
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With

End Sub
```

---

## Alternating Row Colors

Add zebra striping for readability.

```vba
Sub AlternatingRowColors()
    '================================================
    ' Add Alternating Row Colors (Zebra Stripes)
    '================================================

    Dim rng As Range
    Dim row As Range
    Dim i As Long

    Set rng = Selection

    i = 0
    For Each row In rng.Rows
        If i Mod 2 = 1 Then
            row.Interior.Color = RGB(242, 242, 242)  ' Light gray
        Else
            row.Interior.ColorIndex = xlNone
        End If
        i = i + 1
    Next row

End Sub

Sub AlternatingBlueRows()
    '================================================
    ' Alternating Rows - Blue Theme
    '================================================

    Dim rng As Range
    Dim row As Range
    Dim i As Long

    Set rng = Selection

    i = 0
    For Each row In rng.Rows
        If i Mod 2 = 1 Then
            row.Interior.Color = RGB(221, 235, 247)  ' Light blue
        Else
            row.Interior.ColorIndex = xlNone
        End If
        i = i + 1
    Next row

End Sub

Sub RemoveAlternatingColors()
    '================================================
    ' Remove Row Colors
    '================================================

    Selection.Interior.ColorIndex = xlNone

End Sub
```

---

## AutoFit All Columns

Auto-size all columns to fit content.

```vba
Sub AutoFitAllColumns()
    '================================================
    ' AutoFit All Columns in Active Sheet
    '================================================

    Cells.EntireColumn.AutoFit

End Sub

Sub AutoFitSelection()
    '================================================
    ' AutoFit Selected Columns
    '================================================

    Selection.EntireColumn.AutoFit

End Sub

Sub SetColumnWidths()
    '================================================
    ' Set Standard Column Widths
    '================================================

    Dim width As Double

    width = Application.InputBox("Enter column width:", "Set Width", 15, Type:=1)
    If width > 0 Then
        Selection.EntireColumn.ColumnWidth = width
    End If

End Sub

Sub AutoFitRowHeight()
    '================================================
    ' AutoFit All Row Heights
    '================================================

    Cells.EntireRow.AutoFit

End Sub
```

---

## Remove Gridlines

Hide gridlines for cleaner look.

```vba
Sub RemoveGridlines()
    '================================================
    ' Hide Gridlines on Active Sheet
    '================================================

    ActiveWindow.DisplayGridlines = False

End Sub

Sub ShowGridlines()
    '================================================
    ' Show Gridlines on Active Sheet
    '================================================

    ActiveWindow.DisplayGridlines = True

End Sub

Sub ToggleGridlines()
    '================================================
    ' Toggle Gridlines On/Off
    '================================================

    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines

End Sub

Sub RemoveGridlinesAllSheets()
    '================================================
    ' Hide Gridlines on ALL Sheets
    '================================================

    Dim ws As Worksheet
    Dim currentSheet As Worksheet

    Set currentSheet = ActiveSheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ActiveWindow.DisplayGridlines = False
    Next ws

    currentSheet.Activate

    MsgBox "Gridlines removed from all sheets.", vbInformation

End Sub
```

---

## Protect All Sheets

Password protect all worksheets.

```vba
Sub ProtectAllSheets()
    '================================================
    ' Protect All Worksheets
    '================================================

    Dim ws As Worksheet
    Dim pwd As String
    Dim sheetCount As Long

    pwd = InputBox("Enter password (leave blank for no password):", "Protect Sheets")

    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
        sheetCount = sheetCount + 1
    Next ws

    MsgBox "Protected " & sheetCount & " sheets.", vbInformation

End Sub

Sub UnprotectAllSheets()
    '================================================
    ' Unprotect All Worksheets
    '================================================

    Dim ws As Worksheet
    Dim pwd As String
    Dim sheetCount As Long

    pwd = InputBox("Enter password:", "Unprotect Sheets")

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=pwd
        If Err.Number = 0 Then sheetCount = sheetCount + 1
        Err.Clear
    Next ws
    On Error GoTo 0

    MsgBox "Unprotected " & sheetCount & " sheets.", vbInformation

End Sub

Sub ProtectWorkbook()
    '================================================
    ' Protect Workbook Structure
    '================================================

    Dim pwd As String

    pwd = InputBox("Enter password:", "Protect Workbook Structure")

    ThisWorkbook.Protect Password:=pwd, Structure:=True, Windows:=False

    MsgBox "Workbook structure protected.", vbInformation

End Sub
```

---

## Unhide All Sheets

Show all hidden worksheets.

```vba
Sub UnhideAllSheets()
    '================================================
    ' Unhide All Hidden Sheets
    '================================================

    Dim ws As Worksheet
    Dim unhideCount As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            unhideCount = unhideCount + 1
        End If
    Next ws

    MsgBox "Unhid " & unhideCount & " sheets.", vbInformation

End Sub

Sub HideSheet()
    '================================================
    ' Hide Active Sheet
    '================================================

    If ThisWorkbook.Worksheets.Count > 1 Then
        ActiveSheet.Visible = xlSheetHidden
    Else
        MsgBox "Cannot hide the only sheet in workbook.", vbExclamation
    End If

End Sub

Sub VeryHideSheet()
    '================================================
    ' Very Hide Sheet (only visible via VBA)
    '================================================

    ActiveSheet.Visible = xlSheetVeryHidden
    MsgBox "Sheet is now very hidden. Use VBA to unhide.", vbInformation

End Sub

Sub ListHiddenSheets()
    '================================================
    ' List All Hidden Sheets
    '================================================

    Dim ws As Worksheet
    Dim hiddenList As String

    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetHidden Then
            hiddenList = hiddenList & ws.Name & " (Hidden)" & vbCrLf
        ElseIf ws.Visible = xlSheetVeryHidden Then
            hiddenList = hiddenList & ws.Name & " (Very Hidden)" & vbCrLf
        End If
    Next ws

    If hiddenList = "" Then
        MsgBox "No hidden sheets found.", vbInformation
    Else
        MsgBox "Hidden Sheets:" & vbCrLf & vbCrLf & hiddenList, vbInformation
    End If

End Sub
```

---

## Go to Last Cell

Navigate to end of data.

```vba
Sub GoToLastCell()
    '================================================
    ' Go to Last Used Cell
    '================================================

    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Cells(lastRow, lastCol).Select

End Sub

Sub GoToLastRow()
    '================================================
    ' Go to Last Row in Column A
    '================================================

    Cells(Rows.Count, "A").End(xlUp).Select

End Sub

Sub GoToLastColumn()
    '================================================
    ' Go to Last Column in Row 1
    '================================================

    Cells(1, Columns.Count).End(xlToLeft).Select

End Sub

Sub GoToNamedRange()
    '================================================
    ' Go to Named Range
    '================================================

    Dim rangeName As String

    rangeName = InputBox("Enter range name:", "Go To Range")
    If rangeName = "" Then Exit Sub

    On Error Resume Next
    Application.Goto Reference:=Range(rangeName)
    If Err.Number <> 0 Then
        MsgBox "Range '" & rangeName & "' not found.", vbExclamation
    End If
    On Error GoTo 0

End Sub
```

---

## Highlight Negative Values

Color negative numbers red.

```vba
Sub HighlightNegativeValues()
    '================================================
    ' Highlight Negative Numbers in Red
    '================================================

    Dim rng As Range
    Dim cell As Range

    Set rng = Selection

    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value < 0 Then
                cell.Font.Color = RGB(192, 0, 0)  ' Dark red
            Else
                cell.Font.ColorIndex = xlAutomatic
            End If
        End If
    Next cell

End Sub

Sub ConditionalFormatNegatives()
    '================================================
    ' Apply Conditional Formatting for Negatives
    '================================================

    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    Selection.FormatConditions(1).Font.Color = RGB(192, 0, 0)

    MsgBox "Conditional formatting applied for negative values.", vbInformation

End Sub

Sub HighlightZeros()
    '================================================
    ' Highlight Zero Values
    '================================================

    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
    Selection.FormatConditions(1).Interior.Color = RGB(255, 235, 156)  ' Yellow

End Sub

Sub HighlightBlanks()
    '================================================
    ' Highlight Blank Cells
    '================================================

    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlBlanksCondition
    Selection.FormatConditions(1).Interior.Color = RGB(255, 199, 206)  ' Red

End Sub
```

---

## Clean Workbook

Remove issues and optimize workbook.

```vba
Sub CleanWorkbook()
    '================================================
    ' Clean Workbook - Remove Issues
    '================================================

    Dim ws As Worksheet
    Dim response As VbMsgBoxResult
    Dim cleanedItems As String

    response = MsgBox("This will:" & vbCrLf & _
                     "- Remove empty sheets" & vbCrLf & _
                     "- Delete unused cells" & vbCrLf & _
                     "- Clear formatting outside data" & vbCrLf & vbCrLf & _
                     "Continue?", vbYesNo + vbQuestion, "Clean Workbook")

    If response = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ' Remove empty sheets
    For Each ws In ThisWorkbook.Worksheets
        If ThisWorkbook.Worksheets.Count > 1 Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
                cleanedItems = cleanedItems & "Deleted empty sheet" & vbCrLf
            End If
        End If
    Next ws

    ' Reset used range on each remaining sheet
    For Each ws In ThisWorkbook.Worksheets
        Dim lastRow As Long, lastCol As Long

        On Error Resume Next
        lastRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        lastCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        On Error GoTo 0

        If lastRow > 0 And lastCol > 0 Then
            ' Clear formatting outside used range
            If lastRow < ws.Rows.Count Then
                ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(ws.Rows.Count, lastCol)).Clear
            End If
            If lastCol < ws.Columns.Count Then
                ws.Range(ws.Cells(1, lastCol + 1), ws.Cells(lastRow, ws.Columns.Count)).Clear
            End If
        End If
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Workbook cleaned!" & vbCrLf & vbCrLf & cleanedItems, vbInformation

End Sub

Sub CompactWorkbook()
    '================================================
    ' Compact Workbook - Save to Reduce Size
    '================================================

    Dim originalSize As Long
    Dim newSize As Long
    Dim filePath As String

    filePath = ThisWorkbook.FullName

    originalSize = FileLen(filePath)

    ' Save workbook (this often reduces file size)
    ThisWorkbook.Save

    newSize = FileLen(filePath)

    MsgBox "File Size:" & vbCrLf & _
           "Before: " & Format(originalSize / 1024, "#,##0") & " KB" & vbCrLf & _
           "After: " & Format(newSize / 1024, "#,##0") & " KB" & vbCrLf & _
           "Saved: " & Format((originalSize - newSize) / 1024, "#,##0") & " KB", vbInformation

End Sub
```

---

## Additional Utilities

```vba
Sub AddBorders()
    '================================================
    ' Add Borders to Selection
    '================================================

    With Selection.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

End Sub

Sub AddOutlineBorder()
    '================================================
    ' Add Outline Border Only
    '================================================

    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium

End Sub

Sub RemoveBorders()
    '================================================
    ' Remove All Borders
    '================================================

    Selection.Borders.LineStyle = xlNone

End Sub

Sub MergeCentered()
    '================================================
    ' Merge and Center Selection
    '================================================

    With Selection
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

End Sub

Sub UnmergeCells()
    '================================================
    ' Unmerge Selected Cells
    '================================================

    Selection.UnMerge

End Sub

Sub FreezeTopRow()
    '================================================
    ' Freeze Top Row
    '================================================

    ActiveWindow.FreezePanes = False
    Rows(2).Select
    ActiveWindow.FreezePanes = True
    Range("A1").Select

End Sub

Sub FreezeFirstColumn()
    '================================================
    ' Freeze First Column
    '================================================

    ActiveWindow.FreezePanes = False
    Columns("B").Select
    ActiveWindow.FreezePanes = True
    Range("A1").Select

End Sub

Sub UnfreezePanes()
    '================================================
    ' Unfreeze All Panes
    '================================================

    ActiveWindow.FreezePanes = False

End Sub

Sub InsertTimestamp()
    '================================================
    ' Insert Current Date/Time
    '================================================

    Selection.Value = Now
    Selection.NumberFormat = "mm/dd/yyyy hh:mm AM/PM"

End Sub

Sub InsertUsername()
    '================================================
    ' Insert Current Windows Username
    '================================================

    Selection.Value = Environ("USERNAME")

End Sub
```

---

## Quick Format Codes

| Format | Code |
|--------|------|
| Accounting | `"_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"` |
| Number with commas | `"#,##0.00"` |
| Percentage | `"0.0%"` |
| Date | `"mm/dd/yyyy"` |
| Date/Time | `"mm/dd/yyyy hh:mm"` |
| Text | `"@"` |

---

## Color Reference (RGB)

| Color | RGB Code |
|-------|----------|
| Dark Blue | `RGB(0, 51, 102)` |
| Light Blue | `RGB(221, 235, 247)` |
| Green | `RGB(198, 239, 206)` |
| Yellow | `RGB(255, 235, 156)` |
| Red | `RGB(255, 199, 206)` |
| Gray | `RGB(217, 217, 217)` |
| White | `RGB(255, 255, 255)` |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
