# Data Import & Cleanup VBA Macros

> **There's a VBA for That!** - Import, clean, and transform data with one click

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [ImportCSVFile](#import-csv-file) | Import a CSV file into a new sheet |
| [ImportMultipleFiles](#import-multiple-files) | Combine multiple files into one sheet |
| [RemoveDuplicates](#remove-duplicates) | Delete duplicate rows |
| [TrimAllCells](#trim-all-cells) | Remove extra spaces from all cells |
| [CleanText](#clean-all-text) | Remove non-printable characters |
| [SplitTextToColumns](#split-text-to-columns) | Parse delimited text into columns |
| [CombineColumns](#combine-columns) | Merge text from multiple columns |
| [RemoveBlankRows](#remove-blank-rows) | Delete all empty rows |
| [ConvertTextToNumbers](#convert-text-to-numbers) | Fix "numbers stored as text" |
| [StandardizeNames](#standardize-names) | Proper case for names |
| [ExtractNumbers](#extract-numbers-from-text) | Pull numbers from mixed text |
| [FindAndReplaceAll](#find-and-replace-all-sheets) | Find/replace across all sheets |

---

## Import CSV File

Import a single CSV file into a new worksheet.

```vba
Sub ImportCSVFile()
    '================================================
    ' Import CSV File to New Sheet
    ' Usage: Run macro, select CSV file, imports to new sheet
    '================================================

    Dim fd As FileDialog
    Dim filePath As String
    Dim wsNew As Worksheet
    Dim wb As Workbook

    ' Create file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select CSV File to Import"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "Text Files", "*.txt"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False

        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub  ' User cancelled
        End If
    End With

    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Open the CSV file
    Set wb = Workbooks.Open(filePath)

    ' Copy to this workbook
    wb.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    ' Close the CSV without saving
    wb.Close SaveChanges:=False

    ' Rename the new sheet
    wsNew.Name = Left(Replace(Dir(filePath), ".csv", ""), 31)

    Application.ScreenUpdating = True

    MsgBox "Imported: " & filePath, vbInformation

End Sub
```

---

## Import Multiple Files

Combine multiple CSV/Excel files into one master sheet.

```vba
Sub ImportMultipleFiles()
    '================================================
    ' Import Multiple Files into One Sheet
    ' Usage: Select multiple files, combines all into one sheet
    '================================================

    Dim fd As FileDialog
    Dim filePath As Variant
    Dim wsMaster As Worksheet
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim lastRowMaster As Long
    Dim lastRowSource As Long
    Dim lastColSource As Long
    Dim fileCount As Long
    Dim copyHeader As Boolean

    ' Create master sheet
    Set wsMaster = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsMaster.Name = "Combined_Data"

    ' File dialog for multiple selection
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Files to Combine"
        .Filters.Clear
        .Filters.Add "Excel & CSV", "*.xlsx;*.xls;*.csv"
        .AllowMultiSelect = True

        If .Show = -1 Then
            ' Process each selected file
            Application.ScreenUpdating = False
            copyHeader = True

            For Each filePath In .SelectedItems
                fileCount = fileCount + 1

                ' Open source file
                Set wbSource = Workbooks.Open(filePath)
                Set wsSource = wbSource.Sheets(1)

                ' Get dimensions
                lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
                lastColSource = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
                lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row

                If copyHeader Then
                    ' Copy including header (first file only)
                    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRowSource, lastColSource)).Copy _
                        Destination:=wsMaster.Cells(1, 1)
                    copyHeader = False
                Else
                    ' Copy without header (subsequent files)
                    wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRowSource, lastColSource)).Copy _
                        Destination:=wsMaster.Cells(lastRowMaster + 1, 1)
                End If

                ' Close source file
                wbSource.Close SaveChanges:=False
            Next filePath

            Application.ScreenUpdating = True
            MsgBox fileCount & " files combined successfully!", vbInformation
        End If
    End With

End Sub
```

---

## Remove Duplicates

Delete duplicate rows based on all columns or specific columns.

```vba
Sub RemoveDuplicates()
    '================================================
    ' Remove Duplicate Rows
    ' Usage: Select data range first, then run
    '================================================

    Dim rng As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim originalCount As Long
    Dim newCount As Long

    ' Get the used range
    On Error Resume Next
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "Please select a range first.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Count original rows
    originalCount = rng.Rows.Count

    ' Remove duplicates based on all columns
    rng.RemoveDuplicates Columns:=Array(1), Header:=xlYes

    ' Count new rows
    newCount = ActiveSheet.UsedRange.Rows.Count

    MsgBox "Removed " & (originalCount - newCount) & " duplicate rows.", vbInformation

End Sub

Sub RemoveDuplicatesAllColumns()
    '================================================
    ' Remove Duplicates Based on ALL Columns
    '================================================

    Dim rng As Range
    Dim lastCol As Long
    Dim colArray() As Long
    Dim i As Long

    Set rng = ActiveSheet.UsedRange
    lastCol = rng.Columns.Count

    ' Build array of all column numbers
    ReDim colArray(1 To lastCol)
    For i = 1 To lastCol
        colArray(i) = i
    Next i

    rng.RemoveDuplicates Columns:=colArray, Header:=xlYes

    MsgBox "Duplicates removed!", vbInformation

End Sub
```

---

## Trim All Cells

Remove leading, trailing, and extra spaces from all cells.

```vba
Sub TrimAllCells()
    '================================================
    ' Trim All Cells in Selection or Used Range
    ' Removes leading, trailing, and double spaces
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim trimCount As Long

    Application.ScreenUpdating = False

    ' Use selection or entire used range
    If Selection.Cells.Count > 1 Then
        Set rng = Selection
    Else
        Set rng = ActiveSheet.UsedRange
    End If

    For Each cell In rng
        If Not IsEmpty(cell) And Not IsNumeric(cell.Value) Then
            If cell.Value <> Application.WorksheetFunction.Trim(cell.Value) Then
                cell.Value = Application.WorksheetFunction.Trim(cell.Value)
                trimCount = trimCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Trimmed " & trimCount & " cells.", vbInformation

End Sub
```

---

## Clean All Text

Remove non-printable characters from cells.

```vba
Sub CleanAllText()
    '================================================
    ' Clean Non-Printable Characters
    ' Removes CHAR(0) through CHAR(31) except tab, CR, LF
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim cleanCount As Long
    Dim oldVal As String
    Dim newVal As String

    Application.ScreenUpdating = False

    Set rng = ActiveSheet.UsedRange

    For Each cell In rng
        If Not IsEmpty(cell) And Not IsNumeric(cell.Value) Then
            oldVal = cell.Value
            newVal = Application.WorksheetFunction.Clean(oldVal)

            ' Also trim
            newVal = Application.WorksheetFunction.Trim(newVal)

            If oldVal <> newVal Then
                cell.Value = newVal
                cleanCount = cleanCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Cleaned " & cleanCount & " cells.", vbInformation

End Sub
```

---

## Split Text to Columns

Parse delimited text into multiple columns.

```vba
Sub SplitTextToColumns()
    '================================================
    ' Split Text to Columns
    ' Usage: Select the column to split
    '================================================

    Dim rng As Range
    Dim delimiter As String

    Set rng = Selection

    ' Ask for delimiter
    delimiter = InputBox("Enter delimiter:" & vbCrLf & _
                         "(comma, semicolon, tab, space, pipe)", _
                         "Text to Columns", ",")

    If delimiter = "" Then Exit Sub

    ' Convert delimiter text to actual character
    Select Case LCase(delimiter)
        Case "tab"
            delimiter = vbTab
        Case "space"
            delimiter = " "
        Case "pipe"
            delimiter = "|"
    End Select

    ' Text to columns
    rng.TextToColumns _
        Destination:=rng, _
        DataType:=xlDelimited, _
        Other:=True, _
        OtherChar:=delimiter

    MsgBox "Split complete!", vbInformation

End Sub
```

---

## Combine Columns

Merge text from multiple columns into one.

```vba
Sub CombineColumns()
    '================================================
    ' Combine Multiple Columns into One
    ' Usage: Select columns to combine, result goes in next column
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim row As Long
    Dim col As Long
    Dim result As String
    Dim delimiter As String
    Dim destCol As Long

    Set rng = Selection
    destCol = rng.Column + rng.Columns.Count

    ' Ask for delimiter
    delimiter = InputBox("Enter separator between values:", "Combine Columns", " ")

    Application.ScreenUpdating = False

    For row = 1 To rng.Rows.Count
        result = ""
        For col = 1 To rng.Columns.Count
            If rng.Cells(row, col).Value <> "" Then
                If result <> "" Then result = result & delimiter
                result = result & rng.Cells(row, col).Value
            End If
        Next col
        Cells(rng.row + row - 1, destCol).Value = result
    Next row

    Application.ScreenUpdating = True

    MsgBox "Combined into column " & destCol, vbInformation

End Sub
```

---

## Remove Blank Rows

Delete all empty rows in the data range.

```vba
Sub RemoveBlankRows()
    '================================================
    ' Remove All Blank Rows
    '================================================

    Dim rng As Range
    Dim i As Long
    Dim lastRow As Long
    Dim deleteCount As Long

    Application.ScreenUpdating = False

    lastRow = ActiveSheet.UsedRange.Rows.Count

    ' Loop backwards to avoid skipping rows
    For i = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(Rows(i)) = 0 Then
            Rows(i).Delete
            deleteCount = deleteCount + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Deleted " & deleteCount & " blank rows.", vbInformation

End Sub

Sub RemoveRowsWithBlankInColumnA()
    '================================================
    ' Remove Rows Where Column A is Blank
    '================================================

    Dim i As Long
    Dim lastRow As Long
    Dim deleteCount As Long

    Application.ScreenUpdating = False

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 2 Step -1  ' Start at 2 to skip header
        If IsEmpty(Cells(i, 1)) Or Cells(i, 1).Value = "" Then
            Rows(i).Delete
            deleteCount = deleteCount + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Deleted " & deleteCount & " rows.", vbInformation

End Sub
```

---

## Convert Text to Numbers

Fix numbers that are stored as text (green triangle warning).

```vba
Sub ConvertTextToNumbers()
    '================================================
    ' Convert Text to Numbers
    ' Fixes "Number Stored as Text" errors
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim convertCount As Long

    Application.ScreenUpdating = False

    ' Use selection or used range
    If Selection.Cells.Count > 1 Then
        Set rng = Selection
    Else
        Set rng = ActiveSheet.UsedRange
    End If

    For Each cell In rng
        If IsNumeric(cell.Value) And Not IsEmpty(cell) Then
            cell.Value = cell.Value * 1  ' Multiply by 1 converts text to number
            convertCount = convertCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Converted " & convertCount & " cells to numbers.", vbInformation

End Sub

Sub ConvertTextToNumbersAlt()
    '================================================
    ' Alternative: Paste Special Multiply Method
    ' Sometimes more reliable for stubborn text-numbers
    '================================================

    Dim rng As Range
    Dim tempCell As Range

    Set rng = Selection

    ' Put 1 in a temporary cell
    Set tempCell = ActiveSheet.Cells(1, ActiveSheet.UsedRange.Columns.Count + 2)
    tempCell.Value = 1

    ' Copy and paste special multiply
    tempCell.Copy
    rng.PasteSpecial Paste:=xlPasteValues, Operation:=xlMultiply
    Application.CutCopyMode = False

    ' Clear temp cell
    tempCell.Clear

    MsgBox "Conversion complete!", vbInformation

End Sub
```

---

## Standardize Names

Convert names to proper case (Title Case).

```vba
Sub StandardizeNames()
    '================================================
    ' Standardize Names to Proper Case
    ' "JOHN SMITH" or "john smith" becomes "John Smith"
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim fixCount As Long

    Application.ScreenUpdating = False

    Set rng = Selection

    For Each cell In rng
        If Not IsEmpty(cell) And Not IsNumeric(cell.Value) Then
            cell.Value = Application.WorksheetFunction.Proper(cell.Value)
            fixCount = fixCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Standardized " & fixCount & " names.", vbInformation

End Sub

Sub StandardizeNamesAdvanced()
    '================================================
    ' Advanced Name Standardization
    ' Handles LLC, Inc, II, III, etc.
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim val As String

    Application.ScreenUpdating = False

    Set rng = Selection

    For Each cell In rng
        If Not IsEmpty(cell) And Not IsNumeric(cell.Value) Then
            val = Application.WorksheetFunction.Proper(cell.Value)

            ' Fix common abbreviations
            val = Replace(val, " Llc", " LLC")
            val = Replace(val, " Llp", " LLP")
            val = Replace(val, " Inc", " Inc.")
            val = Replace(val, " Corp", " Corp.")
            val = Replace(val, " Ii", " II")
            val = Replace(val, " Iii", " III")
            val = Replace(val, " Iv", " IV")
            val = Replace(val, "Mc", "Mc")  ' McDonald stays McDonald
            val = Replace(val, " Po ", " PO ")
            val = Replace(val, " Cpa", " CPA")

            cell.Value = val
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Names standardized!", vbInformation

End Sub
```

---

## Extract Numbers from Text

Pull numeric values from cells containing mixed text and numbers.

```vba
Sub ExtractNumbersFromText()
    '================================================
    ' Extract Numbers from Text
    ' "Invoice #12345" becomes "12345"
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim destCol As Long

    Set rng = Selection
    destCol = rng.Column + 1

    Application.ScreenUpdating = False

    For Each cell In rng
        result = ""
        For i = 1 To Len(cell.Value)
            char = Mid(cell.Value, i, 1)
            If char Like "[0-9]" Or char = "." Or char = "-" Then
                result = result & char
            End If
        Next i

        ' Put result in next column
        Cells(cell.Row, destCol).Value = result
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Numbers extracted to column " & destCol, vbInformation

End Sub

Sub ExtractTextOnly()
    '================================================
    ' Extract Text Only (Remove Numbers)
    '================================================

    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim destCol As Long

    Set rng = Selection
    destCol = rng.Column + 1

    Application.ScreenUpdating = False

    For Each cell In rng
        result = ""
        For i = 1 To Len(cell.Value)
            char = Mid(cell.Value, i, 1)
            If Not (char Like "[0-9]") Then
                result = result & char
            End If
        Next i

        Cells(cell.Row, destCol).Value = Trim(result)
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Text extracted to column " & destCol, vbInformation

End Sub
```

---

## Find and Replace All Sheets

Find and replace text across all worksheets in workbook.

```vba
Sub FindAndReplaceAllSheets()
    '================================================
    ' Find and Replace Across ALL Worksheets
    '================================================

    Dim ws As Worksheet
    Dim findText As String
    Dim replaceText As String
    Dim replaceCount As Long

    findText = InputBox("Find what:", "Find and Replace")
    If findText = "" Then Exit Sub

    replaceText = InputBox("Replace with:", "Find and Replace", "")

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Replace What:=findText, _
                        Replacement:=replaceText, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        MatchCase:=False
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Find and replace complete across all sheets!", vbInformation

End Sub
```

---

## Bonus: Import from Clipboard

Paste and parse clipboard data that's tab or comma delimited.

```vba
Sub ImportFromClipboard()
    '================================================
    ' Import and Parse Clipboard Data
    '================================================

    Dim ws As Worksheet
    Dim pasteRange As Range

    Set ws = ActiveSheet
    Set pasteRange = Selection

    ' Paste from clipboard
    pasteRange.Select
    ws.Paste

    ' Offer to split if single column
    If MsgBox("Split data into columns?", vbYesNo, "Text to Columns") = vbYes Then
        Selection.TextToColumns Destination:=pasteRange, _
            DataType:=xlDelimited, _
            Tab:=True, _
            Comma:=True

    End If

End Sub
```

---

## Best Practices for Data Import

| Practice | Why |
|----------|-----|
| **Always backup first** | Run macro on a copy, not the original |
| **Check row counts** | Verify before/after counts match expectations |
| **Validate data types** | Numbers are numbers, dates are dates |
| **Review for blanks** | Unexpected blanks can break formulas |
| **Document source** | Note where data came from |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
