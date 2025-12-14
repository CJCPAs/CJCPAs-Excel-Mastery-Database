# Workpaper VBA Macros

> **There's a VBA for That!** - Create, format, and organize workpapers like a pro

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [CreateWorkpaper](#create-workpaper-template) | Generate standard workpaper template |
| [AddWorkpaperHeader](#add-workpaper-header) | Insert professional header |
| [NumberWorkpapers](#auto-number-workpapers) | Auto-number workpaper sheets |
| [CreatePBCList](#create-pbc-list) | Generate Prepared By Client list |
| [AddTickmarkLegend](#add-tickmark-legend) | Insert standard tickmarks |
| [InsertTickmark](#insert-tickmark) | Add tickmark to cell |
| [AddSignOff](#add-sign-off-section) | Preparer/Reviewer sign-off |
| [CrossReferenceCell](#cross-reference) | Link to other workpapers |
| [CreateIndexSheet](#create-index-sheet) | Master workpaper index |
| [ProtectWorkpaper](#protect-workpaper) | Lock completed workpapers |

---

## Create Workpaper Template

Generate a standard audit/tax workpaper template.

```vba
Sub CreateWorkpaperTemplate()
    '================================================
    ' Create Standard Workpaper Template
    '================================================

    Dim ws As Worksheet
    Dim wpRef As String
    Dim wpDesc As String
    Dim clientName As String

    wpRef = InputBox("Enter workpaper reference (e.g., A-1):", "New Workpaper", "A-1")
    If wpRef = "" Then Exit Sub

    wpDesc = InputBox("Enter workpaper description:", "New Workpaper", "Cash Lead Schedule")
    If wpDesc = "" Then Exit Sub

    clientName = InputBox("Enter client name:", "New Workpaper", "Client Name")

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    On Error Resume Next
    ws.Name = Left(wpRef, 31)
    On Error GoTo 0

    Application.ScreenUpdating = False

    With ws
        ' Header Row 1 - Client and WP Ref
        .Range("A1").Value = clientName
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12

        .Range("H1").Value = "W/P Ref: " & wpRef
        .Range("H1").Font.Bold = True
        .Range("H1").HorizontalAlignment = xlRight

        ' Header Row 2 - Description
        .Range("A2").Value = wpDesc
        .Range("A2").Font.Bold = True
        .Range("A2").Font.Size = 11

        ' Header Row 3 - Period
        .Range("A3").Value = "Period Ending: " & Format(Application.WorksheetFunction.EoMonth(Date, 0), "mmmm d, yyyy")

        ' Header border
        .Range("A1:H3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A1:H3").Borders(xlEdgeBottom).Weight = xlMedium

        ' Objective section
        .Range("A5").Value = "OBJECTIVE:"
        .Range("A5").Font.Bold = True
        .Range("A6").Value = "[State the purpose of this workpaper]"
        .Range("A6").Font.Italic = True

        ' Procedures section
        .Range("A8").Value = "PROCEDURES PERFORMED:"
        .Range("A8").Font.Bold = True
        .Range("A9").Value = "1. "
        .Range("A10").Value = "2. "
        .Range("A11").Value = "3. "

        ' Work area
        .Range("A13").Value = "ANALYSIS:"
        .Range("A13").Font.Bold = True

        ' Conclusion section (bottom of standard workpaper area)
        .Range("A40").Value = "CONCLUSION:"
        .Range("A40").Font.Bold = True
        .Range("A41").Value = "[State conclusion based on work performed]"
        .Range("A41").Font.Italic = True

        ' Sign-off section
        .Range("A43").Value = "Prepared By:"
        .Range("B43").Value = Environ("USERNAME")
        .Range("C43").Value = "Date:"
        .Range("D43").Value = Date

        .Range("A44").Value = "Reviewed By:"
        .Range("C44").Value = "Date:"

        ' Tickmark legend area
        .Range("F43").Value = "TICKMARKS"
        .Range("F43").Font.Bold = True
        .Range("F44").Value = ChrW(10003) & " = Agreed to source"
        .Range("F45").Value = "F = Footed"
        .Range("F46").Value = "T = Traced"

        ' Column widths
        .Columns("A").ColumnWidth = 40
        .Columns("B:G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15

        ' Page setup
        .PageSetup.PrintTitleRows = "$1:$4"
        .PageSetup.CenterHorizontally = True

    End With

    Application.ScreenUpdating = True

    MsgBox "Workpaper " & wpRef & " created!", vbInformation

End Sub
```

---

## Add Workpaper Header

Add professional header to any worksheet.

```vba
Sub AddWorkpaperHeader()
    '================================================
    ' Add Standard Workpaper Header
    '================================================

    Dim ws As Worksheet
    Dim clientName As String
    Dim wpRef As String
    Dim wpDesc As String

    Set ws = ActiveSheet

    clientName = InputBox("Client Name:", "Workpaper Header", "Client Name")
    wpRef = InputBox("W/P Reference:", "Workpaper Header", "A-1")
    wpDesc = InputBox("Description:", "Workpaper Header", "Schedule Description")

    ' Insert 4 rows at top
    ws.Rows("1:4").Insert Shift:=xlDown

    With ws
        .Range("A1").Value = clientName
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 12

        .Range("H1").Value = "W/P: " & wpRef
        .Range("H1").Font.Bold = True

        .Range("A2").Value = wpDesc
        .Range("A2").Font.Bold = True

        .Range("A3").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "12/31/yyyy")

        ' Bottom border
        .Range("A3:H3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A3:H3").Borders(xlEdgeBottom).Weight = xlMedium
    End With

    MsgBox "Header added!", vbInformation

End Sub
```

---

## Auto-Number Workpapers

Automatically number/rename workpaper sheets.

```vba
Sub NumberWorkpapers()
    '================================================
    ' Auto-Number Workpaper Sheets
    '================================================

    Dim ws As Worksheet
    Dim prefix As String
    Dim counter As Integer
    Dim i As Integer

    prefix = InputBox("Enter workpaper prefix (e.g., A, B, C):", "Number Workpapers", "A")
    If prefix = "" Then Exit Sub

    counter = 1

    For i = 1 To ThisWorkbook.Sheets.Count
        Set ws = ThisWorkbook.Sheets(i)

        ' Skip index and summary sheets
        If LCase(ws.Name) <> "index" And LCase(ws.Name) <> "summary" Then
            On Error Resume Next
            ws.Name = prefix & "-" & counter
            On Error GoTo 0

            ' Also update W/P ref cell if exists
            If ws.Range("H1").Value Like "W/P*" Then
                ws.Range("H1").Value = "W/P: " & prefix & "-" & counter
            End If

            counter = counter + 1
        End If
    Next i

    MsgBox "Numbered " & (counter - 1) & " workpapers.", vbInformation

End Sub
```

---

## Create PBC List

Generate Prepared By Client request list.

```vba
Sub CreatePBCList()
    '================================================
    ' Create Prepared By Client (PBC) List
    '================================================

    Dim ws As Worksheet
    Dim clientName As String
    Dim periodEnd As String

    clientName = InputBox("Client Name:", "PBC List", "Client Name")
    periodEnd = InputBox("Period End (mm/dd/yyyy):", "PBC List", Format(DateSerial(Year(Date), 12, 31), "mm/dd/yyyy"))

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "PBC List"

    Application.ScreenUpdating = False

    With ws
        ' Header
        .Range("A1").Value = "PREPARED BY CLIENT LIST"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = clientName
        .Range("A3").Value = "Period Ending: " & periodEnd

        ' Column Headers
        .Range("A5").Value = "Item #"
        .Range("B5").Value = "Description"
        .Range("C5").Value = "W/P Ref"
        .Range("D5").Value = "Requested"
        .Range("E5").Value = "Due Date"
        .Range("F5").Value = "Received"
        .Range("G5").Value = "Status"
        .Range("H5").Value = "Notes"

        .Range("A5:H5").Font.Bold = True
        .Range("A5:H5").Interior.Color = RGB(0, 51, 102)
        .Range("A5:H5").Font.Color = RGB(255, 255, 255)

        ' Sample items
        .Range("A6").Value = 1
        .Range("B6").Value = "General Ledger / Trial Balance"
        .Range("D6").Value = Date

        .Range("A7").Value = 2
        .Range("B7").Value = "Bank Statements - All Accounts"
        .Range("D7").Value = Date

        .Range("A8").Value = 3
        .Range("B8").Value = "Bank Reconciliations - All Accounts"
        .Range("D8").Value = Date

        .Range("A9").Value = 4
        .Range("B9").Value = "Accounts Receivable Aging"
        .Range("D9").Value = Date

        .Range("A10").Value = 5
        .Range("B10").Value = "Accounts Payable Aging"
        .Range("D10").Value = Date

        .Range("A11").Value = 6
        .Range("B11").Value = "Fixed Asset Schedule"
        .Range("D11").Value = Date

        .Range("A12").Value = 7
        .Range("B12").Value = "Loan Statements"
        .Range("D12").Value = Date

        .Range("A13").Value = 8
        .Range("B13").Value = "Payroll Reports (941s, W-2s)"
        .Range("D13").Value = Date

        .Range("A14").Value = 9
        .Range("B14").Value = "Prior Year Tax Return"
        .Range("D14").Value = Date

        .Range("A15").Value = 10
        .Range("B15").Value = "Corporate Minutes / Resolutions"
        .Range("D15").Value = Date

        ' Status dropdown (Data Validation)
        Dim statusList As String
        statusList = "Open,Received,Partial,N/A"

        .Range("G6:G100").Validation.Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=statusList

        ' Conditional formatting for status
        .Range("G6:G100").FormatConditions.Add Type:=xlTextString, String:="Received", TextOperator:=xlContains
        .Range("G6:G100").FormatConditions(1).Interior.Color = RGB(198, 239, 206)

        .Range("G6:G100").FormatConditions.Add Type:=xlTextString, String:="Open", TextOperator:=xlContains
        .Range("G6:G100").FormatConditions(2).Interior.Color = RGB(255, 199, 206)

        ' Column widths
        .Columns("A").ColumnWidth = 8
        .Columns("B").ColumnWidth = 40
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 30

        ' Date format
        .Columns("D:F").NumberFormat = "mm/dd/yyyy"

        ' Borders
        .Range("A5:H100").Borders.LineStyle = xlContinuous

    End With

    Application.ScreenUpdating = True

    MsgBox "PBC List created!", vbInformation

End Sub
```

---

## Add Tickmark Legend

Insert standard audit tickmark legend.

```vba
Sub AddTickmarkLegend()
    '================================================
    ' Add Tickmark Legend to Workpaper
    '================================================

    Dim ws As Worksheet
    Dim startRow As Long
    Dim startCol As Long

    Set ws = ActiveSheet

    ' Find good location
    startRow = Application.InputBox("Enter starting row:", "Tickmark Legend", ws.UsedRange.Rows.Count + 3, Type:=1)
    startCol = Application.InputBox("Enter starting column:", "Tickmark Legend", 1, Type:=1)

    With ws
        .Cells(startRow, startCol).Value = "TICKMARK LEGEND"
        .Cells(startRow, startCol).Font.Bold = True
        .Cells(startRow, startCol).Font.Underline = xlUnderlineStyleSingle

        .Cells(startRow + 1, startCol).Value = ChrW(10003)
        .Cells(startRow + 1, startCol + 1).Value = "Agreed to source document"

        .Cells(startRow + 2, startCol).Value = "F"
        .Cells(startRow + 2, startCol + 1).Value = "Footed (mathematically verified)"

        .Cells(startRow + 3, startCol).Value = "CF"
        .Cells(startRow + 3, startCol + 1).Value = "Cross-footed"

        .Cells(startRow + 4, startCol).Value = "T"
        .Cells(startRow + 4, startCol + 1).Value = "Traced to supporting detail"

        .Cells(startRow + 5, startCol).Value = "TB"
        .Cells(startRow + 5, startCol + 1).Value = "Agreed to trial balance"

        .Cells(startRow + 6, startCol).Value = "PY"
        .Cells(startRow + 6, startCol + 1).Value = "Agreed to prior year"

        .Cells(startRow + 7, startCol).Value = "GL"
        .Cells(startRow + 7, startCol + 1).Value = "Agreed to general ledger"

        .Cells(startRow + 8, startCol).Value = "I"
        .Cells(startRow + 8, startCol + 1).Value = "Inquired of client"

        .Cells(startRow + 9, startCol).Value = "C"
        .Cells(startRow + 9, startCol + 1).Value = "Confirmed"

        .Cells(startRow + 10, startCol).Value = "R"
        .Cells(startRow + 10, startCol + 1).Value = "Recalculated"

        .Cells(startRow + 11, startCol).Value = "V"
        .Cells(startRow + 11, startCol + 1).Value = "Vouched to invoice/document"

        .Cells(startRow + 12, startCol).Value = "*"
        .Cells(startRow + 12, startCol + 1).Value = "See comment/note below"

        .Cells(startRow + 13, startCol).Value = "N/A"
        .Cells(startRow + 13, startCol + 1).Value = "Not applicable"

        ' Format tickmark column
        .Range(.Cells(startRow + 1, startCol), .Cells(startRow + 13, startCol)).HorizontalAlignment = xlCenter
        .Range(.Cells(startRow + 1, startCol), .Cells(startRow + 13, startCol)).Font.Bold = True
        .Range(.Cells(startRow + 1, startCol), .Cells(startRow + 13, startCol)).Font.Color = RGB(0, 112, 192)

        ' Border around legend
        .Range(.Cells(startRow, startCol), .Cells(startRow + 13, startCol + 1)).Borders.LineStyle = xlContinuous

    End With

    MsgBox "Tickmark legend added!", vbInformation

End Sub
```

---

## Insert Tickmark

Quickly insert a tickmark into the selected cell.

```vba
Sub InsertTickmark()
    '================================================
    ' Insert Tickmark in Selected Cell
    '================================================

    Dim tickmark As String
    Dim choice As Integer

    choice = Application.InputBox("Select tickmark:" & vbCrLf & vbCrLf & _
                                  "1 = " & ChrW(10003) & " (Agreed)" & vbCrLf & _
                                  "2 = F (Footed)" & vbCrLf & _
                                  "3 = T (Traced)" & vbCrLf & _
                                  "4 = TB (Trial Balance)" & vbCrLf & _
                                  "5 = PY (Prior Year)" & vbCrLf & _
                                  "6 = C (Confirmed)" & vbCrLf & _
                                  "7 = R (Recalculated)" & vbCrLf & _
                                  "8 = V (Vouched)" & vbCrLf & _
                                  "9 = * (See note)", _
                                  "Insert Tickmark", 1, Type:=1)

    Select Case choice
        Case 1: tickmark = ChrW(10003)
        Case 2: tickmark = "F"
        Case 3: tickmark = "T"
        Case 4: tickmark = "TB"
        Case 5: tickmark = "PY"
        Case 6: tickmark = "C"
        Case 7: tickmark = "R"
        Case 8: tickmark = "V"
        Case 9: tickmark = "*"
        Case Else: Exit Sub
    End Select

    With Selection
        .Value = tickmark
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = RGB(0, 112, 192)
    End With

End Sub

Sub TickmarkCheckmark()
    Selection.Value = ChrW(10003)
    Selection.Font.Color = RGB(0, 112, 192)
    Selection.HorizontalAlignment = xlCenter
End Sub

Sub TickmarkFooted()
    Selection.Value = "F"
    Selection.Font.Color = RGB(0, 112, 192)
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
End Sub

Sub TickmarkTraced()
    Selection.Value = "T"
    Selection.Font.Color = RGB(0, 112, 192)
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlCenter
End Sub
```

---

## Add Sign-Off Section

Add preparer and reviewer sign-off.

```vba
Sub AddSignOffSection()
    '================================================
    ' Add Sign-Off Section to Workpaper
    '================================================

    Dim ws As Worksheet
    Dim startRow As Long

    Set ws = ActiveSheet
    startRow = ws.UsedRange.Rows.Count + 3

    With ws
        ' Box border
        .Range(.Cells(startRow, 1), .Cells(startRow + 3, 6)).Borders.LineStyle = xlContinuous

        ' Labels
        .Cells(startRow, 1).Value = "Prepared By:"
        .Cells(startRow, 1).Font.Bold = True
        .Cells(startRow + 1, 1).Value = "Reviewed By:"
        .Cells(startRow + 1, 1).Font.Bold = True
        .Cells(startRow + 2, 1).Value = "Partner Review:"
        .Cells(startRow + 2, 1).Font.Bold = True

        .Cells(startRow, 3).Value = "Date:"
        .Cells(startRow + 1, 3).Value = "Date:"
        .Cells(startRow + 2, 3).Value = "Date:"

        ' Auto-fill preparer
        .Cells(startRow, 2).Value = Environ("USERNAME")
        .Cells(startRow, 4).Value = Date

    End With

    MsgBox "Sign-off section added!", vbInformation

End Sub
```

---

## Cross Reference

Create cross-reference link to another workpaper.

```vba
Sub CrossReference()
    '================================================
    ' Create Cross-Reference to Another Workpaper
    '================================================

    Dim xRef As String
    Dim ws As Worksheet

    xRef = InputBox("Enter cross-reference (e.g., A-1, B-2.1):", "Cross Reference", "A-1")
    If xRef = "" Then Exit Sub

    ' Check if referenced sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(xRef)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Sheet doesn't exist, just add text reference
        With Selection
            .Value = "X-Ref: " & xRef
            .Font.Color = RGB(0, 112, 192)
            .Font.Underline = xlUnderlineStyleSingle
        End With
    Else
        ' Create hyperlink to sheet
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, _
            Address:="", _
            SubAddress:="'" & xRef & "'!A1", _
            TextToDisplay:="X-Ref: " & xRef

        Selection.Font.Color = RGB(0, 112, 192)
    End If

End Sub

Sub CreateBackReference()
    '================================================
    ' Create Back-Reference (return link)
    '================================================

    Dim fromSheet As String

    fromSheet = InputBox("Enter source W/P reference:", "Back Reference", "")
    If fromSheet = "" Then Exit Sub

    ' Add return link
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, _
        Address:="", _
        SubAddress:="'" & fromSheet & "'!A1", _
        TextToDisplay:="<< Return to " & fromSheet

    Selection.Font.Color = RGB(128, 128, 128)

End Sub
```

---

## Create Index Sheet

Create master workpaper index.

```vba
Sub CreateIndexSheet()
    '================================================
    ' Create Workpaper Index Sheet
    '================================================

    Dim wsIndex As Worksheet
    Dim ws As Worksheet
    Dim indexRow As Long
    Dim clientName As String

    clientName = InputBox("Enter client name:", "Workpaper Index", "Client Name")

    ' Create or clear index sheet
    On Error Resume Next
    Set wsIndex = ThisWorkbook.Sheets("INDEX")
    If Not wsIndex Is Nothing Then wsIndex.Delete
    On Error GoTo 0

    Set wsIndex = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    wsIndex.Name = "INDEX"

    Application.ScreenUpdating = False

    With wsIndex
        .Range("A1").Value = "WORKPAPER INDEX"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = clientName
        .Range("A3").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "12/31/yyyy")

        ' Headers
        .Range("A5").Value = "W/P Ref"
        .Range("B5").Value = "Description"
        .Range("C5").Value = "Prepared By"
        .Range("D5").Value = "Reviewed By"
        .Range("E5").Value = "Status"

        .Range("A5:E5").Font.Bold = True
        .Range("A5:E5").Interior.Color = RGB(0, 51, 102)
        .Range("A5:E5").Font.Color = RGB(255, 255, 255)

        indexRow = 6

        ' Loop through all sheets
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> "INDEX" Then
                ' W/P Ref with hyperlink
                .Hyperlinks.Add Anchor:=.Cells(indexRow, 1), _
                    Address:="", _
                    SubAddress:="'" & ws.Name & "'!A1", _
                    TextToDisplay:=ws.Name

                ' Try to get description from cell A2
                On Error Resume Next
                .Cells(indexRow, 2).Value = ws.Range("A2").Value
                On Error GoTo 0

                indexRow = indexRow + 1
            End If
        Next ws

        ' Format
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 45
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 12

        ' Borders
        .Range("A5:E" & indexRow - 1).Borders.LineStyle = xlContinuous

    End With

    Application.ScreenUpdating = True

    MsgBox "Index created with " & (indexRow - 6) & " workpapers.", vbInformation

End Sub
```

---

## Protect Workpaper

Lock completed workpapers to prevent changes.

```vba
Sub ProtectWorkpaper()
    '================================================
    ' Protect Completed Workpaper
    '================================================

    Dim ws As Worksheet
    Dim pwd As String
    Dim confirmPwd As String

    Set ws = ActiveSheet

    If MsgBox("Protect workpaper: " & ws.Name & "?" & vbCrLf & vbCrLf & _
              "This will prevent further edits.", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    pwd = InputBox("Enter password (leave blank for no password):", "Protect Workpaper")

    If pwd <> "" Then
        confirmPwd = InputBox("Confirm password:", "Protect Workpaper")
        If pwd <> confirmPwd Then
            MsgBox "Passwords do not match!", vbExclamation
            Exit Sub
        End If
    End If

    ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True

    ' Visual indicator
    ws.Tab.Color = RGB(0, 176, 80)  ' Green tab = finalized

    MsgBox "Workpaper protected!" & vbCrLf & "Green tab indicates finalized status.", vbInformation

End Sub

Sub UnprotectWorkpaper()
    '================================================
    ' Unprotect Workpaper for Edits
    '================================================

    Dim ws As Worksheet
    Dim pwd As String

    Set ws = ActiveSheet

    pwd = InputBox("Enter password:", "Unprotect Workpaper")

    On Error Resume Next
    ws.Unprotect Password:=pwd
    On Error GoTo 0

    If ws.ProtectContents = False Then
        ws.Tab.ColorIndex = xlNone
        MsgBox "Workpaper unprotected.", vbInformation
    Else
        MsgBox "Incorrect password.", vbExclamation
    End If

End Sub
```

---

## Best Practices for Workpapers

| Practice | Description |
|----------|-------------|
| **Clear reference** | Every page needs W/P reference |
| **State objective** | What is this workpaper proving? |
| **Document procedures** | What work did you do? |
| **Tickmark everything** | Support every number |
| **Cross-reference** | Link to source workpapers |
| **Sign and date** | Preparer and reviewer sign-off |
| **Conclude** | State your conclusion |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
