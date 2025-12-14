# Reporting VBA Macros

> **There's a VBA for That!** - Generate reports, export to PDF, automate distributions

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [ExportToPDF](#export-to-pdf) | Save worksheet or workbook as PDF |
| [ExportAllSheetsToPDF](#export-all-sheets-to-pdf) | Create PDF of entire workbook |
| [CreateReportPackage](#create-report-package) | Bundle multiple reports |
| [EmailReport](#email-report-via-outlook) | Send report via Outlook |
| [PrintAreaSetup](#setup-print-area) | Define and manage print areas |
| [AddPageBreaks](#add-page-breaks) | Insert strategic page breaks |
| [CreateCoverPage](#create-cover-page) | Generate report cover page |
| [AddTableOfContents](#add-table-of-contents) | Auto-generate TOC |
| [DateStampReport](#date-stamp-report) | Add date/time stamps |
| [CreateDashboard](#create-summary-dashboard) | Build KPI dashboard |

---

## Export to PDF

Save current worksheet as PDF file.

```vba
Sub ExportToPDF()
    '================================================
    ' Export Active Sheet to PDF
    '================================================

    Dim filePath As String
    Dim fileName As String
    Dim ws As Worksheet

    Set ws = ActiveSheet

    ' Create default filename
    fileName = ws.Name & "_" & Format(Date, "YYYYMMDD") & ".pdf"

    ' Get save location
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=fileName, _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Save as PDF")

    If filePath = "False" Then Exit Sub

    ' Export to PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Exported to: " & filePath, vbInformation

End Sub

Sub ExportToPDFWithOptions()
    '================================================
    ' Export to PDF with Custom Options
    '================================================

    Dim filePath As String
    Dim ws As Worksheet
    Dim openAfter As VbMsgBoxResult

    Set ws = ActiveSheet

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=ws.Name & "_" & Format(Date, "YYYYMMDD") & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")

    If filePath = "False" Then Exit Sub

    openAfter = MsgBox("Open PDF after export?", vbYesNo + vbQuestion)

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=(openAfter = vbYes)

    MsgBox "PDF created: " & filePath, vbInformation

End Sub
```

---

## Export All Sheets to PDF

Create a single PDF containing all worksheets.

```vba
Sub ExportAllSheetsToPDF()
    '================================================
    ' Export Entire Workbook to PDF
    '================================================

    Dim filePath As String
    Dim wb As Workbook

    Set wb = ThisWorkbook

    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=Replace(wb.Name, ".xlsx", "") & "_" & Format(Date, "YYYYMMDD") & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Export Workbook to PDF")

    If filePath = "False" Then Exit Sub

    ' Export entire workbook
    wb.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Workbook exported to PDF!", vbInformation

End Sub

Sub ExportSelectedSheetsToPDF()
    '================================================
    ' Export Selected Sheets to PDF
    '================================================

    Dim filePath As String
    Dim ws As Worksheet
    Dim sheetList As String
    Dim selectedSheets() As String
    Dim i As Long

    ' Build list of sheets
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        sheetList = sheetList & ws.Name & vbCrLf
    Next ws

    ' Get sheets to include
    Dim sheetsToExport As String
    sheetsToExport = InputBox("Enter sheet names to export (comma-separated):" & vbCrLf & vbCrLf & _
                              "Available sheets:" & vbCrLf & sheetList, _
                              "Select Sheets", "Sheet1,Sheet2")

    If sheetsToExport = "" Then Exit Sub

    selectedSheets = Split(sheetsToExport, ",")

    ' Select the sheets
    ThisWorkbook.Sheets(Trim(selectedSheets(0))).Select
    For i = 1 To UBound(selectedSheets)
        ThisWorkbook.Sheets(Trim(selectedSheets(i))).Select Replace:=False
    Next i

    ' Get save location
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="Selected_Sheets_" & Format(Date, "YYYYMMDD") & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf")

    If filePath = "False" Then Exit Sub

    ' Export selected sheets
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Selected sheets exported!", vbInformation

End Sub
```

---

## Create Report Package

Bundle multiple reports into organized package.

```vba
Sub CreateReportPackage()
    '================================================
    ' Create Report Package (Multiple PDFs)
    '================================================

    Dim folderPath As String
    Dim ws As Worksheet
    Dim reportCount As Long
    Dim includeSheet As VbMsgBoxResult

    ' Get folder for reports
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder for Report Package"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With

    Application.ScreenUpdating = False

    ' Create subfolder with date
    Dim packageFolder As String
    packageFolder = folderPath & "Reports_" & Format(Date, "YYYYMMDD") & "\"

    On Error Resume Next
    MkDir packageFolder
    On Error GoTo 0

    ' Export each sheet
    For Each ws In ThisWorkbook.Worksheets
        includeSheet = MsgBox("Include '" & ws.Name & "' in report package?", vbYesNo + vbQuestion)

        If includeSheet = vbYes Then
            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=packageFolder & ws.Name & ".pdf", _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False

            reportCount = reportCount + 1
        End If
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Created " & reportCount & " reports in:" & vbCrLf & packageFolder, vbInformation

    ' Open folder
    Shell "explorer.exe """ & packageFolder & """", vbNormalFocus

End Sub
```

---

## Email Report via Outlook

Send report as email attachment.

```vba
Sub EmailReportViaOutlook()
    '================================================
    ' Email Report via Outlook
    '================================================

    Dim OutApp As Object
    Dim OutMail As Object
    Dim filePath As String
    Dim ws As Worksheet
    Dim toAddress As String
    Dim subject As String
    Dim body As String

    Set ws = ActiveSheet

    ' Get email details
    toAddress = InputBox("Enter recipient email:", "Email Report", "")
    If toAddress = "" Then Exit Sub

    subject = InputBox("Enter subject:", "Email Report", ws.Name & " Report - " & Format(Date, "mm/dd/yyyy"))
    body = InputBox("Enter message:", "Email Report", "Please find the attached report.")

    ' Create temp PDF
    filePath = Environ("TEMP") & "\" & ws.Name & "_" & Format(Date, "YYYYMMDD") & ".pdf"

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard

    ' Create Outlook email
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then
        Set OutApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If OutApp Is Nothing Then
        MsgBox "Could not open Outlook.", vbExclamation
        Exit Sub
    End If

    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = toAddress
        .Subject = subject
        .body = body
        .Attachments.Add filePath
        .Display  ' Use .Send to send automatically
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

    MsgBox "Email created with report attached!", vbInformation

End Sub

Sub EmailWorkbookAsAttachment()
    '================================================
    ' Email Entire Workbook as Attachment
    '================================================

    Dim OutApp As Object
    Dim OutMail As Object
    Dim toAddress As String

    toAddress = InputBox("Enter recipient email:", "Email Workbook")
    If toAddress = "" Then Exit Sub

    ' Save workbook first
    ThisWorkbook.Save

    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If OutApp Is Nothing Then
        MsgBox "Could not start Outlook.", vbExclamation
        Exit Sub
    End If

    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = toAddress
        .Subject = ThisWorkbook.Name & " - " & Format(Date, "mm/dd/yyyy")
        .body = "Please find the attached workbook."
        .Attachments.Add ThisWorkbook.FullName
        .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
```

---

## Setup Print Area

Define and manage print areas.

```vba
Sub SetupPrintArea()
    '================================================
    ' Set Print Area for Current Sheet
    '================================================

    Dim rng As Range

    On Error Resume Next
    Set rng = Application.InputBox("Select print area:", "Set Print Area", Type:=8)
    On Error GoTo 0

    If rng Is Nothing Then Exit Sub

    ActiveSheet.PageSetup.PrintArea = rng.Address

    MsgBox "Print area set to: " & rng.Address, vbInformation

End Sub

Sub ClearPrintArea()
    '================================================
    ' Clear Print Area
    '================================================

    ActiveSheet.PageSetup.PrintArea = ""
    MsgBox "Print area cleared.", vbInformation

End Sub

Sub SetPrintAreaAllData()
    '================================================
    ' Set Print Area to All Data
    '================================================

    ActiveSheet.PageSetup.PrintArea = ActiveSheet.UsedRange.Address
    MsgBox "Print area set to all data.", vbInformation

End Sub

Sub SetupPrintTitles()
    '================================================
    ' Set Rows to Repeat at Top
    '================================================

    Dim rowsToRepeat As String

    rowsToRepeat = InputBox("Enter rows to repeat (e.g., $1:$3):", "Print Titles", "$1:$1")

    If rowsToRepeat = "" Then Exit Sub

    ActiveSheet.PageSetup.PrintTitleRows = rowsToRepeat

    MsgBox "Print titles set!", vbInformation

End Sub
```

---

## Add Page Breaks

Insert page breaks at strategic locations.

```vba
Sub AddPageBreaks()
    '================================================
    ' Add Page Breaks at Selection
    '================================================

    ' Horizontal page break (above selection)
    ActiveSheet.HPageBreaks.Add Before:=Selection

    MsgBox "Horizontal page break added.", vbInformation

End Sub

Sub AddVerticalPageBreak()
    '================================================
    ' Add Vertical Page Break
    '================================================

    ActiveSheet.VPageBreaks.Add Before:=Selection

    MsgBox "Vertical page break added.", vbInformation

End Sub

Sub RemoveAllPageBreaks()
    '================================================
    ' Remove All Manual Page Breaks
    '================================================

    ActiveSheet.ResetAllPageBreaks

    MsgBox "All page breaks removed.", vbInformation

End Sub

Sub InsertPageBreaksByGroup()
    '================================================
    ' Insert Page Breaks When Group Changes
    ' Useful for reports grouped by account/department
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim groupCol As Integer
    Dim breakCount As Long

    Set ws = ActiveSheet

    groupCol = Application.InputBox("Enter column number for grouping:", "Page Breaks", 1, Type:=1)
    If groupCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, groupCol).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 3 To lastRow
        If ws.Cells(i, groupCol).Value <> ws.Cells(i - 1, groupCol).Value Then
            ws.HPageBreaks.Add Before:=ws.Cells(i, 1)
            breakCount = breakCount + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Added " & breakCount & " page breaks.", vbInformation

End Sub
```

---

## Create Cover Page

Generate professional report cover page.

```vba
Sub CreateCoverPage()
    '================================================
    ' Create Report Cover Page
    '================================================

    Dim wsCover As Worksheet
    Dim reportTitle As String
    Dim clientName As String
    Dim periodEnd As String

    reportTitle = InputBox("Report Title:", "Cover Page", "Financial Statements")
    clientName = InputBox("Client Name:", "Cover Page", "Client Name")
    periodEnd = InputBox("Period Ending:", "Cover Page", Format(DateSerial(Year(Date), 12, 31), "December 31, yyyy"))

    ' Create cover sheet
    Set wsCover = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    wsCover.Name = "Cover"

    Application.ScreenUpdating = False

    With wsCover
        ' Client Name
        .Range("A10").Value = clientName
        .Range("A10").Font.Size = 24
        .Range("A10").Font.Bold = True
        .Range("A10").HorizontalAlignment = xlCenter
        .Range("A10:H10").Merge

        ' Report Title
        .Range("A15").Value = reportTitle
        .Range("A15").Font.Size = 28
        .Range("A15").Font.Bold = True
        .Range("A15").HorizontalAlignment = xlCenter
        .Range("A15:H15").Merge

        ' Period
        .Range("A20").Value = "For the Period Ending"
        .Range("A20").Font.Size = 14
        .Range("A20").HorizontalAlignment = xlCenter
        .Range("A20:H20").Merge

        .Range("A22").Value = periodEnd
        .Range("A22").Font.Size = 18
        .Range("A22").Font.Bold = True
        .Range("A22").HorizontalAlignment = xlCenter
        .Range("A22:H22").Merge

        ' Prepared By
        .Range("A35").Value = "Prepared by:"
        .Range("A35").Font.Size = 12
        .Range("A35").HorizontalAlignment = xlCenter
        .Range("A35:H35").Merge

        .Range("A37").Value = "[Your Firm Name]"
        .Range("A37").Font.Size = 14
        .Range("A37").HorizontalAlignment = xlCenter
        .Range("A37:H37").Merge

        ' Date prepared
        .Range("A40").Value = Format(Date, "mmmm d, yyyy")
        .Range("A40").Font.Size = 11
        .Range("A40").HorizontalAlignment = xlCenter
        .Range("A40:H40").Merge

        ' Remove gridlines
        ActiveWindow.DisplayGridlines = False

        ' Page setup
        .PageSetup.CenterHorizontally = True
        .PageSetup.CenterVertically = True

    End With

    Application.ScreenUpdating = True

    MsgBox "Cover page created!", vbInformation

End Sub
```

---

## Add Table of Contents

Auto-generate table of contents.

```vba
Sub AddTableOfContents()
    '================================================
    ' Create Table of Contents Sheet
    '================================================

    Dim wsTOC As Worksheet
    Dim ws As Worksheet
    Dim tocRow As Long

    ' Create TOC sheet
    On Error Resume Next
    ThisWorkbook.Sheets("Contents").Delete
    On Error GoTo 0

    Set wsTOC = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
    wsTOC.Name = "Contents"

    Application.ScreenUpdating = False

    With wsTOC
        .Range("A1").Value = "TABLE OF CONTENTS"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16

        .Range("A3").Value = "Section"
        .Range("B3").Value = "Description"
        .Range("A3:B3").Font.Bold = True

        tocRow = 4

        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> "Contents" And ws.Name <> "Cover" Then
                ' Add hyperlink to sheet
                .Hyperlinks.Add Anchor:=.Cells(tocRow, 1), _
                    Address:="", _
                    SubAddress:="'" & ws.Name & "'!A1", _
                    TextToDisplay:=ws.Name

                ' Try to get description from A2
                On Error Resume Next
                .Cells(tocRow, 2).Value = ws.Range("A2").Value
                On Error GoTo 0

                tocRow = tocRow + 1
            End If
        Next ws

        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 50

        ' Add borders
        .Range("A3:B" & tocRow - 1).Borders.LineStyle = xlContinuous

    End With

    Application.ScreenUpdating = True

    MsgBox "Table of Contents created!", vbInformation

End Sub
```

---

## Date Stamp Report

Add date and time stamps to report.

```vba
Sub DateStampReport()
    '================================================
    ' Add Date/Time Stamp to Report
    '================================================

    Dim stampLocation As String

    stampLocation = InputBox("Where to add stamp?" & vbCrLf & _
                            "1 = Header" & vbCrLf & _
                            "2 = Footer" & vbCrLf & _
                            "3 = Cell (select first)", _
                            "Date Stamp", "3")

    Select Case stampLocation
        Case "1"
            ActiveSheet.PageSetup.RightHeader = "Printed: &D &T"
            MsgBox "Date stamp added to header.", vbInformation

        Case "2"
            ActiveSheet.PageSetup.RightFooter = "Printed: &D &T"
            MsgBox "Date stamp added to footer.", vbInformation

        Case "3"
            Selection.Value = "Generated: " & Format(Now, "mm/dd/yyyy hh:mm AM/PM")
            Selection.Font.Italic = True
            Selection.Font.Size = 9
    End Select

End Sub

Sub AddPrintedDateToAllSheets()
    '================================================
    ' Add Printed Date to Footer of All Sheets
    '================================================

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ws.PageSetup.RightFooter = "Printed: &D &T"
    Next ws

    MsgBox "Date stamp added to all sheet footers.", vbInformation

End Sub
```

---

## Create Summary Dashboard

Build a KPI summary dashboard.

```vba
Sub CreateSummaryDashboard()
    '================================================
    ' Create KPI Summary Dashboard
    '================================================

    Dim wsDash As Worksheet

    ' Create dashboard sheet
    On Error Resume Next
    ThisWorkbook.Sheets("Dashboard").Delete
    On Error GoTo 0

    Set wsDash = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    wsDash.Name = "Dashboard"

    Application.ScreenUpdating = False

    With wsDash
        ' Title
        .Range("A1").Value = "EXECUTIVE SUMMARY DASHBOARD"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 18
        .Range("A1:H1").Merge

        .Range("A2").Value = "As of " & Format(Date, "mmmm d, yyyy")
        .Range("A2:H2").Merge

        ' KPI Section 1 - Financial
        .Range("A4").Value = "FINANCIAL HIGHLIGHTS"
        .Range("A4").Font.Bold = True
        .Range("A4").Interior.Color = RGB(0, 51, 102)
        .Range("A4").Font.Color = RGB(255, 255, 255)
        .Range("A4:D4").Merge

        .Range("A5").Value = "Total Revenue"
        .Range("A6").Value = "Total Expenses"
        .Range("A7").Value = "Net Income"
        .Range("A8").Value = "Profit Margin"

        .Range("B5:B8").NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
        .Range("B8").NumberFormat = "0.0%"

        ' KPI Section 2 - Balance Sheet
        .Range("A10").Value = "BALANCE SHEET SUMMARY"
        .Range("A10").Font.Bold = True
        .Range("A10").Interior.Color = RGB(0, 51, 102)
        .Range("A10").Font.Color = RGB(255, 255, 255)
        .Range("A10:D10").Merge

        .Range("A11").Value = "Total Assets"
        .Range("A12").Value = "Total Liabilities"
        .Range("A13").Value = "Total Equity"
        .Range("A14").Value = "Current Ratio"

        .Range("B11:B13").NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
        .Range("B14").NumberFormat = "0.00"

        ' KPI Section 3 - Metrics
        .Range("A16").Value = "KEY RATIOS"
        .Range("A16").Font.Bold = True
        .Range("A16").Interior.Color = RGB(0, 51, 102)
        .Range("A16").Font.Color = RGB(255, 255, 255)
        .Range("A16:D16").Merge

        .Range("A17").Value = "Quick Ratio"
        .Range("A18").Value = "Debt-to-Equity"
        .Range("A19").Value = "ROA"
        .Range("A20").Value = "ROE"

        .Range("B17:B18").NumberFormat = "0.00"
        .Range("B19:B20").NumberFormat = "0.0%"

        ' Format
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 15

        ' Remove gridlines
        ActiveWindow.DisplayGridlines = False

    End With

    Application.ScreenUpdating = True

    MsgBox "Dashboard template created!" & vbCrLf & "Link cells to your data sources.", vbInformation

End Sub
```

---

## Quick Reference: Page Setup

| Setting | Code |
|---------|------|
| Portrait | `PageSetup.Orientation = xlPortrait` |
| Landscape | `PageSetup.Orientation = xlLandscape` |
| Fit to 1 page wide | `PageSetup.FitToPagesWide = 1` |
| Fit to 1 page tall | `PageSetup.FitToPagesTall = 1` |
| Margins | `PageSetup.LeftMargin = Application.InchesToPoints(0.5)` |
| Center horizontally | `PageSetup.CenterHorizontally = True` |
| Print gridlines | `PageSetup.PrintGridlines = True` |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
