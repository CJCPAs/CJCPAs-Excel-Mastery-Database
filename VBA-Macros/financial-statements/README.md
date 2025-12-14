# Financial Statements VBA Macros

> **There's a VBA for That!** - Generate balance sheets, income statements, and financial analyses

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [CreateBalanceSheet](#create-balance-sheet) | Generate formatted balance sheet |
| [CreateIncomeStatement](#create-income-statement) | Generate P&L statement |
| [ComparativePeriods](#create-comparative-statements) | Side-by-side period comparison |
| [CalculateRatios](#calculate-financial-ratios) | Key financial ratios |
| [VarianceAnalysis](#variance-analysis) | Budget vs Actual analysis |
| [ConsolidateStatements](#consolidate-statements) | Combine multiple entities |
| [RollForward](#roll-forward-balances) | Roll balances to next period |
| [TrialBalanceToFS](#trial-balance-to-financials) | Convert TB to financial statements |
| [AddFootnotes](#add-footnotes) | Insert financial statement notes |
| [FormatAsFinancials](#format-as-financial-statement) | Apply professional formatting |

---

## Create Balance Sheet

Generate a formatted balance sheet from trial balance data.

```vba
Sub CreateBalanceSheet()
    '================================================
    ' Create Balance Sheet from Trial Balance
    ' Assumes TB has Account#, AccountName, Balance columns
    '================================================

    Dim wsBS As Worksheet
    Dim wsTB As Worksheet
    Dim lastRow As Long
    Dim bsRow As Long
    Dim i As Long
    Dim acctNum As String
    Dim acctName As String
    Dim balance As Double

    ' Get Trial Balance sheet
    On Error Resume Next
    Set wsTB = ThisWorkbook.Sheets("Trial_Balance")
    On Error GoTo 0

    If wsTB Is Nothing Then
        MsgBox "Trial_Balance sheet not found.", vbExclamation
        Exit Sub
    End If

    ' Create Balance Sheet
    Set wsBS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsBS.Name = "Balance_Sheet"

    Application.ScreenUpdating = False

    With wsBS
        ' Header
        .Range("A1").Value = "BALANCE SHEET"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "As of " & Format(Date, "mmmm d, yyyy")

        bsRow = 4

        ' ASSETS
        .Cells(bsRow, 1).Value = "ASSETS"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 1).Font.Size = 12
        bsRow = bsRow + 1

        ' Current Assets
        .Cells(bsRow, 1).Value = "Current Assets"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 1).Font.Italic = True
        bsRow = bsRow + 1

        lastRow = wsTB.Cells(wsTB.Rows.Count, "A").End(xlUp).Row

        ' Loop through TB for asset accounts (1000-1999)
        For i = 2 To lastRow
            acctNum = CStr(wsTB.Cells(i, 1).Value)
            acctName = wsTB.Cells(i, 2).Value
            balance = wsTB.Cells(i, 3).Value

            If Left(acctNum, 1) = "1" Then
                If Val(acctNum) < 1500 Then  ' Current Assets
                    .Cells(bsRow, 2).Value = acctName
                    .Cells(bsRow, 3).Value = balance
                    bsRow = bsRow + 1
                End If
            End If
        Next i

        ' Current Assets Total
        .Cells(bsRow, 2).Value = "Total Current Assets"
        .Cells(bsRow, 2).Font.Bold = True
        .Cells(bsRow, 3).Formula = "=SUM(C6:C" & (bsRow - 1) & ")"
        .Cells(bsRow, 3).Font.Bold = True
        Dim currentAssetsTotal As Long
        currentAssetsTotal = bsRow
        bsRow = bsRow + 2

        ' Fixed Assets
        .Cells(bsRow, 1).Value = "Fixed Assets"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 1).Font.Italic = True
        bsRow = bsRow + 1
        Dim fixedStart As Long
        fixedStart = bsRow

        For i = 2 To lastRow
            acctNum = CStr(wsTB.Cells(i, 1).Value)
            acctName = wsTB.Cells(i, 2).Value
            balance = wsTB.Cells(i, 3).Value

            If Left(acctNum, 1) = "1" And Val(acctNum) >= 1500 Then
                .Cells(bsRow, 2).Value = acctName
                .Cells(bsRow, 3).Value = balance
                bsRow = bsRow + 1
            End If
        Next i

        ' Fixed Assets Total
        .Cells(bsRow, 2).Value = "Total Fixed Assets"
        .Cells(bsRow, 2).Font.Bold = True
        .Cells(bsRow, 3).Formula = "=SUM(C" & fixedStart & ":C" & (bsRow - 1) & ")"
        .Cells(bsRow, 3).Font.Bold = True
        bsRow = bsRow + 1

        ' TOTAL ASSETS
        .Cells(bsRow, 1).Value = "TOTAL ASSETS"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 3).Formula = "=C" & currentAssetsTotal & "+C" & (bsRow - 1)
        .Cells(bsRow, 3).Font.Bold = True
        .Range("A" & bsRow & ":C" & bsRow).Interior.Color = RGB(221, 235, 247)
        Dim totalAssets As Long
        totalAssets = bsRow
        bsRow = bsRow + 2

        ' LIABILITIES
        .Cells(bsRow, 1).Value = "LIABILITIES"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 1).Font.Size = 12
        bsRow = bsRow + 1

        ' Current Liabilities
        .Cells(bsRow, 1).Value = "Current Liabilities"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 1).Font.Italic = True
        bsRow = bsRow + 1
        Dim liabStart As Long
        liabStart = bsRow

        For i = 2 To lastRow
            acctNum = CStr(wsTB.Cells(i, 1).Value)
            acctName = wsTB.Cells(i, 2).Value
            balance = wsTB.Cells(i, 3).Value

            If Left(acctNum, 1) = "2" Then
                .Cells(bsRow, 2).Value = acctName
                .Cells(bsRow, 3).Value = Abs(balance)  ' Credit balance shown positive
                bsRow = bsRow + 1
            End If
        Next i

        ' Total Liabilities
        .Cells(bsRow, 2).Value = "Total Liabilities"
        .Cells(bsRow, 2).Font.Bold = True
        .Cells(bsRow, 3).Formula = "=SUM(C" & liabStart & ":C" & (bsRow - 1) & ")"
        .Cells(bsRow, 3).Font.Bold = True
        Dim totalLiab As Long
        totalLiab = bsRow
        bsRow = bsRow + 2

        ' EQUITY
        .Cells(bsRow, 1).Value = "EQUITY"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 1).Font.Size = 12
        bsRow = bsRow + 1
        Dim equityStart As Long
        equityStart = bsRow

        For i = 2 To lastRow
            acctNum = CStr(wsTB.Cells(i, 1).Value)
            acctName = wsTB.Cells(i, 2).Value
            balance = wsTB.Cells(i, 3).Value

            If Left(acctNum, 1) = "3" Then
                .Cells(bsRow, 2).Value = acctName
                .Cells(bsRow, 3).Value = Abs(balance)
                bsRow = bsRow + 1
            End If
        Next i

        ' Total Equity
        .Cells(bsRow, 2).Value = "Total Equity"
        .Cells(bsRow, 2).Font.Bold = True
        .Cells(bsRow, 3).Formula = "=SUM(C" & equityStart & ":C" & (bsRow - 1) & ")"
        .Cells(bsRow, 3).Font.Bold = True
        Dim totalEquity As Long
        totalEquity = bsRow
        bsRow = bsRow + 1

        ' TOTAL L&E
        .Cells(bsRow, 1).Value = "TOTAL LIABILITIES & EQUITY"
        .Cells(bsRow, 1).Font.Bold = True
        .Cells(bsRow, 3).Formula = "=C" & totalLiab & "+C" & totalEquity
        .Cells(bsRow, 3).Font.Bold = True
        .Range("A" & bsRow & ":C" & bsRow).Interior.Color = RGB(221, 235, 247)

        ' Format
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 35
        .Columns("C").ColumnWidth = 18
        .Range("C:C").NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"

    End With

    Application.ScreenUpdating = True

    MsgBox "Balance Sheet created!", vbInformation

End Sub
```

---

## Create Income Statement

Generate profit and loss statement.

```vba
Sub CreateIncomeStatement()
    '================================================
    ' Create Income Statement from Trial Balance
    '================================================

    Dim wsIS As Worksheet
    Dim wsTB As Worksheet
    Dim lastRow As Long
    Dim isRow As Long
    Dim i As Long
    Dim acctNum As String
    Dim acctName As String
    Dim balance As Double

    On Error Resume Next
    Set wsTB = ThisWorkbook.Sheets("Trial_Balance")
    On Error GoTo 0

    If wsTB Is Nothing Then
        MsgBox "Trial_Balance sheet not found.", vbExclamation
        Exit Sub
    End If

    Set wsIS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsIS.Name = "Income_Statement"

    Application.ScreenUpdating = False

    With wsIS
        ' Header
        .Range("A1").Value = "INCOME STATEMENT"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "For the Period Ending " & Format(Date, "mmmm d, yyyy")

        isRow = 4
        lastRow = wsTB.Cells(wsTB.Rows.Count, "A").End(xlUp).Row

        ' REVENUE
        .Cells(isRow, 1).Value = "REVENUE"
        .Cells(isRow, 1).Font.Bold = True
        isRow = isRow + 1
        Dim revStart As Long
        revStart = isRow

        For i = 2 To lastRow
            acctNum = CStr(wsTB.Cells(i, 1).Value)
            acctName = wsTB.Cells(i, 2).Value
            balance = wsTB.Cells(i, 3).Value

            If Left(acctNum, 1) = "4" Then  ' Revenue accounts
                .Cells(isRow, 2).Value = acctName
                .Cells(isRow, 3).Value = Abs(balance)
                isRow = isRow + 1
            End If
        Next i

        ' Total Revenue
        .Cells(isRow, 2).Value = "Total Revenue"
        .Cells(isRow, 2).Font.Bold = True
        .Cells(isRow, 3).Formula = "=SUM(C" & revStart & ":C" & (isRow - 1) & ")"
        .Cells(isRow, 3).Font.Bold = True
        .Range("A" & isRow & ":C" & isRow).Interior.Color = RGB(198, 239, 206)
        Dim totalRev As Long
        totalRev = isRow
        isRow = isRow + 2

        ' COST OF GOODS SOLD
        .Cells(isRow, 1).Value = "COST OF GOODS SOLD"
        .Cells(isRow, 1).Font.Bold = True
        isRow = isRow + 1
        Dim cogsStart As Long
        cogsStart = isRow

        For i = 2 To lastRow
            acctNum = CStr(wsTB.Cells(i, 1).Value)
            acctName = wsTB.Cells(i, 2).Value
            balance = wsTB.Cells(i, 3).Value

            If Left(acctNum, 1) = "5" Then  ' COGS accounts
                .Cells(isRow, 2).Value = acctName
                .Cells(isRow, 3).Value = balance
                isRow = isRow + 1
            End If
        Next i

        ' Total COGS
        .Cells(isRow, 2).Value = "Total Cost of Goods Sold"
        .Cells(isRow, 2).Font.Bold = True
        .Cells(isRow, 3).Formula = "=SUM(C" & cogsStart & ":C" & (isRow - 1) & ")"
        .Cells(isRow, 3).Font.Bold = True
        Dim totalCOGS As Long
        totalCOGS = isRow
        isRow = isRow + 1

        ' GROSS PROFIT
        .Cells(isRow, 1).Value = "GROSS PROFIT"
        .Cells(isRow, 1).Font.Bold = True
        .Cells(isRow, 3).Formula = "=C" & totalRev & "-C" & totalCOGS
        .Cells(isRow, 3).Font.Bold = True
        .Range("A" & isRow & ":C" & isRow).Interior.Color = RGB(221, 235, 247)
        Dim grossProfit As Long
        grossProfit = isRow
        isRow = isRow + 2

        ' OPERATING EXPENSES
        .Cells(isRow, 1).Value = "OPERATING EXPENSES"
        .Cells(isRow, 1).Font.Bold = True
        isRow = isRow + 1
        Dim expStart As Long
        expStart = isRow

        For i = 2 To lastRow
            acctNum = CStr(wsTB.Cells(i, 1).Value)
            acctName = wsTB.Cells(i, 2).Value
            balance = wsTB.Cells(i, 3).Value

            If Left(acctNum, 1) = "6" Or Left(acctNum, 1) = "7" Then  ' Expense accounts
                .Cells(isRow, 2).Value = acctName
                .Cells(isRow, 3).Value = balance
                isRow = isRow + 1
            End If
        Next i

        ' Total Expenses
        .Cells(isRow, 2).Value = "Total Operating Expenses"
        .Cells(isRow, 2).Font.Bold = True
        .Cells(isRow, 3).Formula = "=SUM(C" & expStart & ":C" & (isRow - 1) & ")"
        .Cells(isRow, 3).Font.Bold = True
        .Range("A" & isRow & ":C" & isRow).Interior.Color = RGB(255, 199, 206)
        Dim totalExp As Long
        totalExp = isRow
        isRow = isRow + 2

        ' NET INCOME
        .Cells(isRow, 1).Value = "NET INCOME"
        .Cells(isRow, 1).Font.Bold = True
        .Cells(isRow, 1).Font.Size = 12
        .Cells(isRow, 3).Formula = "=C" & grossProfit & "-C" & totalExp
        .Cells(isRow, 3).Font.Bold = True
        .Range("A" & isRow & ":C" & isRow).Interior.Color = RGB(0, 51, 102)
        .Range("A" & isRow & ":C" & isRow).Font.Color = RGB(255, 255, 255)

        ' Format
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 35
        .Columns("C").ColumnWidth = 18
        .Range("C:C").NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"

    End With

    Application.ScreenUpdating = True

    MsgBox "Income Statement created!", vbInformation

End Sub
```

---

## Create Comparative Statements

Generate side-by-side period comparison.

```vba
Sub CreateComparativeStatements()
    '================================================
    ' Create Comparative Financial Statements
    ' Current Period vs Prior Period with Variance
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    ' Assumes structure:
    ' Column A: Account Name
    ' Column B: Current Period
    ' Column C: Prior Period (or Budget)

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Add headers for variance columns
    ws.Range("D1").Value = "$ Variance"
    ws.Range("E1").Value = "% Variance"
    ws.Range("D1:E1").Font.Bold = True

    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, "B").Value) And IsNumeric(ws.Cells(i, "C").Value) Then
            ' $ Variance
            ws.Cells(i, "D").Formula = "=B" & i & "-C" & i

            ' % Variance (avoid divide by zero)
            ws.Cells(i, "E").Formula = "=IF(C" & i & "=0,0,(B" & i & "-C" & i & ")/ABS(C" & i & "))"
        End If
    Next i

    ' Format
    ws.Range("D:D").NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
    ws.Range("E:E").NumberFormat = "0.0%"

    ' Conditional formatting for variances
    ws.Range("D2:D" & lastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    ws.Range("D2:D" & lastRow).FormatConditions(1).Font.Color = RGB(192, 0, 0)

    ws.Range("E2:E" & lastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    ws.Range("E2:E" & lastRow).FormatConditions(1).Font.Color = RGB(192, 0, 0)

    MsgBox "Comparative analysis added!", vbInformation

End Sub
```

---

## Calculate Financial Ratios

Calculate key financial ratios.

```vba
Sub CalculateFinancialRatios()
    '================================================
    ' Calculate Key Financial Ratios
    '================================================

    Dim wsRatios As Worksheet
    Dim currentAssets As Double
    Dim currentLiab As Double
    Dim inventory As Double
    Dim totalAssets As Double
    Dim totalLiab As Double
    Dim totalEquity As Double
    Dim revenue As Double
    Dim netIncome As Double
    Dim receivables As Double

    ' Get inputs
    currentAssets = Application.InputBox("Enter Current Assets:", "Ratios", Type:=1)
    currentLiab = Application.InputBox("Enter Current Liabilities:", "Ratios", Type:=1)
    inventory = Application.InputBox("Enter Inventory:", "Ratios", Type:=1)
    totalAssets = Application.InputBox("Enter Total Assets:", "Ratios", Type:=1)
    totalLiab = Application.InputBox("Enter Total Liabilities:", "Ratios", Type:=1)
    totalEquity = Application.InputBox("Enter Total Equity:", "Ratios", Type:=1)
    revenue = Application.InputBox("Enter Total Revenue:", "Ratios", Type:=1)
    netIncome = Application.InputBox("Enter Net Income:", "Ratios", Type:=1)
    receivables = Application.InputBox("Enter Accounts Receivable:", "Ratios", Type:=1)

    ' Create ratios sheet
    Set wsRatios = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsRatios.Name = "Financial_Ratios"

    With wsRatios
        .Range("A1").Value = "FINANCIAL RATIO ANALYSIS"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        ' Liquidity Ratios
        .Range("A3").Value = "LIQUIDITY RATIOS"
        .Range("A3").Font.Bold = True
        .Range("A3").Interior.Color = RGB(0, 51, 102)
        .Range("A3").Font.Color = RGB(255, 255, 255)

        .Range("A4").Value = "Current Ratio"
        .Range("B4").Value = currentAssets / currentLiab
        .Range("C4").Value = "Current Assets / Current Liabilities"

        .Range("A5").Value = "Quick Ratio"
        .Range("B5").Value = (currentAssets - inventory) / currentLiab
        .Range("C5").Value = "(Current Assets - Inventory) / Current Liabilities"

        ' Leverage Ratios
        .Range("A7").Value = "LEVERAGE RATIOS"
        .Range("A7").Font.Bold = True
        .Range("A7").Interior.Color = RGB(0, 51, 102)
        .Range("A7").Font.Color = RGB(255, 255, 255)

        .Range("A8").Value = "Debt-to-Equity"
        .Range("B8").Value = totalLiab / totalEquity
        .Range("C8").Value = "Total Liabilities / Total Equity"

        .Range("A9").Value = "Debt Ratio"
        .Range("B9").Value = totalLiab / totalAssets
        .Range("C9").Value = "Total Liabilities / Total Assets"

        ' Profitability Ratios
        .Range("A11").Value = "PROFITABILITY RATIOS"
        .Range("A11").Font.Bold = True
        .Range("A11").Interior.Color = RGB(0, 51, 102)
        .Range("A11").Font.Color = RGB(255, 255, 255)

        .Range("A12").Value = "Profit Margin"
        .Range("B12").Value = netIncome / revenue
        .Range("C12").Value = "Net Income / Revenue"

        .Range("A13").Value = "Return on Assets (ROA)"
        .Range("B13").Value = netIncome / totalAssets
        .Range("C13").Value = "Net Income / Total Assets"

        .Range("A14").Value = "Return on Equity (ROE)"
        .Range("B14").Value = netIncome / totalEquity
        .Range("C14").Value = "Net Income / Total Equity"

        ' Activity Ratios
        .Range("A16").Value = "ACTIVITY RATIOS"
        .Range("A16").Font.Bold = True
        .Range("A16").Interior.Color = RGB(0, 51, 102)
        .Range("A16").Font.Color = RGB(255, 255, 255)

        .Range("A17").Value = "Receivables Turnover"
        .Range("B17").Value = revenue / receivables
        .Range("C17").Value = "Revenue / Accounts Receivable"

        .Range("A18").Value = "Days Sales Outstanding"
        .Range("B18").Value = 365 / (revenue / receivables)
        .Range("C18").Value = "365 / Receivables Turnover"

        .Range("A19").Value = "Asset Turnover"
        .Range("B19").Value = revenue / totalAssets
        .Range("C19").Value = "Revenue / Total Assets"

        ' Format
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 45
        .Range("B:B").NumberFormat = "0.00"
        .Range("B12:B14").NumberFormat = "0.00%"

    End With

    MsgBox "Financial ratios calculated!", vbInformation

End Sub
```

---

## Variance Analysis

Create budget vs actual variance analysis.

```vba
Sub VarianceAnalysis()
    '================================================
    ' Budget vs Actual Variance Analysis
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim threshold As Double

    Set ws = ActiveSheet

    ' Get variance threshold
    threshold = Application.InputBox("Enter variance threshold % to highlight (e.g., 10 for 10%):", "Variance Analysis", 10, Type:=1)
    threshold = threshold / 100

    ' Assumes:
    ' Column A: Account
    ' Column B: Budget
    ' Column C: Actual

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Add variance columns
    ws.Range("D1").Value = "$ Variance"
    ws.Range("E1").Value = "% Variance"
    ws.Range("F1").Value = "Status"
    ws.Range("D1:F1").Font.Bold = True

    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, "B").Value) And IsNumeric(ws.Cells(i, "C").Value) Then

            ' $ Variance (Actual - Budget)
            ws.Cells(i, "D").Formula = "=C" & i & "-B" & i

            ' % Variance
            ws.Cells(i, "E").Formula = "=IF(B" & i & "=0,0,(C" & i & "-B" & i & ")/ABS(B" & i & "))"

            ' Status based on threshold
            ws.Cells(i, "F").Formula = "=IF(ABS(E" & i & ")>" & threshold & ",""INVESTIGATE"",""OK"")"

        End If
    Next i

    ' Format
    ws.Range("D:D").NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"
    ws.Range("E:E").NumberFormat = "0.0%"

    ' Conditional formatting
    ws.Range("F2:F" & lastRow).FormatConditions.Add Type:=xlTextString, String:="INVESTIGATE", TextOperator:=xlContains
    ws.Range("F2:F" & lastRow).FormatConditions(1).Interior.Color = RGB(255, 199, 206)

    MsgBox "Variance analysis complete!" & vbCrLf & "Items over " & Format(threshold, "0%") & " are flagged.", vbInformation

End Sub
```

---

## Consolidate Statements

Combine multiple entity financial statements.

```vba
Sub ConsolidateStatements()
    '================================================
    ' Consolidate Financial Statements from Multiple Sheets
    '================================================

    Dim wsConsol As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim consolRow As Long
    Dim i As Long
    Dim col As Long
    Dim entityCount As Long

    ' Create consolidation sheet
    Set wsConsol = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsConsol.Name = "Consolidated"

    Application.ScreenUpdating = False

    With wsConsol
        .Range("A1").Value = "CONSOLIDATED FINANCIAL STATEMENT"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A3").Value = "Account"
        .Range("A3").Font.Bold = True

        col = 2
        entityCount = 0

        ' Loop through sheets to build headers and structure
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> "Consolidated" And ws.Name <> "Chart_of_Accounts" Then
                ' Add entity column header
                .Cells(3, col).Value = ws.Name
                .Cells(3, col).Font.Bold = True

                entityCount = entityCount + 1

                ' If first entity, copy account structure
                If col = 2 Then
                    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                    ws.Range("A4:A" & lastRow).Copy Destination:=.Range("A4")
                End If

                ' Copy amounts
                lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                For i = 4 To lastRow
                    On Error Resume Next
                    .Cells(i, col).Value = ws.Cells(i, "C").Value  ' Assumes amounts in column C
                    On Error GoTo 0
                Next i

                col = col + 1
            End If
        Next ws

        ' Add consolidated column
        .Cells(3, col).Value = "CONSOLIDATED"
        .Cells(3, col).Font.Bold = True
        .Cells(3, col).Interior.Color = RGB(0, 51, 102)
        .Cells(3, col).Font.Color = RGB(255, 255, 255)

        ' Sum across entities
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        For i = 4 To lastRow
            If .Cells(i, "B").Value <> "" Then
                .Cells(i, col).Formula = "=SUM(B" & i & ":" & Chr(64 + col - 1) & i & ")"
            End If
        Next i

        ' Format
        .Columns("A").ColumnWidth = 35
        .Range(.Cells(1, 2), .Cells(lastRow, col)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"

    End With

    Application.ScreenUpdating = True

    MsgBox "Consolidated " & entityCount & " entities!", vbInformation

End Sub
```

---

## Roll Forward Balances

Roll balances to next period.

```vba
Sub RollForwardBalances()
    '================================================
    ' Roll Forward Balances to Next Period
    '================================================

    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim newPeriod As String
    Dim lastRow As Long

    Set wsSource = ActiveSheet

    newPeriod = InputBox("Enter new period name:", "Roll Forward", Format(DateAdd("m", 1, Date), "yyyy-mm"))
    If newPeriod = "" Then Exit Sub

    ' Copy sheet
    wsSource.Copy After:=wsSource
    Set wsNew = ActiveSheet
    wsNew.Name = newPeriod

    With wsNew
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        ' Move current to prior period column
        ' Assumes: Column B = Prior, Column C = Current

        Dim i As Long
        For i = 2 To lastRow
            ' Prior period = what was current
            .Cells(i, "B").Value = .Cells(i, "C").Value
            ' Clear current period
            .Cells(i, "C").ClearContents
        Next i

        ' Update period header
        .Range("C1").Value = newPeriod
        .Range("B1").Value = wsSource.Range("C1").Value

    End With

    MsgBox "Rolled forward to: " & newPeriod, vbInformation

End Sub
```

---

## Trial Balance to Financials

Convert trial balance to financial statement format.

```vba
Sub TrialBalanceToFinancials()
    '================================================
    ' Convert Trial Balance to Financial Statements
    '================================================

    ' Simply calls the Balance Sheet and Income Statement macros
    Call CreateBalanceSheet
    Call CreateIncomeStatement

    MsgBox "Financial statements generated from trial balance!", vbInformation

End Sub
```

---

## Add Footnotes

Insert financial statement footnotes.

```vba
Sub AddFootnotes()
    '================================================
    ' Add Footnote Reference and Note
    '================================================

    Dim noteNum As Integer
    Dim noteText As String
    Dim notesRow As Long
    Dim ws As Worksheet

    Set ws = ActiveSheet

    noteNum = Application.InputBox("Enter footnote number:", "Add Footnote", 1, Type:=1)
    noteText = InputBox("Enter footnote text:", "Add Footnote", "")
    If noteText = "" Then Exit Sub

    ' Add superscript to current cell
    With Selection
        .Value = .Value & " (" & noteNum & ")"
    End With

    ' Find or create footnotes section
    notesRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 3

    If ws.Cells(notesRow - 1, "A").Value <> "NOTES:" Then
        ws.Cells(notesRow - 1, "A").Value = "NOTES:"
        ws.Cells(notesRow - 1, "A").Font.Bold = True
    End If

    ' Add footnote
    ws.Cells(notesRow, "A").Value = "(" & noteNum & ") " & noteText
    ws.Cells(notesRow, "A").Font.Size = 9

    MsgBox "Footnote added!", vbInformation

End Sub
```

---

## Format as Financial Statement

Apply professional financial statement formatting.

```vba
Sub FormatAsFinancialStatement()
    '================================================
    ' Apply Professional Financial Statement Formatting
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    Set ws = ActiveSheet
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.UsedRange.Columns.Count

    Application.ScreenUpdating = False

    With ws
        ' Font
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11

        ' Number format for currency columns (B onwards)
        .Range(.Cells(1, 2), .Cells(lastRow, lastCol)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""_);_(@_)"

        ' Column widths
        .Columns("A").ColumnWidth = 40
        Dim col As Long
        For col = 2 To lastCol
            .Columns(col).ColumnWidth = 15
        Next col

        ' Gridlines off
        ActiveWindow.DisplayGridlines = False

        ' Page setup
        .PageSetup.CenterHorizontally = True
        .PageSetup.TopMargin = Application.InchesToPoints(1)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.75)
        .PageSetup.PrintTitleRows = "$1:$3"

    End With

    Application.ScreenUpdating = True

    MsgBox "Financial statement formatting applied!", vbInformation

End Sub
```

---

## Best Practices for Financial Statements

| Practice | Description |
|----------|-------------|
| **Consistent periods** | Always compare like periods |
| **Round appropriately** | Use consistent rounding |
| **Show all notes** | Reference footnotes for significant items |
| **Verify cross-foots** | Assets = Liabilities + Equity |
| **Document sources** | Track where numbers come from |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
