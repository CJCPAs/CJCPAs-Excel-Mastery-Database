# Operating Expenses Audit VBA

> **OpEx Testing** - Complete VBA for auditing operating expenses per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 6000-6999, 7000-7999 (typically) |
| **Assertions** | Occurrence, Completeness, Accuracy, Classification |
| **Risk Level** | Low-Moderate (routine transactions) |
| **Categories** | SG&A, R&D, Depreciation, Other |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for operating expense accounts

### Input Sheet 2: `OpEx_Budget`
Budget vs actual by account (optional)

| Column | Header | Example |
|--------|--------|---------|
| A | `Account` | 6100 |
| B | `Description` | Salaries & Wages |
| C | `PY_Actual` | 1200000 |
| D | `CY_Budget` | 1300000 |
| E | `CY_Actual` | 1350000 |

---

## Audit Procedures

```vba
Sub AuditOperatingExpenses()
    '================================================
    ' OPERATING EXPENSES - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with expense transactions
    '   - Sheet "OpEx_Budget" with budget data (optional)
    '
    ' OUTPUTS:
    '   - Creates "OpEx_Audit" worksheet
    '   - Analyzes expenses by account
    '   - Performs YoY and budget variance analysis
    '   - Identifies unusual items
    '
    ' ASSERTIONS TESTED:
    '   - Occurrence (expenses are real)
    '   - Accuracy (amounts correct)
    '   - Classification (proper account)
    '================================================

    Dim wsGL As Worksheet
    Dim wsBudget As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const VARIANCE_THRESHOLD As Double = 0.15  ' 15% triggers review
    Const LARGE_ITEM_THRESHOLD As Double = 50000

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsBudget = ThisWorkbook.Sheets("OpEx_Budget")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("OpEx_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "OpEx_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "OPERATING EXPENSES - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: EXPENSE SUMMARY BY ACCOUNT
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: EXPENSE SUMMARY BY ACCOUNT"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        ' Aggregate by account
        Dim expDict As Object
        Set expDict = CreateObject("Scripting.Dictionary")
        Dim totalOpEx As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            Dim acctName As String
            Dim acctAmt As Double

            acctNum = CStr(wsGL.Cells(i, 3).Value)

            ' Operating expense accounts (6xxx, 7xxx)
            If Left(acctNum, 1) = "6" Or Left(acctNum, 1) = "7" Then
                acctName = wsGL.Cells(i, 4).Value
                acctAmt = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value

                If expDict.Exists(acctNum) Then
                    expDict(acctNum) = Array(acctName, expDict(acctNum)(1) + acctAmt)
                Else
                    expDict.Add acctNum, Array(acctName, acctAmt)
                End If

                totalOpEx = totalOpEx + acctAmt
            End If
        Next i

        .Cells(auditRow, 1).Value = "Account"
        .Cells(auditRow, 2).Value = "Description"
        .Cells(auditRow, 3).Value = "Balance"
        .Cells(auditRow, 4).Value = "% of Total"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        Dim acctStart As Long
        acctStart = auditRow

        Dim key As Variant
        For Each key In expDict.Keys
            Dim data As Variant
            data = expDict(key)
            .Cells(auditRow, 1).Value = key
            .Cells(auditRow, 2).Value = data(0)
            .Cells(auditRow, 3).Value = data(1)
            If totalOpEx <> 0 Then
                .Cells(auditRow, 4).Value = data(1) / totalOpEx
            End If

            ' Highlight large accounts
            If Abs(data(1)) > LARGE_ITEM_THRESHOLD Then
                .Cells(auditRow, 3).Font.Bold = True
            End If

            auditRow = auditRow + 1
        Next key

        .Cells(auditRow, 1).Value = "TOTAL OPERATING EXPENSES"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 3).Value = totalOpEx
        .Cells(auditRow, 3).Font.Bold = True
        .Cells(auditRow, 4).Value = 1
        auditRow = auditRow + 1

        .Range(.Cells(acctStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(acctStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "0.0%"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 2: BUDGET VS ACTUAL ANALYSIS
        ' ========================================
        If Not wsBudget Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 2: BUDGET VS ACTUAL ANALYSIS"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Account"
            .Cells(auditRow, 2).Value = "Budget"
            .Cells(auditRow, 3).Value = "Actual"
            .Cells(auditRow, 4).Value = "Variance"
            .Cells(auditRow, 5).Value = "Var %"
            .Cells(auditRow, 6).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
            auditRow = auditRow + 1

            Dim budStart As Long
            budStart = auditRow

            lastRow = wsBudget.Cells(wsBudget.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim budget As Double
                Dim actual As Double
                Dim variance As Double
                Dim varPct As Double

                budget = wsBudget.Cells(i, 4).Value
                actual = wsBudget.Cells(i, 5).Value
                variance = actual - budget

                If budget <> 0 Then
                    varPct = variance / Abs(budget)
                Else
                    varPct = 0
                End If

                .Cells(auditRow, 1).Value = wsBudget.Cells(i, 2).Value
                .Cells(auditRow, 2).Value = budget
                .Cells(auditRow, 3).Value = actual
                .Cells(auditRow, 4).Value = variance
                .Cells(auditRow, 5).Value = varPct

                If Abs(varPct) > VARIANCE_THRESHOLD Then
                    .Cells(auditRow, 6).Value = "INVESTIGATE"
                    .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                Else
                    .Cells(auditRow, 6).Value = "Reasonable"
                    .Cells(auditRow, 6).Interior.Color = RGB(198, 239, 206)
                End If

                auditRow = auditRow + 1
            Next i

            .Range(.Cells(budStart, 2), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Range(.Cells(budStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "0.0%"

            auditRow = auditRow + 2
        End If

        ' ========================================
        ' TEST 3: YEAR-OVER-YEAR ANALYSIS
        ' ========================================
        If Not wsBudget Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 3: YEAR-OVER-YEAR ANALYSIS"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Account"
            .Cells(auditRow, 2).Value = "PY Actual"
            .Cells(auditRow, 3).Value = "CY Actual"
            .Cells(auditRow, 4).Value = "$ Change"
            .Cells(auditRow, 5).Value = "% Change"
            .Cells(auditRow, 6).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
            auditRow = auditRow + 1

            Dim yoyStart As Long
            yoyStart = auditRow

            lastRow = wsBudget.Cells(wsBudget.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim pyActual As Double
                Dim cyActual As Double
                Dim dollarChg As Double
                Dim pctChg As Double

                pyActual = wsBudget.Cells(i, 3).Value
                cyActual = wsBudget.Cells(i, 5).Value
                dollarChg = cyActual - pyActual

                If pyActual <> 0 Then
                    pctChg = dollarChg / Abs(pyActual)
                Else
                    pctChg = 0
                End If

                .Cells(auditRow, 1).Value = wsBudget.Cells(i, 2).Value
                .Cells(auditRow, 2).Value = pyActual
                .Cells(auditRow, 3).Value = cyActual
                .Cells(auditRow, 4).Value = dollarChg
                .Cells(auditRow, 5).Value = pctChg

                If Abs(pctChg) > VARIANCE_THRESHOLD Then
                    .Cells(auditRow, 6).Value = "INVESTIGATE"
                    .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                Else
                    .Cells(auditRow, 6).Value = "Reasonable"
                    .Cells(auditRow, 6).Interior.Color = RGB(198, 239, 206)
                End If

                auditRow = auditRow + 1
            Next i

            .Range(.Cells(yoyStart, 2), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Range(.Cells(yoyStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "0.0%"

            auditRow = auditRow + 2
        End If

        ' ========================================
        ' TEST 4: MONTHLY TREND ANALYSIS
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: MONTHLY EXPENSE TREND"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Month"
        .Cells(auditRow, 2).Value = "Total OpEx"
        .Cells(auditRow, 3).Value = "Avg"
        .Cells(auditRow, 4).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        ' Aggregate by month
        Dim monthExp(1 To 12) As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            acctNum = CStr(wsGL.Cells(i, 3).Value)
            If Left(acctNum, 1) = "6" Or Left(acctNum, 1) = "7" Then
                Dim trxDate As Date
                On Error Resume Next
                trxDate = wsGL.Cells(i, 1).Value
                On Error GoTo 0

                If IsDate(trxDate) Then
                    Dim mo As Integer
                    mo = Month(trxDate)
                    monthExp(mo) = monthExp(mo) + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
                End If
            End If
        Next i

        Dim trendStart As Long
        trendStart = auditRow

        Dim avgMonthly As Double
        avgMonthly = totalOpEx / 12

        Dim monthNames As Variant
        monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

        For mo = 1 To 12
            .Cells(auditRow, 1).Value = monthNames(mo - 1)
            .Cells(auditRow, 2).Value = monthExp(mo)
            .Cells(auditRow, 3).Value = avgMonthly

            ' Flag months significantly above average
            If monthExp(mo) > avgMonthly * 1.25 Then
                .Cells(auditRow, 4).Value = "HIGH"
                .Cells(auditRow, 4).Interior.Color = RGB(255, 235, 156)
            ElseIf monthExp(mo) < avgMonthly * 0.75 Then
                .Cells(auditRow, 4).Value = "LOW"
                .Cells(auditRow, 4).Interior.Color = RGB(255, 235, 156)
            Else
                .Cells(auditRow, 4).Value = "Normal"
            End If

            auditRow = auditRow + 1
        Next mo

        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = totalOpEx
        .Cells(auditRow, 2).Font.Bold = True
        auditRow = auditRow + 1

        .Range(.Cells(trendStart, 2), .Cells(auditRow - 1, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 5: UNUSUAL ITEM IDENTIFICATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 5: UNUSUAL ITEMS IDENTIFICATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Date"
        .Cells(auditRow, 2).Value = "Account"
        .Cells(auditRow, 3).Value = "Description"
        .Cells(auditRow, 4).Value = "Amount"
        .Cells(auditRow, 5).Value = "Flag"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim unusualStart As Long
        unusualStart = auditRow

        ' Scan for unusual items
        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            acctNum = CStr(wsGL.Cells(i, 3).Value)
            If Left(acctNum, 1) = "6" Or Left(acctNum, 1) = "7" Then
                Dim itemAmt As Double
                Dim itemDesc As String

                itemAmt = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
                itemDesc = LCase(wsGL.Cells(i, 5).Value)

                ' Flag large individual items
                If Abs(itemAmt) > LARGE_ITEM_THRESHOLD Then
                    .Cells(auditRow, 1).Value = wsGL.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsGL.Cells(i, 4).Value
                    .Cells(auditRow, 3).Value = wsGL.Cells(i, 5).Value
                    .Cells(auditRow, 4).Value = itemAmt
                    .Cells(auditRow, 5).Value = "LARGE ITEM"
                    .Cells(auditRow, 5).Interior.Color = RGB(255, 235, 156)
                    auditRow = auditRow + 1
                ' Flag suspicious keywords
                ElseIf InStr(itemDesc, "adjust") > 0 Or InStr(itemDesc, "correct") > 0 Or _
                       InStr(itemDesc, "reclass") > 0 Or InStr(itemDesc, "related party") > 0 Then
                    .Cells(auditRow, 1).Value = wsGL.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsGL.Cells(i, 4).Value
                    .Cells(auditRow, 3).Value = wsGL.Cells(i, 5).Value
                    .Cells(auditRow, 4).Value = itemAmt
                    .Cells(auditRow, 5).Value = "KEYWORD"
                    .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
                    auditRow = auditRow + 1
                End If
            End If
        Next i

        If auditRow = unusualStart Then
            .Cells(auditRow, 1).Value = "No unusual items identified above threshold"
            auditRow = auditRow + 1
        Else
            .Range(.Cells(unusualStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        End If

        auditRow = auditRow + 2

        ' ========================================
        ' AUDIT SUMMARY
        ' ========================================
        .Cells(auditRow, 1).Value = "AUDIT SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Total Operating Expenses:"
        .Cells(auditRow, 2).Value = totalOpEx
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Account balance analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Monthly trend analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Unusual item identification"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Budget variance analysis (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Vouching of selected items (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 25
        .Columns("B:F").ColumnWidth = 14

    End With

    Application.ScreenUpdating = True

    MsgBox "Operating Expenses Audit Complete!" & vbCrLf & _
           "Total OpEx: " & Format(totalOpEx, "$#,##0"), vbInformation

End Sub
```

---

## Output Examples

### OpEx_Audit Worksheet

The `AuditOperatingExpenses` procedure generates a comprehensive worksheet:

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ OPERATING EXPENSES - AUDIT WORKPAPER                                                │
│ Period: December 31, 2024                                                           │
│ Prepared: AUDITOR on 12/15/2024 4:30:00 PM                                         │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

**TEST 1: EXPENSE SUMMARY BY ACCOUNT**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 1: EXPENSE SUMMARY BY ACCOUNT                                                  │
├─────────────┬────────────────────────────────┬────────────────┬─────────────────────┤
│ Account     │ Description                    │ Balance        │ % of Total          │
├─────────────┼────────────────────────────────┼────────────────┼─────────────────────┤
│ 6100        │ Salaries & Wages               │ $1,450,000     │ 48.3%               │
│ 6200        │ Employee Benefits              │ $290,000       │ 9.7%                │
│ 6300        │ Rent Expense                   │ $240,000       │ 8.0%                │
│ 6400        │ Utilities                      │ $72,000        │ 2.4%                │
│ 6500        │ Professional Fees              │ $185,000       │ 6.2%                │
│ 6600        │ Travel & Entertainment         │ $95,000        │ 3.2%                │
│ 6700        │ Depreciation                   │ $350,000       │ 11.7%               │
│ 6800        │ Insurance                      │ $125,000       │ 4.2%                │
│ 6900        │ Repairs & Maintenance          │ $85,000        │ 2.8%                │
│ 7100        │ Marketing & Advertising        │ $108,000       │ 3.6%                │
├─────────────┴────────────────────────────────┼────────────────┼─────────────────────┤
│ TOTAL OPERATING EXPENSES                     │ $3,000,000     │ 100.0%              │
└──────────────────────────────────────────────┴────────────────┴─────────────────────┘
```

**TEST 2: BUDGET VS ACTUAL ANALYSIS**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 2: BUDGET VS ACTUAL ANALYSIS                                                               │
├────────────────────────────────┬────────────────┬────────────────┬────────────┬────────┬────────┤
│ Account                        │ Budget         │ Actual         │ Variance   │ Var %  │ Status │
├────────────────────────────────┼────────────────┼────────────────┼────────────┼────────┼────────┤
│ Salaries & Wages               │ $1,400,000     │ $1,450,000     │ $50,000    │ 3.6%   │ ✓      │
│ Employee Benefits              │ $280,000       │ $290,000       │ $10,000    │ 3.6%   │ ✓      │
│ Rent Expense                   │ $240,000       │ $240,000       │ $0         │ 0.0%   │ ✓      │
│ Utilities                      │ $65,000        │ $72,000        │ $7,000     │ 10.8%  │ ✓      │
│ Professional Fees              │ $150,000       │ $185,000       │ $35,000    │ 23.3%  │ ⚠ INV  │
│ Travel & Entertainment         │ $80,000        │ $95,000        │ $15,000    │ 18.8%  │ ⚠ INV  │
│ Depreciation                   │ $340,000       │ $350,000       │ $10,000    │ 2.9%   │ ✓      │
│ Insurance                      │ $120,000       │ $125,000       │ $5,000     │ 4.2%   │ ✓      │
│ Repairs & Maintenance          │ $75,000        │ $85,000        │ $10,000    │ 13.3%  │ ✓      │
│ Marketing & Advertising        │ $100,000       │ $108,000       │ $8,000     │ 8.0%   │ ✓      │
└────────────────────────────────┴────────────────┴────────────────┴────────────┴────────┴────────┘
  Status: ⚠ INV = Investigate variance > 15%
```

**TEST 3: YEAR-OVER-YEAR ANALYSIS**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 3: YEAR-OVER-YEAR ANALYSIS                                                                 │
├────────────────────────────────┬────────────────┬────────────────┬────────────┬────────┬────────┤
│ Account                        │ PY Actual      │ CY Actual      │ $ Change   │ % Change│ Status │
├────────────────────────────────┼────────────────┼────────────────┼────────────┼────────┼────────┤
│ Salaries & Wages               │ $1,350,000     │ $1,450,000     │ $100,000   │ 7.4%   │ ✓      │
│ Employee Benefits              │ $265,000       │ $290,000       │ $25,000    │ 9.4%   │ ✓      │
│ Rent Expense                   │ $230,000       │ $240,000       │ $10,000    │ 4.3%   │ ✓      │
│ Utilities                      │ $68,000        │ $72,000        │ $4,000     │ 5.9%   │ ✓      │
│ Professional Fees              │ $140,000       │ $185,000       │ $45,000    │ 32.1%  │ ⚠ INV  │
│ Travel & Entertainment         │ $75,000        │ $95,000        │ $20,000    │ 26.7%  │ ⚠ INV  │
│ Depreciation                   │ $320,000       │ $350,000       │ $30,000    │ 9.4%   │ ✓      │
│ Insurance                      │ $115,000       │ $125,000       │ $10,000    │ 8.7%   │ ✓      │
│ Repairs & Maintenance          │ $82,000        │ $85,000        │ $3,000     │ 3.7%   │ ✓      │
│ Marketing & Advertising        │ $95,000        │ $108,000       │ $13,000    │ 13.7%  │ ✓      │
└────────────────────────────────┴────────────────┴────────────────┴────────────┴────────┴────────┘
```

**TEST 4: MONTHLY EXPENSE TREND**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 4: MONTHLY EXPENSE TREND                                                       │
├──────────────┬────────────────┬────────────────┬────────────────────────────────────┤
│ Month        │ Total OpEx     │ Avg            │ Status                             │
├──────────────┼────────────────┼────────────────┼────────────────────────────────────┤
│ Jan          │ $242,000       │ $250,000       │ Normal                             │
│ Feb          │ $238,000       │ $250,000       │ Normal                             │
│ Mar          │ $255,000       │ $250,000       │ Normal                             │
│ Apr          │ $248,000       │ $250,000       │ Normal                             │
│ May          │ $252,000       │ $250,000       │ Normal                             │
│ Jun          │ $265,000       │ $250,000       │ Normal                             │
│ Jul          │ $235,000       │ $250,000       │ Normal                             │
│ Aug          │ $240,000       │ $250,000       │ Normal                             │
│ Sep          │ $258,000       │ $250,000       │ Normal                             │
│ Oct          │ $262,000       │ $250,000       │ Normal                             │
│ Nov          │ $275,000       │ $250,000       │ Normal                             │
│ Dec          │ $330,000       │ $250,000       │ ⚠ HIGH                             │
├──────────────┼────────────────┼────────────────┼────────────────────────────────────┤
│ TOTAL        │ $3,000,000     │                │                                    │
└──────────────┴────────────────┴────────────────┴────────────────────────────────────┘
  Status: ⚠ HIGH = Significantly above average (>125%)
```

**TEST 5: UNUSUAL ITEMS IDENTIFICATION**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 5: UNUSUAL ITEMS IDENTIFICATION                                                                        │
├────────────┬────────────────────────────┬────────────────────────────────────────┬────────────┬─────────────┤
│ Date       │ Account                    │ Description                            │ Amount     │ Flag        │
├────────────┼────────────────────────────┼────────────────────────────────────────┼────────────┼─────────────┤
│ 03/15/2024 │ Professional Fees          │ Litigation settlement consulting       │ $65,000    │ ⚠ LARGE     │
│ 06/30/2024 │ Repairs & Maintenance      │ HVAC replacement - correction          │ $52,000    │ ✗ KEYWORD   │
│ 09/22/2024 │ Travel & Entertainment     │ Executive retreat expenses             │ $48,000    │ ⚠ LARGE     │
│ 11/15/2024 │ Professional Fees          │ Year-end audit fee adjustment          │ $35,000    │ ✗ KEYWORD   │
│ 12/20/2024 │ Marketing & Advertising    │ Reclass from prepaid                   │ $28,000    │ ✗ KEYWORD   │
└────────────┴────────────────────────────┴────────────────────────────────────────┴────────────┴─────────────┘
  ⚠ LARGE = Over $50,000 threshold | ✗ KEYWORD = Contains adjust/correct/reclass
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                               │
├─────────────────────────────────────────────────────────────────────────────┤
│ Total Operating Expenses: $3,000,000                                        │
├─────────────────────────────────────────────────────────────────────────────┤
│ Procedures Performed:                                                       │
│   ✓ Account balance analysis                                                │
│   ✓ Monthly trend analysis                                                  │
│   ✓ Unusual item identification                                             │
│   ☐ Budget variance analysis (manual)                                       │
│   ☐ Vouching of selected items (manual)                                     │
├─────────────────────────────────────────────────────────────────────────────┤
│ CONCLUSION: [Document conclusion]                                           │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Expense Vouching Template

```vba
Sub GenerateVouchingSample()
    '================================================
    ' GENERATE EXPENSE VOUCHING SAMPLE
    '
    ' Creates sample selection for vouching
    '================================================

    Dim wsGL As Worksheet
    Dim wsSample As Worksheet
    Dim lastRow As Long, i As Long
    Dim sampleRow As Long

    Const SAMPLE_SIZE As Long = 25
    Const MIN_AMOUNT As Double = 5000

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Sheets("Expense_Vouching").Delete
    On Error GoTo 0

    Set wsSample = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsSample.Name = "Expense_Vouching"

    With wsSample
        .Range("A1").Value = "EXPENSE VOUCHING SAMPLE"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Sample Size: " & SAMPLE_SIZE & " items"
        .Range("A3").Value = "Minimum Amount: " & Format(MIN_AMOUNT, "$#,##0")

        sampleRow = 5

        .Cells(sampleRow, 1).Value = "#"
        .Cells(sampleRow, 2).Value = "Date"
        .Cells(sampleRow, 3).Value = "Account"
        .Cells(sampleRow, 4).Value = "Description"
        .Cells(sampleRow, 5).Value = "Amount"
        .Cells(sampleRow, 6).Value = "Invoice"
        .Cells(sampleRow, 7).Value = "Approval"
        .Cells(sampleRow, 8).Value = "GL Agree"
        .Cells(sampleRow, 9).Value = "Status"
        .Range(.Cells(sampleRow, 1), .Cells(sampleRow, 9)).Font.Bold = True
        .Range(.Cells(sampleRow, 1), .Cells(sampleRow, 9)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(sampleRow, 1), .Cells(sampleRow, 9)).Font.Color = RGB(255, 255, 255)
        sampleRow = sampleRow + 1

        ' Build list of eligible items
        Dim eligible As Collection
        Set eligible = New Collection

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            acctNum = CStr(wsGL.Cells(i, 3).Value)

            If Left(acctNum, 1) = "6" Or Left(acctNum, 1) = "7" Then
                Dim amt As Double
                amt = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value

                If Abs(amt) >= MIN_AMOUNT Then
                    eligible.Add i
                End If
            End If
        Next i

        ' Random sample selection
        Dim selected As Object
        Set selected = CreateObject("Scripting.Dictionary")

        Randomize

        Dim sampleNum As Long
        sampleNum = 1

        Dim attempts As Long
        Do While selected.Count < Application.WorksheetFunction.Min(SAMPLE_SIZE, eligible.Count) And attempts < 1000
            Dim randIdx As Long
            randIdx = Int((eligible.Count) * Rnd + 1)
            Dim rowNum As Long
            rowNum = eligible(randIdx)

            If Not selected.Exists(rowNum) Then
                selected.Add rowNum, True

                .Cells(sampleRow, 1).Value = sampleNum
                .Cells(sampleRow, 2).Value = wsGL.Cells(rowNum, 1).Value
                .Cells(sampleRow, 3).Value = wsGL.Cells(rowNum, 4).Value
                .Cells(sampleRow, 4).Value = wsGL.Cells(rowNum, 5).Value
                .Cells(sampleRow, 5).Value = wsGL.Cells(rowNum, 6).Value - wsGL.Cells(rowNum, 7).Value
                .Cells(sampleRow, 5).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

                ' Leave documentation columns blank for completion
                .Cells(sampleRow, 6).Interior.Color = RGB(255, 255, 204)
                .Cells(sampleRow, 7).Interior.Color = RGB(255, 255, 204)
                .Cells(sampleRow, 8).Interior.Color = RGB(255, 255, 204)

                sampleRow = sampleRow + 1
                sampleNum = sampleNum + 1
            End If
            attempts = attempts + 1
        Loop

        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 35
        .Columns("E:I").ColumnWidth = 12

    End With

    MsgBox "Vouching Sample Generated!" & vbCrLf & _
           (sampleRow - 6) & " items selected for testing", vbInformation

End Sub
```

---

## Common Expense Categories

| Category | Accounts | Key Tests |
|----------|----------|-----------|
| **Salaries & Wages** | 6100 | Payroll testing |
| **Benefits** | 6200 | Plan documents |
| **Rent** | 6300 | Lease agreements |
| **Utilities** | 6400 | Subsequent payments |
| **Professional Fees** | 6500 | Contracts, invoices |
| **Travel & Entertainment** | 6600 | Expense reports |
| **Depreciation** | 6700 | FA register |
| **Insurance** | 6800 | Policies |
| **Repairs & Maintenance** | 6900 | Cap vs. expense |
| **Other** | 7xxx | Varies |

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ COGS](./cogs.md) | [➡️ Payroll](./payroll.md)
