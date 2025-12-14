# Revenue Audit VBA

> **Revenue Recognition Testing** - Complete VBA for auditing revenue per ASC 606 and GAAS

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 4000-4999 (typically) |
| **Assertions** | Occurrence, Completeness, Accuracy, Cutoff, Classification |
| **Risk Level** | **CRITICAL** (fraud risk, significant estimate) |
| **Key Standards** | ASC 606 (Revenue from Contracts with Customers) |

**Key Audit Focus:** Revenue is a presumed fraud risk per GAAS. Focus on **occurrence** (did the sale happen?) and **cutoff** (correct period).

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for revenue accounts

| Column | Header | Example |
|--------|--------|---------|
| A | `Date` | 12/31/2024 |
| B | `JE_Number` | JE-2024-9999 |
| C | `Account` | 4100 |
| D | `Account_Name` | Product Revenue |
| E | `Description` | Invoice #5001 |
| F | `Debit` | 0 |
| G | `Credit` | 50000 |
| H | `Source` | AR |

### Input Sheet 2: `Sales_Detail`
Individual sales transactions

| Column | Header | Example |
|--------|--------|---------|
| A | `Invoice_Number` | INV-5001 |
| B | `Invoice_Date` | 12/28/2024 |
| C | `Customer` | Acme Corp |
| D | `Ship_Date` | 12/29/2024 |
| E | `Product` | Widget A |
| F | `Quantity` | 100 |
| G | `Unit_Price` | 500 |
| H | `Total` | 50000 |
| I | `Payment_Received` | 01/15/2025 |

### Input Sheet 3: `Prior_Year_Revenue`
Prior year revenue by month for comparison

| Column | Header | Example |
|--------|--------|---------|
| A | `Month` | January |
| B | `PY_Revenue` | 500000 |

---

## Audit Procedures

### 1. Complete Revenue Audit Module

```vba
Sub AuditRevenue()
    '================================================
    ' REVENUE - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with revenue transactions
    '   - Sheet "Sales_Detail" with invoice detail
    '   - Sheet "Prior_Year_Revenue" for analytics (optional)
    '
    ' OUTPUTS:
    '   - Creates "Revenue_Audit" worksheet
    '   - Performs analytical procedures
    '   - Tests cutoff and occurrence
    '   - Identifies unusual transactions
    '
    ' ASSERTIONS TESTED:
    '   - Occurrence (did the sale happen?)
    '   - Accuracy (amounts correct)
    '   - Cutoff (correct period)
    '   - Classification (proper account)
    '================================================

    Dim wsGL As Worksheet
    Dim wsSales As Worksheet
    Dim wsPY As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    ' Materiality thresholds
    Const MATERIALITY As Double = 50000
    Const TRIVIAL As Double = 2500
    Const VARIANCE_THRESHOLD As Double = 0.15  ' 15% variance triggers investigation

    ' Validate required sheets
    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsSales = ThisWorkbook.Sheets("Sales_Detail")
    Set wsPY = ThisWorkbook.Sheets("Prior_Year_Revenue")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Revenue_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Revenue_Audit"

    Application.ScreenUpdating = False

    ' ========================================
    ' HEADER
    ' ========================================
    With wsAudit
        .Range("A1").Value = "REVENUE - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: Year Ended " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now
        .Range("A4").Value = "Materiality: " & Format(MATERIALITY, "$#,##0")

        .Range("A5").Value = "NOTE: Revenue is a PRESUMED FRAUD RISK per AU-C 240"
        .Range("A5").Font.Bold = True
        .Range("A5").Interior.Color = RGB(255, 199, 206)

        auditRow = 7
    End With

    ' ========================================
    ' TEST 1: REVENUE BY MONTH - ANALYTICAL PROCEDURES
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 1: MONTHLY REVENUE ANALYSIS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Month"
        .Cells(auditRow, 2).Value = "CY Revenue"
        .Cells(auditRow, 3).Value = "PY Revenue"
        .Cells(auditRow, 4).Value = "$ Change"
        .Cells(auditRow, 5).Value = "% Change"
        .Cells(auditRow, 6).Value = "CY % of Total"
        .Cells(auditRow, 7).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
        auditRow = auditRow + 1

        Dim monthStart As Long
        monthStart = auditRow

        ' Calculate revenue by month from GL
        Dim monthlyRev(1 To 12) As Double
        Dim totalRevenue As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 1) = "4" Then  ' Revenue accounts
                Dim transDate As Date
                Dim transMonth As Integer
                Dim transAmt As Double

                If IsDate(wsGL.Cells(i, 1).Value) Then
                    transDate = wsGL.Cells(i, 1).Value
                    transMonth = Month(transDate)
                    transAmt = wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value  ' Credit - Debit

                    monthlyRev(transMonth) = monthlyRev(transMonth) + transAmt
                    totalRevenue = totalRevenue + transAmt
                End If
            End If
        Next i

        ' Output monthly analysis
        Dim monthNames As Variant
        monthNames = Array("", "January", "February", "March", "April", "May", "June", _
                          "July", "August", "September", "October", "November", "December")

        Dim pyRev As Double
        Dim cyRev As Double
        Dim pctChange As Double

        For i = 1 To 12
            .Cells(auditRow, 1).Value = monthNames(i)
            .Cells(auditRow, 2).Value = monthlyRev(i)

            ' Get PY revenue if available
            If Not wsPY Is Nothing Then
                On Error Resume Next
                pyRev = wsPY.Cells(i + 1, 2).Value
                On Error GoTo 0
                .Cells(auditRow, 3).Value = pyRev
                .Cells(auditRow, 4).Value = monthlyRev(i) - pyRev

                If pyRev <> 0 Then
                    pctChange = (monthlyRev(i) - pyRev) / Abs(pyRev)
                    .Cells(auditRow, 5).Value = pctChange

                    If Abs(pctChange) > VARIANCE_THRESHOLD Then
                        .Cells(auditRow, 7).Value = "INVESTIGATE"
                        .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                    Else
                        .Cells(auditRow, 7).Value = "OK"
                        .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
                    End If
                End If
            End If

            .Cells(auditRow, 6).Value = monthlyRev(i) / totalRevenue

            auditRow = auditRow + 1
        Next i

        ' Totals
        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Formula = "=SUM(B" & monthStart & ":B" & (auditRow - 1) & ")"
        .Cells(auditRow, 2).Font.Bold = True

        ' Format
        .Range(.Cells(monthStart, 2), .Cells(auditRow, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(monthStart, 5), .Cells(auditRow - 1, 6)).NumberFormat = "0.0%"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 2: REVENUE BY ACCOUNT
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 2: REVENUE BY ACCOUNT"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Account"
        .Cells(auditRow, 2).Value = "Account Name"
        .Cells(auditRow, 3).Value = "Balance"
        .Cells(auditRow, 4).Value = "% of Total"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        Dim acctStart As Long
        acctStart = auditRow

        ' Aggregate by account
        Dim acctDict As Object
        Set acctDict = CreateObject("Scripting.Dictionary")

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 1) = "4" Then
                Dim acctNum As String
                Dim acctName As String
                Dim acctAmt As Double

                acctNum = wsGL.Cells(i, 3).Value
                acctName = wsGL.Cells(i, 4).Value
                acctAmt = wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value

                If acctDict.Exists(acctNum) Then
                    acctDict(acctNum) = Array(acctName, acctDict(acctNum)(1) + acctAmt)
                Else
                    acctDict.Add acctNum, Array(acctName, acctAmt)
                End If
            End If
        Next i

        ' Output accounts
        Dim key As Variant
        Dim acctData As Variant

        For Each key In acctDict.Keys
            acctData = acctDict(key)
            .Cells(auditRow, 1).Value = key
            .Cells(auditRow, 2).Value = acctData(0)
            .Cells(auditRow, 3).Value = acctData(1)
            .Cells(auditRow, 4).Value = acctData(1) / totalRevenue
            auditRow = auditRow + 1
        Next key

        .Range(.Cells(acctStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(acctStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "0.0%"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 3: SALES CUTOFF TESTING
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 3: REVENUE CUTOFF TESTING"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Testing sales within 5 days of year-end for proper cutoff"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Invoice #"
        .Cells(auditRow, 2).Value = "Invoice Date"
        .Cells(auditRow, 3).Value = "Customer"
        .Cells(auditRow, 4).Value = "Ship Date"
        .Cells(auditRow, 5).Value = "Amount"
        .Cells(auditRow, 6).Value = "Days from Y/E"
        .Cells(auditRow, 7).Value = "Cutoff Issue?"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
        auditRow = auditRow + 1

        Dim cutoffStart As Long
        cutoffStart = auditRow

        If Not wsSales Is Nothing Then
            Dim yearEnd As Date
            yearEnd = DateSerial(Year(Date), 12, 31)

            lastRow = wsSales.Cells(wsSales.Rows.Count, "A").End(xlUp).Row

            For i = 2 To lastRow
                Dim invDate As Date
                Dim shipDate As Date
                Dim daysFromYE As Long

                If IsDate(wsSales.Cells(i, 2).Value) Then
                    invDate = wsSales.Cells(i, 2).Value
                    daysFromYE = invDate - yearEnd

                    If Abs(daysFromYE) <= 5 Then
                        .Cells(auditRow, 1).Value = wsSales.Cells(i, 1).Value
                        .Cells(auditRow, 2).Value = invDate
                        .Cells(auditRow, 3).Value = wsSales.Cells(i, 3).Value

                        If IsDate(wsSales.Cells(i, 4).Value) Then
                            shipDate = wsSales.Cells(i, 4).Value
                            .Cells(auditRow, 4).Value = shipDate
                        End If

                        .Cells(auditRow, 5).Value = wsSales.Cells(i, 8).Value
                        .Cells(auditRow, 6).Value = daysFromYE

                        ' Check for cutoff issues
                        If daysFromYE <= 0 Then  ' Recorded in CY
                            If IsDate(wsSales.Cells(i, 4).Value) Then
                                If shipDate > yearEnd Then
                                    .Cells(auditRow, 7).Value = "SHIP AFTER Y/E - CUTOFF ERROR"
                                    .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                                Else
                                    .Cells(auditRow, 7).Value = "OK"
                                    .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
                                End If
                            Else
                                .Cells(auditRow, 7).Value = "VERIFY SHIP DATE"
                                .Cells(auditRow, 7).Interior.Color = RGB(255, 235, 156)
                            End If
                        Else  ' Recorded after Y/E
                            If IsDate(wsSales.Cells(i, 4).Value) Then
                                If shipDate <= yearEnd Then
                                    .Cells(auditRow, 7).Value = "SHIP BEFORE Y/E - RECORD IN CY?"
                                    .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                                Else
                                    .Cells(auditRow, 7).Value = "OK - Next year"
                                    .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
                                End If
                            End If
                        End If

                        auditRow = auditRow + 1
                    End If
                End If
            Next i
        Else
            .Cells(auditRow, 1).Value = "Sales_Detail sheet not found - manual cutoff testing required"
            .Cells(auditRow, 1).Font.Italic = True
            auditRow = auditRow + 1
        End If

        .Range(.Cells(cutoffStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 4: LARGE/UNUSUAL TRANSACTIONS
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 4: LARGE & UNUSUAL REVENUE TRANSACTIONS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Date"
        .Cells(auditRow, 2).Value = "JE Number"
        .Cells(auditRow, 3).Value = "Description"
        .Cells(auditRow, 4).Value = "Amount"
        .Cells(auditRow, 5).Value = "% of Materiality"
        .Cells(auditRow, 6).Value = "Action Required"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        Dim largeStart As Long
        largeStart = auditRow

        Dim largeThreshold As Double
        largeThreshold = MATERIALITY * 0.25

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 1) = "4" Then
                Dim revAmt As Double
                revAmt = wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value

                If Abs(revAmt) >= largeThreshold Then
                    .Cells(auditRow, 1).Value = wsGL.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsGL.Cells(i, 2).Value
                    .Cells(auditRow, 3).Value = wsGL.Cells(i, 5).Value
                    .Cells(auditRow, 4).Value = revAmt
                    .Cells(auditRow, 5).Value = Abs(revAmt) / MATERIALITY

                    If Abs(revAmt) >= MATERIALITY Then
                        .Cells(auditRow, 6).Value = "MATERIAL - VOUCH TO SUPPORT"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                    Else
                        .Cells(auditRow, 6).Value = "Review support"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 235, 156)
                    End If

                    auditRow = auditRow + 1
                End If
            End If
        Next i

        .Range(.Cells(largeStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(largeStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "0%"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 5: CREDIT MEMOS / REVERSALS
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 5: CREDIT MEMOS & REVENUE REVERSALS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Date"
        .Cells(auditRow, 2).Value = "JE Number"
        .Cells(auditRow, 3).Value = "Description"
        .Cells(auditRow, 4).Value = "Amount"
        .Cells(auditRow, 5).Value = "Review Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim creditStart As Long
        creditStart = auditRow

        Dim creditTotal As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 1) = "4" Then
                ' Credit memos = debits to revenue
                If wsGL.Cells(i, 6).Value > 0 Then
                    .Cells(auditRow, 1).Value = wsGL.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsGL.Cells(i, 2).Value
                    .Cells(auditRow, 3).Value = wsGL.Cells(i, 5).Value
                    .Cells(auditRow, 4).Value = wsGL.Cells(i, 6).Value * -1
                    .Cells(auditRow, 5).Value = "VERIFY REASON"
                    .Cells(auditRow, 5).Interior.Color = RGB(255, 235, 156)

                    creditTotal = creditTotal + wsGL.Cells(i, 6).Value
                    auditRow = auditRow + 1
                End If
            End If
        Next i

        ' Summary
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "Total Credit Memos:"
        .Cells(auditRow, 4).Value = creditTotal * -1
        .Cells(auditRow, 5).Value = creditTotal / totalRevenue
        .Cells(auditRow, 5).NumberFormat = "0.0%"

        .Range(.Cells(creditStart, 4), .Cells(auditRow, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' AUDIT SUMMARY
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "AUDIT SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Total Revenue:"
        .Cells(auditRow, 2).Value = totalRevenue
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Monthly revenue analytics"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Revenue by account analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Cutoff testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Large transaction review"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Credit memo analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Detail testing / vouching (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Contract review for ASC 606 (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion - address fraud risk considerations]"
        .Cells(auditRow, 1).Font.Italic = True

        ' Column widths
        .Columns("A").ColumnWidth = 25
        .Columns("B:H").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "Revenue Audit Complete!" & vbCrLf & vbCrLf & _
           "Total Revenue: " & Format(totalRevenue, "$#,##0") & vbCrLf & _
           "Review the Revenue_Audit worksheet.", vbInformation

End Sub
```

---

## Output Examples

### Revenue_Audit Worksheet

The `AuditRevenue` procedure generates a comprehensive worksheet:

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ REVENUE - AUDIT WORKPAPER                                                           │
│ Period: Year Ended December 31, 2024                                                │
│ Prepared: AUDITOR on 12/15/2024 3:30:00 PM                                         │
│ Materiality: $50,000                                                                │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ NOTE: Revenue is a PRESUMED FRAUD RISK per AU-C 240                     ▓▓▓ RED    │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

**TEST 1: MONTHLY REVENUE ANALYSIS**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 1: MONTHLY REVENUE ANALYSIS                                                                        │
├────────────┬────────────────┬────────────────┬────────────────┬────────────┬────────────┬───────────────┤
│ Month      │ CY Revenue     │ PY Revenue     │ $ Change       │ % Change   │ CY % Total │ Status        │
├────────────┼────────────────┼────────────────┼────────────────┼────────────┼────────────┼───────────────┤
│ January    │ $850,000       │ $780,000       │ $70,000        │ 9.0%       │ 7.8%       │ ✓ OK          │
│ February   │ $820,000       │ $750,000       │ $70,000        │ 9.3%       │ 7.5%       │ ✓ OK          │
│ March      │ $910,000       │ $800,000       │ $110,000       │ 13.8%      │ 8.4%       │ ✓ OK          │
│ April      │ $875,000       │ $820,000       │ $55,000        │ 6.7%       │ 8.0%       │ ✓ OK          │
│ May        │ $925,000       │ $850,000       │ $75,000        │ 8.8%       │ 8.5%       │ ✓ OK          │
│ June       │ $950,000       │ $875,000       │ $75,000        │ 8.6%       │ 8.7%       │ ✓ OK          │
│ July       │ $860,000       │ $810,000       │ $50,000        │ 6.2%       │ 7.9%       │ ✓ OK          │
│ August     │ $880,000       │ $830,000       │ $50,000        │ 6.0%       │ 8.1%       │ ✓ OK          │
│ September  │ $920,000       │ $860,000       │ $60,000        │ 7.0%       │ 8.5%       │ ✓ OK          │
│ October    │ $940,000       │ $890,000       │ $50,000        │ 5.6%       │ 8.6%       │ ✓ OK          │
│ November   │ $980,000       │ $920,000       │ $60,000        │ 6.5%       │ 9.0%       │ ✓ OK          │
│ December   │ $1,090,000     │ $850,000       │ $240,000       │ 28.2%      │ 10.0%      │ ⚠ INVESTIGATE │
├────────────┼────────────────┼────────────────┼────────────────┼────────────┼────────────┼───────────────┤
│ TOTAL      │ $10,900,000    │ $10,035,000    │ $865,000       │ 8.6%       │ 100.0%     │               │
└────────────┴────────────────┴────────────────┴────────────────┴────────────┴────────────┴───────────────┘
  Status: ⚠ INVESTIGATE = Variance > 15%
```

**TEST 2: REVENUE BY ACCOUNT**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 2: REVENUE BY ACCOUNT                                                          │
├─────────────┬────────────────────────────────┬────────────────┬─────────────────────┤
│ Account     │ Account Name                   │ Balance        │ % of Total          │
├─────────────┼────────────────────────────────┼────────────────┼─────────────────────┤
│ 4100        │ Product Revenue                │ $7,850,000     │ 72.0%               │
│ 4200        │ Service Revenue                │ $2,400,000     │ 22.0%               │
│ 4300        │ Licensing Revenue              │ $500,000       │ 4.6%                │
│ 4900        │ Other Revenue                  │ $150,000       │ 1.4%                │
├─────────────┴────────────────────────────────┼────────────────┼─────────────────────┤
│ TOTAL                                        │ $10,900,000    │ 100.0%              │
└──────────────────────────────────────────────┴────────────────┴─────────────────────┘
```

**TEST 3: REVENUE CUTOFF TESTING**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 3: REVENUE CUTOFF TESTING                                                                              │
│ Testing sales within 5 days of year-end for proper cutoff                                                   │
├────────────────┬────────────┬─────────────────────┬────────────┬────────────┬─────────┬─────────────────────┤
│ Invoice #      │ Inv Date   │ Customer            │ Ship Date  │ Amount     │Days frYE│ Cutoff Issue?       │
├────────────────┼────────────┼─────────────────────┼────────────┼────────────┼─────────┼─────────────────────┤
│ INV-2024-5847  │ 12/28/2024 │ Acme Corp           │ 12/28/2024 │ $45,000    │ -3      │ ✓ OK                │
│ INV-2024-5848  │ 12/29/2024 │ Beta Industries     │ 12/30/2024 │ $32,000    │ -2      │ ✓ OK                │
│ INV-2024-5849  │ 12/30/2024 │ Gamma LLC           │ 01/02/2025 │ $58,000    │ -1      │ ✗ SHIP AFTER Y/E    │
│ INV-2024-5850  │ 12/31/2024 │ Delta Corp          │ 12/31/2024 │ $27,500    │ 0       │ ✓ OK                │
│ INV-2025-0001  │ 01/02/2025 │ Epsilon Inc         │ 12/30/2024 │ $41,000    │ +2      │ ✗ SHIP BEFORE Y/E   │
│ INV-2025-0002  │ 01/03/2025 │ Zeta Partners       │ 01/03/2025 │ $19,000    │ +3      │ ✓ OK - Next year    │
└────────────────┴────────────┴─────────────────────┴────────────┴────────────┴─────────┴─────────────────────┘
  ✗ = Potential cutoff error requiring adjustment
```

**TEST 4: LARGE & UNUSUAL REVENUE TRANSACTIONS**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 4: LARGE & UNUSUAL REVENUE TRANSACTIONS                                                                │
├────────────┬────────────────┬────────────────────────────────────────┬────────────┬──────────┬──────────────┤
│ Date       │ JE Number      │ Description                            │ Amount     │% of Matl │ Action       │
├────────────┼────────────────┼────────────────────────────────────────┼────────────┼──────────┼──────────────┤
│ 12/15/2024 │ JE-2024-1234   │ Large contract - XYZ Corp              │ $185,000   │ 370%     │ ✗ MATERIAL   │
│ 12/28/2024 │ JE-2024-1298   │ Year-end sale - ABC Industries         │ $125,000   │ 250%     │ ✗ MATERIAL   │
│ 06/30/2024 │ JE-2024-0876   │ Equipment sale - one-time              │ $75,000    │ 150%     │ ✗ MATERIAL   │
│ 09/15/2024 │ JE-2024-1045   │ Settlement revenue                     │ $42,000    │ 84%      │ ⚠ Review     │
└────────────┴────────────────┴────────────────────────────────────────┴────────────┴──────────┴──────────────┘
  ✗ MATERIAL = Vouch to supporting documentation
```

**TEST 5: CREDIT MEMOS & REVENUE REVERSALS**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 5: CREDIT MEMOS & REVENUE REVERSALS                                                                    │
├────────────┬────────────────┬────────────────────────────────────────┬────────────┬─────────────────────────┤
│ Date       │ JE Number      │ Description                            │ Amount     │ Review Status           │
├────────────┼────────────────┼────────────────────────────────────────┼────────────┼─────────────────────────┤
│ 02/15/2024 │ CM-2024-0045   │ Customer return - defective product    │ ($12,500)  │ ⚠ VERIFY REASON         │
│ 05/22/2024 │ CM-2024-0156   │ Pricing adjustment per contract        │ ($8,750)   │ ⚠ VERIFY REASON         │
│ 08/30/2024 │ CM-2024-0287   │ Quantity dispute resolution            │ ($15,000)  │ ⚠ VERIFY REASON         │
│ 11/14/2024 │ CM-2024-0398   │ Service credit - quality issue         │ ($6,200)   │ ⚠ VERIFY REASON         │
├────────────┴────────────────┴────────────────────────────────────────┼────────────┼─────────────────────────┤
│ Total Credit Memos:                                                  │ ($42,450)  │ 0.4% of Revenue         │
└──────────────────────────────────────────────────────────────────────┴────────────┴─────────────────────────┘
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                               │
├─────────────────────────────────────────────────────────────────────────────┤
│ Total Revenue: $10,900,000                                                  │
├─────────────────────────────────────────────────────────────────────────────┤
│ Procedures Performed:                                                       │
│   ✓ Monthly revenue analytics                                               │
│   ✓ Revenue by account analysis                                             │
│   ✓ Cutoff testing                                                          │
│   ✓ Large transaction review                                                │
│   ✓ Credit memo analysis                                                    │
│   ☐ Detail testing / vouching (manual)                                      │
│   ☐ Contract review for ASC 606 (manual)                                    │
├─────────────────────────────────────────────────────────────────────────────┤
│ CONCLUSION: [Document conclusion - address fraud risk considerations]       │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## ASC 606 - Five-Step Model

Revenue recognition under ASC 606:

| Step | Description | Audit Consideration |
|------|-------------|---------------------|
| 1 | Identify the contract | Review significant contracts |
| 2 | Identify performance obligations | Multiple deliverables? |
| 3 | Determine transaction price | Variable consideration? |
| 4 | Allocate to obligations | Standalone selling prices |
| 5 | Recognize when satisfied | Point in time vs. over time |

---

## Assertions Tested

| Assertion | Test | Pass Criteria |
|-----------|------|---------------|
| **Occurrence** | Vouch to support, analytics | Transaction happened |
| **Completeness** | Analytical procedures | All revenue recorded |
| **Accuracy** | Recalculation, invoice tie | Amounts correct |
| **Cutoff** | Ship dates vs. invoice dates | Correct period |
| **Classification** | Review account coding | Proper accounts |

---

## Fraud Considerations

Revenue is a **presumed fraud risk** per AU-C 240. Consider:

| Red Flag | Audit Response |
|----------|----------------|
| Unusual revenue near period-end | Cutoff testing |
| Round dollar transactions | Vouch to support |
| Revenue without cash collection | Subsequent receipts |
| Related party sales | Confirm, review pricing |
| Channel stuffing | Customer confirmations |
| Bill-and-hold arrangements | Verify criteria met |

---

[⬅️ Back to FS Auditing](../README.md) | [➡️ COGS](./cogs.md)
