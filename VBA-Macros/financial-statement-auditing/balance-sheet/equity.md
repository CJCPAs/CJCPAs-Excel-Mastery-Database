# Stockholders' Equity Audit VBA

> **Equity Testing** - Complete VBA for auditing stockholders' equity per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 3000-3999 (typically) |
| **Assertions** | Existence, Completeness, Valuation, Rights, Presentation |
| **Risk Level** | Low-Moderate (complex transactions) |
| **Key Standards** | ASC 505 (Equity), ASC 718 (Stock Compensation) |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for equity accounts

### Input Sheet 2: `Equity_Schedule`
Equity rollforward schedule

| Column | Header | Example |
|--------|--------|---------|
| A | `Account` | 3100 |
| B | `Description` | Common Stock |
| C | `PY_Balance` | 100000 |
| D | `Issuances` | 25000 |
| E | `Repurchases` | 0 |
| F | `Dividends` | 0 |
| G | `Net_Income` | 0 |
| H | `Other` | 0 |
| I | `CY_Balance` | 125000 |

### Input Sheet 3: `Stock_Transactions`
Detail of stock transactions during the year

| Column | Header | Example |
|--------|--------|---------|
| A | `Date` | 03/15/2024 |
| B | `Type` | Issuance |
| C | `Shares` | 10000 |
| D | `Par_Value` | 1 |
| E | `Issue_Price` | 25 |
| F | `Total_Proceeds` | 250000 |
| G | `Common_Stock` | 10000 |
| H | `APIC` | 240000 |
| I | `Board_Approval` | 03/01/2024 |

---

## Audit Procedures

```vba
Sub AuditEquity()
    '================================================
    ' STOCKHOLDERS' EQUITY - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with equity transactions
    '   - Sheet "Equity_Schedule" with rollforward
    '   - Sheet "Stock_Transactions" with transaction detail
    '
    ' OUTPUTS:
    '   - Creates "Equity_Audit" worksheet
    '   - Reconciles schedule to GL
    '   - Tests equity rollforward
    '   - Verifies stock transaction accounting
    '
    ' ASSERTIONS TESTED:
    '   - Existence (equity accounts exist)
    '   - Completeness (all transactions recorded)
    '   - Valuation (amounts accurate)
    '   - Rights (proper authorization)
    '================================================

    Dim wsGL As Worksheet
    Dim wsEquity As Worksheet
    Dim wsTrans As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsEquity = ThisWorkbook.Sheets("Equity_Schedule")
    Set wsTrans = ThisWorkbook.Sheets("Stock_Transactions")
    On Error GoTo 0

    If wsGL Is Nothing Or wsEquity Is Nothing Then
        MsgBox "GL_Detail and Equity_Schedule sheets required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Equity_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Equity_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "STOCKHOLDERS' EQUITY - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: EQUITY SCHEDULE TO GL RECONCILIATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: EQUITY SCHEDULE TO GL RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        ' Build dictionary of GL balances by account
        Dim glDict As Object
        Set glDict = CreateObject("Scripting.Dictionary")

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            Dim acctBal As Double

            acctNum = CStr(wsGL.Cells(i, 3).Value)

            ' Equity accounts (3xxx) - credit balance
            If Left(acctNum, 1) = "3" Then
                acctBal = wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value

                If glDict.Exists(acctNum) Then
                    glDict(acctNum) = glDict(acctNum) + acctBal
                Else
                    glDict.Add acctNum, acctBal
                End If
            End If
        Next i

        .Cells(auditRow, 1).Value = "Account"
        .Cells(auditRow, 2).Value = "Description"
        .Cells(auditRow, 3).Value = "GL Balance"
        .Cells(auditRow, 4).Value = "Schedule"
        .Cells(auditRow, 5).Value = "Difference"
        .Cells(auditRow, 6).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        Dim reconStart As Long
        reconStart = auditRow

        Dim totalGLEquity As Double, totalSLEquity As Double

        lastRow = wsEquity.Cells(wsEquity.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim eqAcct As String
            Dim eqDesc As String
            Dim eqBal As Double
            Dim glBal As Double

            eqAcct = CStr(wsEquity.Cells(i, 1).Value)
            eqDesc = wsEquity.Cells(i, 2).Value
            eqBal = wsEquity.Cells(i, 9).Value

            glBal = 0
            If glDict.Exists(eqAcct) Then
                glBal = glDict(eqAcct)
            End If

            .Cells(auditRow, 1).Value = eqAcct
            .Cells(auditRow, 2).Value = eqDesc
            .Cells(auditRow, 3).Value = glBal
            .Cells(auditRow, 4).Value = eqBal
            .Cells(auditRow, 5).Value = glBal - eqBal

            If Abs(glBal - eqBal) < 1 Then
                .Cells(auditRow, 6).Value = "RECONCILED"
                .Cells(auditRow, 6).Interior.Color = RGB(198, 239, 206)
            Else
                .Cells(auditRow, 6).Value = "DIFFERENCE"
                .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
            End If

            totalGLEquity = totalGLEquity + glBal
            totalSLEquity = totalSLEquity + eqBal
            auditRow = auditRow + 1
        Next i

        ' Total
        .Cells(auditRow, 1).Value = "TOTAL EQUITY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 3).Value = totalGLEquity
        .Cells(auditRow, 3).Font.Bold = True
        .Cells(auditRow, 4).Value = totalSLEquity
        .Cells(auditRow, 4).Font.Bold = True
        .Cells(auditRow, 5).Value = totalGLEquity - totalSLEquity
        auditRow = auditRow + 1

        .Range(.Cells(reconStart, 3), .Cells(auditRow - 1, 5)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 2: EQUITY ROLLFORWARD
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 2: EQUITY ROLLFORWARD"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Account"
        .Cells(auditRow, 2).Value = "PY Bal"
        .Cells(auditRow, 3).Value = "Issuances"
        .Cells(auditRow, 4).Value = "Repurchases"
        .Cells(auditRow, 5).Value = "Dividends"
        .Cells(auditRow, 6).Value = "Net Inc"
        .Cells(auditRow, 7).Value = "Other"
        .Cells(auditRow, 8).Value = "Expected"
        .Cells(auditRow, 9).Value = "CY Bal"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Font.Bold = True
        auditRow = auditRow + 1

        Dim rollStart As Long
        rollStart = auditRow

        lastRow = wsEquity.Cells(wsEquity.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim pyBal As Double
            Dim issues As Double
            Dim repurch As Double
            Dim divs As Double
            Dim netInc As Double
            Dim other As Double
            Dim cyBal As Double
            Dim expected As Double

            pyBal = wsEquity.Cells(i, 3).Value
            issues = wsEquity.Cells(i, 4).Value
            repurch = wsEquity.Cells(i, 5).Value
            divs = wsEquity.Cells(i, 6).Value
            netInc = wsEquity.Cells(i, 7).Value
            other = wsEquity.Cells(i, 8).Value
            cyBal = wsEquity.Cells(i, 9).Value

            expected = pyBal + issues - repurch - divs + netInc + other

            .Cells(auditRow, 1).Value = wsEquity.Cells(i, 2).Value
            .Cells(auditRow, 2).Value = pyBal
            .Cells(auditRow, 3).Value = issues
            .Cells(auditRow, 4).Value = repurch
            .Cells(auditRow, 5).Value = divs
            .Cells(auditRow, 6).Value = netInc
            .Cells(auditRow, 7).Value = other
            .Cells(auditRow, 8).Value = expected
            .Cells(auditRow, 9).Value = cyBal

            ' Color code if expected doesn't match actual
            If Abs(expected - cyBal) > 1 Then
                .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
                .Cells(auditRow, 9).Interior.Color = RGB(255, 199, 206)
            End If

            auditRow = auditRow + 1
        Next i

        .Range(.Cells(rollStart, 2), .Cells(auditRow - 1, 9)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 3: STOCK TRANSACTION TESTING
        ' ========================================
        If Not wsTrans Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 3: STOCK TRANSACTION TESTING"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Date"
            .Cells(auditRow, 2).Value = "Type"
            .Cells(auditRow, 3).Value = "Shares"
            .Cells(auditRow, 4).Value = "Price"
            .Cells(auditRow, 5).Value = "Total"
            .Cells(auditRow, 6).Value = "C/S"
            .Cells(auditRow, 7).Value = "APIC"
            .Cells(auditRow, 8).Value = "Calc Check"
            .Cells(auditRow, 9).Value = "Auth?"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Font.Bold = True
            auditRow = auditRow + 1

            Dim transStart As Long
            transStart = auditRow

            lastRow = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim trShares As Double
                Dim trPar As Double
                Dim trPrice As Double
                Dim trTotal As Double
                Dim trCS As Double
                Dim trAPIC As Double
                Dim calcCS As Double
                Dim calcAPIC As Double

                trShares = wsTrans.Cells(i, 3).Value
                trPar = wsTrans.Cells(i, 4).Value
                trPrice = wsTrans.Cells(i, 5).Value
                trTotal = wsTrans.Cells(i, 6).Value
                trCS = wsTrans.Cells(i, 7).Value
                trAPIC = wsTrans.Cells(i, 8).Value

                calcCS = trShares * trPar
                calcAPIC = trTotal - calcCS

                .Cells(auditRow, 1).Value = wsTrans.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsTrans.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = trShares
                .Cells(auditRow, 4).Value = trPrice
                .Cells(auditRow, 5).Value = trTotal
                .Cells(auditRow, 6).Value = trCS
                .Cells(auditRow, 7).Value = trAPIC

                ' Verify calculation
                If Abs(trCS - calcCS) < 1 And Abs(trAPIC - calcAPIC) < 1 Then
                    .Cells(auditRow, 8).Value = "VERIFIED"
                    .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
                Else
                    .Cells(auditRow, 8).Value = "ERROR"
                    .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
                End If

                ' Check board approval
                If IsDate(wsTrans.Cells(i, 9).Value) Then
                    .Cells(auditRow, 9).Value = "YES"
                    .Cells(auditRow, 9).Interior.Color = RGB(198, 239, 206)
                Else
                    .Cells(auditRow, 9).Value = "NO"
                    .Cells(auditRow, 9).Interior.Color = RGB(255, 199, 206)
                End If

                auditRow = auditRow + 1
            Next i

            .Range(.Cells(transStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "#,##0"
            .Range(.Cells(transStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "$#,##0.00"
            .Range(.Cells(transStart, 5), .Cells(auditRow - 1, 7)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 2
        End If

        ' ========================================
        ' TEST 4: RETAINED EARNINGS RECONCILIATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: RETAINED EARNINGS RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Beginning Retained Earnings"
        .Cells(auditRow, 2).Value = "[Enter from PY financials]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Add: Net Income"
        .Cells(auditRow, 2).Value = "[Enter from IS]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Less: Dividends Declared"
        .Cells(auditRow, 2).Value = "[Enter]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Less: Treasury Stock"
        .Cells(auditRow, 2).Value = "[Enter]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Other Adjustments"
        .Cells(auditRow, 2).Value = "[Enter]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Calculated Ending RE"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Formula]"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Per GL"
        .Cells(auditRow, 2).Value = "[Enter]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Difference"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Formula]"
        auditRow = auditRow + 1

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

        .Cells(auditRow, 1).Value = "Total Stockholders' Equity:"
        .Cells(auditRow, 2).Value = totalSLEquity
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Schedule to GL reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Equity rollforward testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Stock transaction testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Board minutes review (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Stock certificate inspection (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Transfer agent confirmation (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 30
        .Columns("B:I").ColumnWidth = 14

    End With

    Application.ScreenUpdating = True

    MsgBox "Equity Audit Complete!" & vbCrLf & _
           "Total Equity: " & Format(totalSLEquity, "$#,##0"), vbInformation

End Sub
```

---

## Dividend Testing

```vba
Sub TestDividends()
    '================================================
    ' DIVIDEND TESTING
    '
    ' Tests dividend declarations and payments for:
    '   - Proper authorization
    '   - Correct calculation
    '   - Proper recording
    '================================================

    Dim wsAudit As Worksheet
    Dim auditRow As Long

    On Error Resume Next
    ThisWorkbook.Sheets("Dividend_Testing").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Dividend_Testing"

    With wsAudit
        .Range("A1").Value = "DIVIDEND TESTING WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")

        auditRow = 4

        ' Dividend declarations
        .Cells(auditRow, 1).Value = "DIVIDEND DECLARATIONS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Declaration Date"
        .Cells(auditRow, 2).Value = "Record Date"
        .Cells(auditRow, 3).Value = "Payment Date"
        .Cells(auditRow, 4).Value = "Type"
        .Cells(auditRow, 5).Value = "Per Share"
        .Cells(auditRow, 6).Value = "Shares O/S"
        .Cells(auditRow, 7).Value = "Total Div"
        .Cells(auditRow, 8).Value = "Board Auth"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Font.Bold = True
        auditRow = auditRow + 1

        ' Sample rows for input
        Dim r As Long
        For r = 1 To 4
            .Cells(auditRow, 1).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 4).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 5).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 6).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 7).Value = "=E" & auditRow & "*F" & auditRow
            .Cells(auditRow, 8).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next r

        auditRow = auditRow + 2

        ' Testing checklist
        .Cells(auditRow, 1).Value = "DIVIDEND TESTING CHECKLIST"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 2

        Dim checks As Variant
        checks = Array( _
            "Board approved dividend declaration (review minutes)", _
            "Dividend per share agrees to board resolution", _
            "Record date and payment date agree to resolution", _
            "Shares outstanding at record date verified", _
            "Total dividend calculation verified", _
            "Dividend payable recorded at declaration date", _
            "Cash payment recorded at payment date", _
            "Retained earnings properly reduced", _
            "No dividends in violation of debt covenants", _
            "Adequate retained earnings for dividend")

        Dim chk As Variant
        For Each chk In checks
            .Cells(auditRow, 1).Value = ChrW(9744) & " " & chk
            .Cells(auditRow, 3).Value = "[Y/N/NA]"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next chk

        .Columns("A").ColumnWidth = 50
        .Columns("B:H").ColumnWidth = 14

    End With

    MsgBox "Dividend Testing Template Created!", vbInformation

End Sub
```

---

## Stock Compensation (ASC 718)

```vba
Sub TestStockCompensation()
    '================================================
    ' STOCK COMPENSATION TESTING (ASC 718)
    '
    ' For companies with stock options or restricted stock
    '================================================

    Dim wsAudit As Worksheet
    Dim auditRow As Long

    On Error Resume Next
    ThisWorkbook.Sheets("StockComp_Testing").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "StockComp_Testing"

    With wsAudit
        .Range("A1").Value = "STOCK COMPENSATION TESTING (ASC 718)"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")

        auditRow = 4

        ' Stock option grants
        .Cells(auditRow, 1).Value = "STOCK OPTION GRANTS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Grant Date"
        .Cells(auditRow, 2).Value = "Grantee"
        .Cells(auditRow, 3).Value = "Options"
        .Cells(auditRow, 4).Value = "Strike"
        .Cells(auditRow, 5).Value = "FV/Option"
        .Cells(auditRow, 6).Value = "Total FV"
        .Cells(auditRow, 7).Value = "Vesting"
        .Cells(auditRow, 8).Value = "CY Expense"
        .Cells(auditRow, 9).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Font.Bold = True
        auditRow = auditRow + 1

        ' Sample rows
        For r = 1 To 5
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next r

        auditRow = auditRow + 2

        ' Valuation inputs
        .Cells(auditRow, 1).Value = "OPTION VALUATION INPUTS (Black-Scholes)"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        Dim inputs As Variant
        inputs = Array( _
            Array("Stock price at grant date", "$", "Stock price"), _
            Array("Exercise price", "$", "Per agreement"), _
            Array("Expected term (years)", "", "SAB 110"), _
            Array("Risk-free rate", "%", "Treasury rate"), _
            Array("Expected volatility", "%", "Historical"), _
            Array("Expected dividends", "%", "Dividend policy"))

        Dim inp As Variant
        For Each inp In inputs
            .Cells(auditRow, 1).Value = inp(0)
            .Cells(auditRow, 2).Value = "[Input]"
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 3).Value = inp(2)
            auditRow = auditRow + 1
        Next inp

        auditRow = auditRow + 2

        ' Testing checklist
        .Cells(auditRow, 1).Value = "ASC 718 COMPLIANCE CHECKLIST"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 2

        checks = Array( _
            "Grant date properly determined", _
            "Fair value measured at grant date", _
            "Valuation model appropriate (Black-Scholes/Binomial)", _
            "Volatility assumption supportable", _
            "Expected term assumption reasonable", _
            "Risk-free rate from zero-coupon Treasury", _
            "Forfeitures estimated or actual method", _
            "Expense recognized over requisite service period", _
            "Modifications properly accounted for", _
            "Disclosures complete per ASC 718-10-50")

        For Each chk In checks
            .Cells(auditRow, 1).Value = ChrW(9744) & " " & chk
            .Cells(auditRow, 3).Value = "[Y/N/NA]"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next chk

        .Columns("A").ColumnWidth = 45
        .Columns("B:I").ColumnWidth = 12

    End With

    MsgBox "Stock Compensation Testing Template Created!", vbInformation

End Sub
```

---

## Key Equity Components

| Component | Audit Focus |
|-----------|------------|
| **Common Stock** | Par value, authorized vs. issued shares |
| **Preferred Stock** | Terms, liquidation preference, conversion |
| **APIC** | Premium over par calculation |
| **Retained Earnings** | Rollforward, dividend restrictions |
| **Treasury Stock** | Cost vs. par method, reissuances |
| **AOCI** | Pension, FX, available-for-sale securities |
| **NCI** | Subsidiary ownership changes |

---

## Documents to Obtain

- Articles of incorporation (authorized shares)
- Board minutes (dividends, stock issuances)
- Stock certificate books or transfer agent reports
- Stock option plans and grant agreements
- Treasury stock transaction support
- Prior year audited financial statements

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Debt](./debt.md)
