# Debt & Notes Payable Audit VBA

> **Debt Testing** - Complete VBA for auditing long-term debt and notes payable per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 2200-2399 (short-term), 2500-2699 (long-term) |
| **Assertions** | Completeness, Valuation, Rights/Obligations, Classification |
| **Risk Level** | Moderate-High (covenants, classification) |
| **Key Standards** | ASC 470 (Debt), ASC 835 (Interest) |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for debt and interest accounts

### Input Sheet 2: `Debt_Schedule`
Complete debt schedule with all borrowings

| Column | Header | Example |
|--------|--------|---------|
| A | `Lender` | First National Bank |
| B | `Loan_Type` | Term Loan |
| C | `Original_Amount` | 5000000 |
| D | `Interest_Rate` | 0.065 |
| E | `Origination_Date` | 01/15/2020 |
| F | `Maturity_Date` | 01/15/2027 |
| G | `Monthly_Payment` | 75000 |
| H | `PY_Balance` | 3750000 |
| I | `CY_Balance` | 3250000 |
| J | `Current_Portion` | 450000 |
| K | `Collateral` | Real Estate |
| L | `Covenant_Type` | Debt Service |

### Input Sheet 3: `Interest_Payments`
Interest payments made during the year

| Column | Header | Example |
|--------|--------|---------|
| A | `Date` | 01/15/2024 |
| B | `Lender` | First National Bank |
| C | `Payment_Amount` | 75000 |
| D | `Principal` | 54000 |
| E | `Interest` | 21000 |
| F | `Reference` | Check #4521 |

---

## Audit Procedures

```vba
Sub AuditDebt()
    '================================================
    ' DEBT & NOTES PAYABLE - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with debt transactions
    '   - Sheet "Debt_Schedule" with loan details
    '   - Sheet "Interest_Payments" with payment history
    '
    ' OUTPUTS:
    '   - Creates "Debt_Audit" worksheet
    '   - Reconciles schedule to GL
    '   - Recalculates interest expense
    '   - Tests current/long-term classification
    '   - Generates confirmation requests
    '
    ' ASSERTIONS TESTED:
    '   - Completeness (all debt recorded)
    '   - Valuation (balances accurate)
    '   - Classification (current vs long-term)
    '   - Rights/Obligations (terms verified)
    '================================================

    Dim wsGL As Worksheet
    Dim wsDebt As Worksheet
    Dim wsInt As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const MATERIALITY As Double = 50000

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsDebt = ThisWorkbook.Sheets("Debt_Schedule")
    Set wsInt = ThisWorkbook.Sheets("Interest_Payments")
    On Error GoTo 0

    If wsGL Is Nothing Or wsDebt Is Nothing Then
        MsgBox "GL_Detail and Debt_Schedule sheets required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Debt_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Debt_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "DEBT & NOTES PAYABLE - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: DEBT SCHEDULE TO GL RECONCILIATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: DEBT SCHEDULE TO GL RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        ' Calculate GL balances
        Dim glSTDebt As Double   ' Short-term (22xx-23xx)
        Dim glLTDebt As Double   ' Long-term (25xx-26xx)
        Dim slSTDebt As Double
        Dim slLTDebt As Double
        Dim slTotalDebt As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            acctNum = CStr(wsGL.Cells(i, 3).Value)

            ' Short-term debt (22xx, 23xx) - credit balance
            If Left(acctNum, 2) = "22" Or Left(acctNum, 2) = "23" Then
                glSTDebt = glSTDebt + wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value
            End If

            ' Long-term debt (25xx, 26xx) - credit balance
            If Left(acctNum, 2) = "25" Or Left(acctNum, 2) = "26" Then
                glLTDebt = glLTDebt + wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value
            End If
        Next i

        ' Calculate subledger balances
        lastRow = wsDebt.Cells(wsDebt.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim cyBal As Double
            Dim curPortion As Double

            cyBal = wsDebt.Cells(i, 9).Value
            curPortion = wsDebt.Cells(i, 10).Value

            slSTDebt = slSTDebt + curPortion
            slLTDebt = slLTDebt + (cyBal - curPortion)
            slTotalDebt = slTotalDebt + cyBal
        Next i

        .Cells(auditRow, 1).Value = "Category"
        .Cells(auditRow, 2).Value = "GL Balance"
        .Cells(auditRow, 3).Value = "Schedule"
        .Cells(auditRow, 4).Value = "Difference"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim reconStart As Long
        reconStart = auditRow

        ' Short-term debt
        .Cells(auditRow, 1).Value = "Current Portion of LT Debt"
        .Cells(auditRow, 2).Value = glSTDebt
        .Cells(auditRow, 3).Value = slSTDebt
        .Cells(auditRow, 4).Value = glSTDebt - slSTDebt

        If Abs(glSTDebt - slSTDebt) < 1 Then
            .Cells(auditRow, 5).Value = "RECONCILED"
            .Cells(auditRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 5).Value = "DIFFERENCE"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Long-term debt
        .Cells(auditRow, 1).Value = "Long-Term Debt"
        .Cells(auditRow, 2).Value = glLTDebt
        .Cells(auditRow, 3).Value = slLTDebt
        .Cells(auditRow, 4).Value = glLTDebt - slLTDebt

        If Abs(glLTDebt - slLTDebt) < 1 Then
            .Cells(auditRow, 5).Value = "RECONCILED"
            .Cells(auditRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 5).Value = "DIFFERENCE"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Total
        .Cells(auditRow, 1).Value = "TOTAL DEBT"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glSTDebt + glLTDebt
        .Cells(auditRow, 2).Font.Bold = True
        .Cells(auditRow, 3).Value = slTotalDebt
        .Cells(auditRow, 3).Font.Bold = True
        .Cells(auditRow, 4).Value = (glSTDebt + glLTDebt) - slTotalDebt
        auditRow = auditRow + 1

        .Range(.Cells(reconStart, 2), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 2: DEBT ROLLFORWARD
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 2: DEBT ROLLFORWARD BY LENDER"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Lender"
        .Cells(auditRow, 2).Value = "PY Balance"
        .Cells(auditRow, 3).Value = "Additions"
        .Cells(auditRow, 4).Value = "Principal Paid"
        .Cells(auditRow, 5).Value = "Expected CY"
        .Cells(auditRow, 6).Value = "Actual CY"
        .Cells(auditRow, 7).Value = "Difference"
        .Cells(auditRow, 8).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Font.Bold = True
        auditRow = auditRow + 1

        Dim rollStart As Long
        rollStart = auditRow

        Dim totalPY As Double, totalCY As Double

        lastRow = wsDebt.Cells(wsDebt.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim lender As String
            Dim pyBal As Double
            Dim actualCY As Double
            Dim monthlyPmt As Double
            Dim principalPaid As Double
            Dim expectedCY As Double

            lender = wsDebt.Cells(i, 1).Value
            pyBal = wsDebt.Cells(i, 8).Value
            actualCY = wsDebt.Cells(i, 9).Value
            monthlyPmt = wsDebt.Cells(i, 7).Value

            ' Calculate principal paid from interest payments sheet
            principalPaid = 0
            If Not wsInt Is Nothing Then
                Dim intLastRow As Long
                intLastRow = wsInt.Cells(wsInt.Rows.Count, "A").End(xlUp).Row
                Dim j As Long
                For j = 2 To intLastRow
                    If wsInt.Cells(j, 2).Value = lender Then
                        principalPaid = principalPaid + wsInt.Cells(j, 4).Value
                    End If
                Next j
            End If

            expectedCY = pyBal - principalPaid

            .Cells(auditRow, 1).Value = lender
            .Cells(auditRow, 2).Value = pyBal
            .Cells(auditRow, 3).Value = 0  ' New borrowings (manual input)
            .Cells(auditRow, 4).Value = principalPaid
            .Cells(auditRow, 5).Value = expectedCY
            .Cells(auditRow, 6).Value = actualCY
            .Cells(auditRow, 7).Value = expectedCY - actualCY

            If Abs(expectedCY - actualCY) < 100 Then
                .Cells(auditRow, 8).Value = "RECONCILED"
                .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
            Else
                .Cells(auditRow, 8).Value = "INVESTIGATE"
                .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
            End If

            totalPY = totalPY + pyBal
            totalCY = totalCY + actualCY
            auditRow = auditRow + 1
        Next i

        .Range(.Cells(rollStart, 2), .Cells(auditRow - 1, 7)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 3: INTEREST EXPENSE RECALCULATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 3: INTEREST EXPENSE RECALCULATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Lender"
        .Cells(auditRow, 2).Value = "Avg Balance"
        .Cells(auditRow, 3).Value = "Rate"
        .Cells(auditRow, 4).Value = "Expected Int"
        .Cells(auditRow, 5).Value = "Actual Int"
        .Cells(auditRow, 6).Value = "Difference"
        .Cells(auditRow, 7).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
        auditRow = auditRow + 1

        Dim intStart As Long
        intStart = auditRow

        Dim totalExpectedInt As Double, totalActualInt As Double

        lastRow = wsDebt.Cells(wsDebt.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim avgBalance As Double
            Dim intRate As Double
            Dim expectedInt As Double
            Dim actualInt As Double

            lender = wsDebt.Cells(i, 1).Value
            pyBal = wsDebt.Cells(i, 8).Value
            cyBal = wsDebt.Cells(i, 9).Value
            intRate = wsDebt.Cells(i, 4).Value

            avgBalance = (pyBal + cyBal) / 2
            expectedInt = avgBalance * intRate

            ' Sum actual interest from payments
            actualInt = 0
            If Not wsInt Is Nothing Then
                intLastRow = wsInt.Cells(wsInt.Rows.Count, "A").End(xlUp).Row
                For j = 2 To intLastRow
                    If wsInt.Cells(j, 2).Value = lender Then
                        actualInt = actualInt + wsInt.Cells(j, 5).Value
                    End If
                Next j
            End If

            .Cells(auditRow, 1).Value = lender
            .Cells(auditRow, 2).Value = avgBalance
            .Cells(auditRow, 3).Value = intRate
            .Cells(auditRow, 4).Value = expectedInt
            .Cells(auditRow, 5).Value = actualInt
            .Cells(auditRow, 6).Value = expectedInt - actualInt

            ' Allow 5% variance for timing
            If Abs(expectedInt - actualInt) / expectedInt < 0.05 Then
                .Cells(auditRow, 7).Value = "REASONABLE"
                .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
            Else
                .Cells(auditRow, 7).Value = "REVIEW"
                .Cells(auditRow, 7).Interior.Color = RGB(255, 235, 156)
            End If

            totalExpectedInt = totalExpectedInt + expectedInt
            totalActualInt = totalActualInt + actualInt
            auditRow = auditRow + 1
        Next i

        ' Total
        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 4).Value = totalExpectedInt
        .Cells(auditRow, 4).Font.Bold = True
        .Cells(auditRow, 5).Value = totalActualInt
        .Cells(auditRow, 5).Font.Bold = True
        .Cells(auditRow, 6).Value = totalExpectedInt - totalActualInt
        auditRow = auditRow + 1

        .Range(.Cells(intStart, 2), .Cells(auditRow - 1, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(intStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "0.00%"
        .Range(.Cells(intStart, 4), .Cells(auditRow - 1, 6)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 4: CURRENT PORTION CLASSIFICATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: CURRENT PORTION CLASSIFICATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Lender"
        .Cells(auditRow, 2).Value = "CY Balance"
        .Cells(auditRow, 3).Value = "Maturity"
        .Cells(auditRow, 4).Value = "Mths to Mat"
        .Cells(auditRow, 5).Value = "Est Current"
        .Cells(auditRow, 6).Value = "Per Books"
        .Cells(auditRow, 7).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
        auditRow = auditRow + 1

        Dim classStart As Long
        classStart = auditRow

        lastRow = wsDebt.Cells(wsDebt.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim matDate As Date
            Dim monthsToMat As Long
            Dim estCurrent As Double
            Dim perBooks As Double
            Dim moPayment As Double

            lender = wsDebt.Cells(i, 1).Value
            cyBal = wsDebt.Cells(i, 9).Value
            perBooks = wsDebt.Cells(i, 10).Value
            moPayment = wsDebt.Cells(i, 7).Value

            On Error Resume Next
            matDate = wsDebt.Cells(i, 6).Value
            On Error GoTo 0

            If IsDate(matDate) Then
                monthsToMat = DateDiff("m", DateSerial(Year(Date), 12, 31), matDate)

                ' If matures within 12 months, entire balance is current
                If monthsToMat <= 12 Then
                    estCurrent = cyBal
                Else
                    ' Estimate next 12 months of principal
                    estCurrent = moPayment * 12 * 0.7  ' Rough estimate (avg principal portion)
                End If
            Else
                monthsToMat = 0
                estCurrent = 0
            End If

            .Cells(auditRow, 1).Value = lender
            .Cells(auditRow, 2).Value = cyBal
            .Cells(auditRow, 3).Value = matDate
            .Cells(auditRow, 4).Value = monthsToMat
            .Cells(auditRow, 5).Value = estCurrent
            .Cells(auditRow, 6).Value = perBooks

            If Abs(estCurrent - perBooks) / cyBal < 0.1 Then
                .Cells(auditRow, 7).Value = "REASONABLE"
                .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
            Else
                .Cells(auditRow, 7).Value = "REVIEW CLASS"
                .Cells(auditRow, 7).Interior.Color = RGB(255, 235, 156)
            End If

            auditRow = auditRow + 1
        Next i

        .Range(.Cells(classStart, 2), .Cells(auditRow - 1, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(classStart, 5), .Cells(auditRow - 1, 6)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 3

        ' ========================================
        ' AUDIT SUMMARY
        ' ========================================
        .Cells(auditRow, 1).Value = "AUDIT SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Total Debt:"
        .Cells(auditRow, 2).Value = slTotalDebt
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  Current Portion:"
        .Cells(auditRow, 2).Value = slSTDebt
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  Long-Term:"
        .Cells(auditRow, 2).Value = slLTDebt
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Interest Expense:"
        .Cells(auditRow, 2).Value = totalActualInt
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Schedule to GL reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Debt rollforward testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Interest expense recalculation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Current portion classification"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Lender confirmations (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Covenant compliance (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 25
        .Columns("B:H").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "Debt Audit Complete!" & vbCrLf & _
           "Total Debt: " & Format(slTotalDebt, "$#,##0"), vbInformation

End Sub
```

---

## Debt Confirmation Generator

```vba
Sub GenerateDebtConfirmations()
    '================================================
    ' GENERATE DEBT CONFIRMATION REQUESTS
    '
    ' Creates confirmation letters for all lenders
    ' per AU-C 505 External Confirmation requirements
    '================================================

    Dim wsDebt As Worksheet
    Dim wsConf As Worksheet
    Dim lastRow As Long, i As Long
    Dim confRow As Long

    On Error Resume Next
    Set wsDebt = ThisWorkbook.Sheets("Debt_Schedule")
    On Error GoTo 0

    If wsDebt Is Nothing Then
        MsgBox "Debt_Schedule sheet required.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Sheets("Debt_Confirmations").Delete
    On Error GoTo 0

    Set wsConf = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsConf.Name = "Debt_Confirmations"

    With wsConf
        .Range("A1").Value = "DEBT CONFIRMATION CONTROL"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "As of: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")

        confRow = 4

        .Cells(confRow, 1).Value = "Lender"
        .Cells(confRow, 2).Value = "Balance"
        .Cells(confRow, 3).Value = "Date Sent"
        .Cells(confRow, 4).Value = "Date Rec'd"
        .Cells(confRow, 5).Value = "Confirmed Bal"
        .Cells(confRow, 6).Value = "Difference"
        .Cells(confRow, 7).Value = "Status"
        .Range(.Cells(confRow, 1), .Cells(confRow, 7)).Font.Bold = True
        .Range(.Cells(confRow, 1), .Cells(confRow, 7)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(confRow, 1), .Cells(confRow, 7)).Font.Color = RGB(255, 255, 255)
        confRow = confRow + 1

        lastRow = wsDebt.Cells(wsDebt.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            .Cells(confRow, 1).Value = wsDebt.Cells(i, 1).Value
            .Cells(confRow, 2).Value = wsDebt.Cells(i, 9).Value
            .Cells(confRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Cells(confRow, 7).Value = "PENDING"
            .Cells(confRow, 7).Interior.Color = RGB(255, 235, 156)
            confRow = confRow + 1
        Next i

        confRow = confRow + 2

        ' Confirmation template
        .Cells(confRow, 1).Value = "CONFIRMATION LETTER TEMPLATE"
        .Cells(confRow, 1).Font.Bold = True
        confRow = confRow + 2

        .Cells(confRow, 1).Value = "[Date]"
        confRow = confRow + 2
        .Cells(confRow, 1).Value = "[Lender Name]"
        confRow = confRow + 1
        .Cells(confRow, 1).Value = "[Address]"
        confRow = confRow + 2
        .Cells(confRow, 1).Value = "RE: Confirmation of Loan Balance"
        confRow = confRow + 2
        .Cells(confRow, 1).Value = "Dear Sir or Madam:"
        confRow = confRow + 2
        .Cells(confRow, 1).Value = "In connection with the audit of [Company Name], please confirm directly"
        confRow = confRow + 1
        .Cells(confRow, 1).Value = "to our auditors the following information as of [Date]:"
        confRow = confRow + 2
        .Cells(confRow, 1).Value = "  1. Outstanding principal balance"
        confRow = confRow + 1
        .Cells(confRow, 1).Value = "  2. Interest rate"
        confRow = confRow + 1
        .Cells(confRow, 1).Value = "  3. Maturity date"
        confRow = confRow + 1
        .Cells(confRow, 1).Value = "  4. Collateral pledged"
        confRow = confRow + 1
        .Cells(confRow, 1).Value = "  5. Financial covenant requirements and compliance status"
        confRow = confRow + 1
        .Cells(confRow, 1).Value = "  6. Any amounts past due"

        .Columns("A").ColumnWidth = 60
        .Columns("B:G").ColumnWidth = 15

    End With

    MsgBox "Debt Confirmation Control Sheet Created!", vbInformation

End Sub
```

---

## Covenant Compliance Testing

```vba
Sub TestDebtCovenants()
    '================================================
    ' DEBT COVENANT COMPLIANCE TESTING
    '
    ' Tests common financial covenants:
    '   - Debt service coverage ratio
    '   - Current ratio
    '   - Debt to equity ratio
    '   - Working capital minimum
    '================================================

    Dim wsAudit As Worksheet
    Dim auditRow As Long

    On Error Resume Next
    ThisWorkbook.Sheets("Covenant_Testing").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Covenant_Testing"

    With wsAudit
        .Range("A1").Value = "DEBT COVENANT COMPLIANCE TESTING"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "As of: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")

        auditRow = 4

        ' Input section
        .Cells(auditRow, 1).Value = "FINANCIAL DATA INPUTS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 3)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Metric"
        .Cells(auditRow, 2).Value = "Amount"
        .Cells(auditRow, 3).Value = "Source"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 3)).Font.Bold = True
        auditRow = auditRow + 1

        Dim inputs As Variant
        inputs = Array( _
            Array("EBITDA", 0, "Income Statement"), _
            Array("Annual Debt Service", 0, "Debt Schedule"), _
            Array("Current Assets", 0, "Balance Sheet"), _
            Array("Current Liabilities", 0, "Balance Sheet"), _
            Array("Total Debt", 0, "Debt Schedule"), _
            Array("Total Equity", 0, "Balance Sheet"), _
            Array("Working Capital", 0, "Calculated"))

        Dim inp As Variant
        For Each inp In inputs
            .Cells(auditRow, 1).Value = inp(0)
            .Cells(auditRow, 2).Value = inp(1)
            .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Cells(auditRow, 3).Value = inp(2)
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)  ' Yellow for input
            auditRow = auditRow + 1
        Next inp

        auditRow = auditRow + 2

        ' Covenant testing
        .Cells(auditRow, 1).Value = "COVENANT TESTING"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Covenant"
        .Cells(auditRow, 2).Value = "Required"
        .Cells(auditRow, 3).Value = "Actual"
        .Cells(auditRow, 4).Value = "Cushion"
        .Cells(auditRow, 5).Value = "Status"
        .Cells(auditRow, 6).Value = "Lender"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        Dim covenants As Variant
        covenants = Array( _
            Array("Debt Service Coverage", ">= 1.25x", "", "", ""), _
            Array("Current Ratio", ">= 1.50x", "", "", ""), _
            Array("Debt to Equity", "<= 3.00x", "", "", ""), _
            Array("Minimum Working Capital", ">= $500,000", "", "", ""))

        Dim cov As Variant
        For Each cov In covenants
            .Cells(auditRow, 1).Value = cov(0)
            .Cells(auditRow, 2).Value = cov(1)
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)  ' Yellow for input
            .Cells(auditRow, 5).Value = "[Calculate]"
            auditRow = auditRow + 1
        Next cov

        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "NOTES:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "1. Review loan agreements for exact covenant definitions"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "2. Confirm calculation methodology with lender if needed"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "3. Document any covenant waivers obtained"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "4. Consider going concern implications of covenant violations"

        .Columns("A").ColumnWidth = 30
        .Columns("B:F").ColumnWidth = 15

    End With

    MsgBox "Covenant Testing Template Created!" & vbCrLf & _
           "Enter financial data and covenant requirements.", vbInformation

End Sub
```

---

## Key Audit Considerations

| Consideration | Audit Response |
|--------------|----------------|
| **New Borrowings** | Obtain and read loan agreements |
| **Modifications** | Test for extinguishment vs. modification accounting |
| **Line of Credit** | Confirm availability and terms |
| **Related Party Loans** | Enhanced disclosure requirements |
| **Covenant Violations** | Evaluate going concern |
| **Classification** | Test current vs. long-term split |
| **Interest Capitalization** | Test if applicable (construction) |
| **Debt Issuance Costs** | Verify amortization |

---

## Output Examples

### Generated `Debt_Audit` Worksheet

**TEST 1: DEBT SCHEDULE TO GL RECONCILIATION**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ DEBT & NOTES PAYABLE - AUDIT WORKPAPER                                              │
│ Period: 12/31/2024                                                                  │
│ Materiality: $50,000                                                                │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ TEST 1: DEBT SCHEDULE TO GL RECONCILIATION                                          │
├────────────────────────────────┬────────────────┬────────────────┬──────────────────┤
│ Category                       │ GL Balance     │ Schedule       │ Status           │
├────────────────────────────────┼────────────────┼────────────────┼──────────────────┤
│ Current Portion of LT Debt     │ $450,000       │ $450,000       │ ✓ RECONCILED     │
│ Long-Term Debt                 │ $2,800,000     │ $2,800,000     │ ✓ RECONCILED     │
│ TOTAL DEBT                     │ $3,250,000     │ $3,250,000     │ ✓ RECONCILED     │
└────────────────────────────────┴────────────────┴────────────────┴──────────────────┘
```

**TEST 2: DEBT ROLLFORWARD BY LENDER**
```
┌──────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 2: DEBT ROLLFORWARD BY LENDER                                                                   │
├─────────────────────┬────────────┬────────────┬────────────┬────────────┬────────────┬──────────────┤
│ Lender              │ PY Balance │ Additions  │ Principal  │ Expected   │ Actual CY  │ Status       │
├─────────────────────┼────────────┼────────────┼────────────┼────────────┼────────────┼──────────────┤
│ First National Bank │ $1,500,000 │ $0         │ $300,000   │ $1,200,000 │ $1,200,000 │ ✓ RECONCILED │
│ Wells Fargo         │ $800,000   │ $0         │ $150,000   │ $650,000   │ $650,000   │ ✓ RECONCILED │
│ SBA Loan            │ $450,000   │ $0         │ $50,000    │ $400,000   │ $400,000   │ ✓ RECONCILED │
│ Equipment Finance   │ $1,250,000 │ $0         │ $250,000   │ $1,000,000 │ $1,000,000 │ ✓ RECONCILED │
├─────────────────────┼────────────┼────────────┼────────────┼────────────┼────────────┼──────────────┤
│ TOTAL               │ $4,000,000 │ $0         │ $750,000   │ $3,250,000 │ $3,250,000 │              │
└─────────────────────┴────────────┴────────────┴────────────┴────────────┴────────────┴──────────────┘
```

**TEST 3: INTEREST EXPENSE RECALCULATION**
```
┌──────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 3: INTEREST EXPENSE RECALCULATION                                                               │
├─────────────────────┬────────────┬──────────┬────────────┬────────────┬────────────┬────────────────┤
│ Lender              │ Avg Balance│ Rate     │ Expected   │ Actual     │ Difference │ Status         │
├─────────────────────┼────────────┼──────────┼────────────┼────────────┼────────────┼────────────────┤
│ First National Bank │ $1,350,000 │ 6.50%    │ $87,750    │ $86,500    │ ($1,250)   │ ✓ REASONABLE   │
│ Wells Fargo         │ $725,000   │ 7.00%    │ $50,750    │ $51,200    │ $450       │ ✓ REASONABLE   │
│ SBA Loan            │ $425,000   │ 5.00%    │ $21,250    │ $21,250    │ $0         │ ✓ REASONABLE   │
│ Equipment Finance   │ $1,125,000 │ 8.50%    │ $95,625    │ $96,000    │ $375       │ ✓ REASONABLE   │
├─────────────────────┼────────────┼──────────┼────────────┼────────────┼────────────┼────────────────┤
│ TOTAL               │            │          │ $255,375   │ $254,950   │ ($425)     │ 0.2% Variance  │
└─────────────────────┴────────────┴──────────┴────────────┴────────────┴────────────┴────────────────┘
```

**TEST 4: CURRENT PORTION CLASSIFICATION**
```
┌──────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 4: CURRENT PORTION CLASSIFICATION                                                               │
├─────────────────────┬────────────┬────────────┬──────────────┬────────────┬────────────┬─────────────┤
│ Lender              │ CY Balance │ Maturity   │ Mths to Mat  │ Est Current│ Per Books  │ Status      │
├─────────────────────┼────────────┼────────────┼──────────────┼────────────┼────────────┼─────────────┤
│ First National Bank │ $1,200,000 │ 01/15/2029 │ 49           │ $300,000   │ $300,000   │ ✓ REASONABLE│
│ Wells Fargo         │ $650,000   │ 03/01/2027 │ 26           │ $150,000   │ $150,000   │ ✓ REASONABLE│
│ SBA Loan            │ $400,000   │ 06/15/2032 │ 90           │ $0         │ $0         │ ✓ REASONABLE│
│ Equipment Finance   │ $1,000,000 │ 12/31/2028 │ 48           │ $0         │ $0         │ ⚠ REVIEW    │
├─────────────────────┴────────────┴────────────┴──────────────┴────────────┴────────────┴─────────────┤
│ Total Current Portion: $450,000  |  Total Long-Term: $2,800,000                                      │
└──────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                   │
├─────────────────────────────────────────────────────────────────┤
│ Total Debt:                              $3,250,000             │
│   Current Portion:                       $450,000               │
│   Long-Term:                             $2,800,000             │
│ Interest Expense:                        $254,950               │
│                                                                 │
│ Procedures Performed:                                           │
│   ✓ Schedule to GL reconciliation                               │
│   ✓ Debt rollforward testing                                    │
│   ✓ Interest expense recalculation                              │
│   ✓ Current portion classification                              │
│   ☐ Lender confirmations (manual)                               │
│   ☐ Covenant compliance (manual)                                │
│                                                                 │
│ CONCLUSION:                                                     │
│ [Document conclusion]                                           │
└─────────────────────────────────────────────────────────────────┘
```

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ PP&E](./ppe.md) | [➡️ Equity](./equity.md)
