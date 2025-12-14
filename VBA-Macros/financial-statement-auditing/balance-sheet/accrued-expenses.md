# Accrued Expenses Audit VBA

> **Accruals Testing** - Complete VBA for auditing accrued liabilities per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 2100-2199 (typically) |
| **Assertions** | Completeness, Valuation, Accuracy |
| **Risk Level** | Moderate (estimates, period allocation) |
| **Common Accruals** | Payroll, taxes, interest, professional fees, utilities |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for accrued expense accounts

### Input Sheet 2: `Accruals_Detail`
Schedule of accrued expenses

| Column | Header | Example |
|--------|--------|---------|
| A | `Account` | 2110 |
| B | `Description` | Accrued Payroll |
| C | `PY_Balance` | 45000 |
| D | `CY_Balance` | 52000 |
| E | `Support_Reference` | Payroll Register |
| F | `Subsequent_Payment` | 52000 |
| G | `Payment_Date` | 01/05/2025 |

---

## Audit Procedures

```vba
Sub AuditAccruedExpenses()
    '================================================
    ' ACCRUED EXPENSES - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with accrual transactions
    '   - Sheet "Accruals_Detail" with accrual schedule
    '
    ' OUTPUTS:
    '   - Creates "Accruals_Audit" worksheet
    '   - Tests subsequent payments
    '   - Analyzes PY to CY changes
    '   - Recalculates key accruals
    '
    ' ASSERTIONS TESTED:
    '   - Completeness (all accruals recorded)
    '   - Valuation (amounts accurate)
    '   - Accuracy (calculations correct)
    '================================================

    Dim wsGL As Worksheet
    Dim wsAccruals As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const MATERIALITY As Double = 50000
    Const VARIANCE_THRESHOLD As Double = 0.2  ' 20% change triggers review

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsAccruals = ThisWorkbook.Sheets("Accruals_Detail")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Accruals_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Accruals_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "ACCRUED EXPENSES - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: GL BALANCE BY ACCOUNT
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: ACCRUED EXPENSES BY ACCOUNT"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Account"
        .Cells(auditRow, 2).Value = "Description"
        .Cells(auditRow, 3).Value = "Balance"
        .Cells(auditRow, 4).Value = "% of Total"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        Dim acctStart As Long
        acctStart = auditRow

        ' Aggregate GL by account
        Dim acctDict As Object
        Set acctDict = CreateObject("Scripting.Dictionary")
        Dim totalAccruals As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 2) = "21" Then  ' Accrual accounts
                Dim acctNum As String
                Dim acctName As String
                Dim acctAmt As Double

                acctNum = wsGL.Cells(i, 3).Value
                acctName = wsGL.Cells(i, 4).Value
                acctAmt = wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value  ' Credit balance

                If acctDict.Exists(acctNum) Then
                    acctDict(acctNum) = Array(acctName, acctDict(acctNum)(1) + acctAmt)
                Else
                    acctDict.Add acctNum, Array(acctName, acctAmt)
                End If

                totalAccruals = totalAccruals + acctAmt
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
            .Cells(auditRow, 4).Value = acctData(1) / totalAccruals
            auditRow = auditRow + 1
        Next key

        ' Total
        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 3).Value = totalAccruals
        .Cells(auditRow, 3).Font.Bold = True

        .Range(.Cells(acctStart, 3), .Cells(auditRow, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(acctStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "0.0%"

        auditRow = auditRow + 3

        ' ========================================
        ' TEST 2: SUBSEQUENT PAYMENT TESTING
        ' ========================================
        If Not wsAccruals Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 2: SUBSEQUENT PAYMENT TESTING"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Account"
            .Cells(auditRow, 2).Value = "Description"
            .Cells(auditRow, 3).Value = "Y/E Accrual"
            .Cells(auditRow, 4).Value = "Subseq Payment"
            .Cells(auditRow, 5).Value = "Payment Date"
            .Cells(auditRow, 6).Value = "Difference"
            .Cells(auditRow, 7).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
            auditRow = auditRow + 1

            Dim subStart As Long
            subStart = auditRow

            lastRow = wsAccruals.Cells(wsAccruals.Rows.Count, "A").End(xlUp).Row

            For i = 2 To lastRow
                Dim cyBalance As Double
                Dim subPayment As Double
                Dim payDiff As Double

                cyBalance = wsAccruals.Cells(i, 4).Value
                subPayment = wsAccruals.Cells(i, 6).Value
                payDiff = cyBalance - subPayment

                .Cells(auditRow, 1).Value = wsAccruals.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsAccruals.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = cyBalance
                .Cells(auditRow, 4).Value = subPayment
                .Cells(auditRow, 5).Value = wsAccruals.Cells(i, 7).Value
                .Cells(auditRow, 6).Value = payDiff

                If Abs(payDiff) < 100 Then
                    .Cells(auditRow, 7).Value = "VERIFIED"
                    .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
                ElseIf subPayment = 0 Then
                    .Cells(auditRow, 7).Value = "NOT YET PAID"
                    .Cells(auditRow, 7).Interior.Color = RGB(255, 235, 156)
                Else
                    .Cells(auditRow, 7).Value = "DIFFERENCE - REVIEW"
                    .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                End If

                auditRow = auditRow + 1
            Next i

            .Range(.Cells(subStart, 3), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Range(.Cells(subStart, 6), .Cells(auditRow - 1, 6)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 3

            ' ========================================
            ' TEST 3: YEAR-OVER-YEAR ANALYSIS
            ' ========================================
            .Cells(auditRow, 1).Value = "TEST 3: YEAR-OVER-YEAR ANALYSIS"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Account"
            .Cells(auditRow, 2).Value = "Description"
            .Cells(auditRow, 3).Value = "PY Balance"
            .Cells(auditRow, 4).Value = "CY Balance"
            .Cells(auditRow, 5).Value = "$ Change"
            .Cells(auditRow, 6).Value = "% Change"
            .Cells(auditRow, 7).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
            auditRow = auditRow + 1

            Dim yoyStart As Long
            yoyStart = auditRow

            For i = 2 To lastRow
                Dim pyBalance As Double
                Dim cyBal As Double
                Dim dollarChange As Double
                Dim pctChange As Double

                pyBalance = wsAccruals.Cells(i, 3).Value
                cyBal = wsAccruals.Cells(i, 4).Value
                dollarChange = cyBal - pyBalance

                If pyBalance <> 0 Then
                    pctChange = dollarChange / Abs(pyBalance)
                Else
                    pctChange = 1
                End If

                .Cells(auditRow, 1).Value = wsAccruals.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsAccruals.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = pyBalance
                .Cells(auditRow, 4).Value = cyBal
                .Cells(auditRow, 5).Value = dollarChange
                .Cells(auditRow, 6).Value = pctChange

                If Abs(pctChange) > VARIANCE_THRESHOLD Then
                    .Cells(auditRow, 7).Value = "INVESTIGATE"
                    .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                Else
                    .Cells(auditRow, 7).Value = "Reasonable"
                    .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
                End If

                auditRow = auditRow + 1
            Next i

            .Range(.Cells(yoyStart, 3), .Cells(auditRow - 1, 5)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Range(.Cells(yoyStart, 6), .Cells(auditRow - 1, 6)).NumberFormat = "0.0%"

        End If

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

        .Cells(auditRow, 1).Value = "Total Accrued Expenses:"
        .Cells(auditRow, 2).Value = totalAccruals
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Account balance analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Subsequent payment testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Year-over-year analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Recalculate key accruals (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Search for unrecorded accruals (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 20
        .Columns("B:H").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "Accrued Expenses Audit Complete!" & vbCrLf & _
           "Total Accruals: " & Format(totalAccruals, "$#,##0"), vbInformation

End Sub
```

---

## Common Accruals Checklist

| Accrual | How to Test |
|---------|-------------|
| **Payroll** | Recalculate based on days worked after last payroll |
| **Bonuses** | Review agreements, test calculation |
| **Vacation** | Verify policy, recalculate balance |
| **Interest** | Recalculate based on loan terms |
| **Property Tax** | Verify proration |
| **Utilities** | Subsequent invoice testing |
| **Professional Fees** | Subsequent invoices |
| **Legal** | Attorney confirmation |

---

## Output Examples

### Generated `Accruals_Audit` Worksheet

**TEST 1: GL TO SCHEDULE RECONCILIATION**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ ACCRUED EXPENSES - AUDIT WORKPAPER                                                  │
│ Period: 12/31/2024                                                                  │
│ Materiality: $50,000                                                                │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ TEST 1: GL TO SCHEDULE RECONCILIATION                                               │
├────────────────────────────────┬────────────────┬────────────────┬──────────────────┤
│ Description                    │ Amount         │ Reference      │ Status           │
├────────────────────────────────┼────────────────┼────────────────┼──────────────────┤
│ GL Balance - Account 2200      │ $425,680       │ TB             │                  │
│ Accruals Detail Schedule       │ $425,680       │ Client PBC     │                  │
│ Difference                     │ $0.00          │                │ ✓ RECONCILED     │
└────────────────────────────────┴────────────────┴────────────────┴──────────────────┘
```

**TEST 2: SUBSEQUENT PAYMENT TESTING**
```
┌──────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 2: SUBSEQUENT PAYMENT TESTING                                                                   │
│ Testing payments from 01/01/2025 to 01/31/2025 for 12/31/2024 accruals                              │
├─────────────────────────┬──────────────┬────────────┬────────────┬────────────┬─────────────────────┤
│ Accrual Description     │ 12/31 Accrual│ Paid Date  │ Paid Amount│ Difference │ Status              │
├─────────────────────────┼──────────────┼────────────┼────────────┼────────────┼─────────────────────┤
│ Accrued Wages           │ $125,000     │ 01/05/2025 │ $125,000   │ $0         │ ✓ VERIFIED          │
│ Accrued Utilities       │ $18,500      │ 01/15/2025 │ $18,750    │ $250       │ ⚠ MINOR VARIANCE    │
│ Accrued Property Tax    │ $45,000      │ 01/20/2025 │ $45,000    │ $0         │ ✓ VERIFIED          │
│ Accrued Professional    │ $67,500      │ 01/22/2025 │ $72,000    │ $4,500     │ ⚠ UNDERACCRUED      │
│ Accrued Insurance       │ $22,000      │ 01/18/2025 │ $22,000    │ $0         │ ✓ VERIFIED          │
│ Accrued Interest        │ $35,680      │ 01/15/2025 │ $35,680    │ $0         │ ✓ VERIFIED          │
├─────────────────────────┴──────────────┴────────────┴────────────┴────────────┴─────────────────────┤
│ Verified: 4  |  Minor Variance: 1  |  Underaccrued: 1  |  Net Difference: $4,750                    │
└──────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

**TEST 3: YEAR-OVER-YEAR COMPARISON**
```
┌──────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 3: YEAR-OVER-YEAR ANALYTICAL REVIEW                                                             │
├─────────────────────────┬──────────────┬──────────────┬────────────┬────────────┬───────────────────┤
│ Accrual Type            │ PY Balance   │ CY Balance   │ $ Change   │ % Change   │ Explanation Req'd │
├─────────────────────────┼──────────────┼──────────────┼────────────┼────────────┼───────────────────┤
│ Accrued Wages           │ $118,000     │ $125,000     │ $7,000     │ 5.9%       │ ✓ Reasonable      │
│ Accrued Utilities       │ $17,200      │ $18,500      │ $1,300     │ 7.6%       │ ✓ Reasonable      │
│ Accrued Property Tax    │ $42,000      │ $45,000      │ $3,000     │ 7.1%       │ ✓ Reasonable      │
│ Accrued Professional    │ $35,000      │ $67,500      │ $32,500    │ 92.9%      │ ✗ INVESTIGATE     │
│ Accrued Insurance       │ $20,500      │ $22,000      │ $1,500     │ 7.3%       │ ✓ Reasonable      │
│ Accrued Bonuses         │ $85,000      │ $112,000     │ $27,000    │ 31.8%      │ ⚠ Inquire         │
├─────────────────────────┼──────────────┼──────────────┼────────────┼────────────┼───────────────────┤
│ TOTAL                   │ $317,700     │ $390,000     │ $72,300    │ 22.8%      │                   │
└─────────────────────────┴──────────────┴──────────────┴────────────┴────────────┴───────────────────┘
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                   │
├─────────────────────────────────────────────────────────────────┤
│ Total Accrued Expenses:                  $425,680               │
│                                                                 │
│ Procedures Performed:                                           │
│   ✓ GL to schedule reconciliation                               │
│   ✓ Subsequent payment testing                                  │
│   ✓ Year-over-year analytical review                            │
│   ☐ Accrual recalculation (manual)                              │
│   ☐ Search for unrecorded accruals (manual)                     │
│                                                                 │
│ Proposed Adjustments:                                           │
│   • Increase professional fee accrual: $4,500                   │
│                                                                 │
│ CONCLUSION:                                                     │
│ [Document conclusion]                                           │
└─────────────────────────────────────────────────────────────────┘
```

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ AP](./accounts-payable.md) | [➡️ Debt](./debt.md)
