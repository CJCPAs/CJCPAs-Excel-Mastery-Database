# Prepaid Expenses & Other Assets Audit VBA

> **Prepaids Testing** - Complete VBA for auditing prepaid expenses per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 1300-1399 (typically) |
| **Assertions** | Existence, Valuation, Cutoff, Classification |
| **Risk Level** | Low (routine, predictable) |
| **Common Prepaids** | Insurance, rent, subscriptions, maintenance contracts |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for prepaid accounts

### Input Sheet 2: `Prepaids_Schedule`
Schedule of prepaid expenses

| Column | Header | Example |
|--------|--------|---------|
| A | `Account` | 1310 |
| B | `Description` | Prepaid Insurance |
| C | `Vendor` | ABC Insurance Co |
| D | `Policy_Period_Start` | 07/01/2024 |
| E | `Policy_Period_End` | 06/30/2025 |
| F | `Total_Premium` | 24000 |
| G | `PY_Balance` | 12000 |
| H | `CY_Additions` | 24000 |
| I | `CY_Amortization` | 12000 |
| J | `CY_Balance` | 12000 |

---

## Audit Procedures

```vba
Sub AuditPrepaids()
    '================================================
    ' PREPAID EXPENSES - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with prepaid transactions
    '   - Sheet "Prepaids_Schedule" with prepaid detail
    '
    ' OUTPUTS:
    '   - Creates "Prepaids_Audit" worksheet
    '   - Reconciles schedule to GL
    '   - Recalculates amortization
    '   - Tests for proper cutoff
    '
    ' ASSERTIONS TESTED:
    '   - Existence (assets have future benefit)
    '   - Valuation (amortization accurate)
    '   - Cutoff (proper period allocation)
    '================================================

    Dim wsGL As Worksheet
    Dim wsPrepaids As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const MATERIALITY As Double = 25000

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsPrepaids = ThisWorkbook.Sheets("Prepaids_Schedule")
    On Error GoTo 0

    If wsGL Is Nothing Or wsPrepaids Is Nothing Then
        MsgBox "GL_Detail and Prepaids_Schedule sheets required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Prepaids_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Prepaids_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "PREPAID EXPENSES - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: SCHEDULE TO GL RECONCILIATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: SCHEDULE TO GL RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        ' Calculate GL balance
        Dim glPrepaids As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            acctNum = CStr(wsGL.Cells(i, 3).Value)

            ' Prepaid accounts (13xx) - debit balance
            If Left(acctNum, 2) = "13" Then
                glPrepaids = glPrepaids + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
            End If
        Next i

        ' Calculate schedule balance
        Dim slPrepaids As Double

        lastRow = wsPrepaids.Cells(wsPrepaids.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            slPrepaids = slPrepaids + wsPrepaids.Cells(i, 10).Value
        Next i

        .Cells(auditRow, 1).Value = "GL Balance"
        .Cells(auditRow, 2).Value = glPrepaids
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Schedule Balance"
        .Cells(auditRow, 2).Value = slPrepaids
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Difference"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glPrepaids - slPrepaids

        If Abs(glPrepaids - slPrepaids) < 1 Then
            .Cells(auditRow, 3).Value = "RECONCILED"
            .Cells(auditRow, 3).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 3).Value = "DIFFERENCE"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 199, 206)
        End If

        .Range(.Cells(auditRow - 2, 2), .Cells(auditRow, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 3

        ' ========================================
        ' TEST 2: AMORTIZATION RECALCULATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 2: AMORTIZATION RECALCULATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 10)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Description"
        .Cells(auditRow, 2).Value = "Start"
        .Cells(auditRow, 3).Value = "End"
        .Cells(auditRow, 4).Value = "Total"
        .Cells(auditRow, 5).Value = "Term (Mo)"
        .Cells(auditRow, 6).Value = "Mo in CY"
        .Cells(auditRow, 7).Value = "Calc Amort"
        .Cells(auditRow, 8).Value = "Actual"
        .Cells(auditRow, 9).Value = "Diff"
        .Cells(auditRow, 10).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 10)).Font.Bold = True
        auditRow = auditRow + 1

        Dim amortStart As Long
        amortStart = auditRow

        Dim yearEnd As Date
        yearEnd = DateSerial(Year(Date), 12, 31)

        Dim yearStart As Date
        yearStart = DateSerial(Year(Date), 1, 1)

        lastRow = wsPrepaids.Cells(wsPrepaids.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim polStart As Date
            Dim polEnd As Date
            Dim totalPrem As Double
            Dim actualAmort As Double
            Dim termMonths As Long
            Dim monthsInCY As Long
            Dim calcAmort As Double
            Dim amortDiff As Double

            On Error Resume Next
            polStart = wsPrepaids.Cells(i, 4).Value
            polEnd = wsPrepaids.Cells(i, 5).Value
            On Error GoTo 0

            totalPrem = wsPrepaids.Cells(i, 6).Value
            actualAmort = wsPrepaids.Cells(i, 9).Value

            If IsDate(polStart) And IsDate(polEnd) Then
                ' Calculate term in months
                termMonths = DateDiff("m", polStart, polEnd) + 1

                ' Calculate months in current year
                Dim effStart As Date
                Dim effEnd As Date

                If polStart < yearStart Then
                    effStart = yearStart
                Else
                    effStart = polStart
                End If

                If polEnd > yearEnd Then
                    effEnd = yearEnd
                Else
                    effEnd = polEnd
                End If

                If effEnd >= effStart Then
                    monthsInCY = DateDiff("m", effStart, effEnd) + 1
                Else
                    monthsInCY = 0
                End If

                ' Calculate expected amortization
                If termMonths > 0 Then
                    calcAmort = (totalPrem / termMonths) * monthsInCY
                Else
                    calcAmort = 0
                End If

                amortDiff = actualAmort - calcAmort

                .Cells(auditRow, 1).Value = wsPrepaids.Cells(i, 2).Value
                .Cells(auditRow, 2).Value = polStart
                .Cells(auditRow, 3).Value = polEnd
                .Cells(auditRow, 4).Value = totalPrem
                .Cells(auditRow, 5).Value = termMonths
                .Cells(auditRow, 6).Value = monthsInCY
                .Cells(auditRow, 7).Value = calcAmort
                .Cells(auditRow, 8).Value = actualAmort
                .Cells(auditRow, 9).Value = amortDiff

                ' Allow 5% variance for timing
                If Abs(amortDiff) < totalPrem * 0.05 Or Abs(amortDiff) < 500 Then
                    .Cells(auditRow, 10).Value = "REASONABLE"
                    .Cells(auditRow, 10).Interior.Color = RGB(198, 239, 206)
                Else
                    .Cells(auditRow, 10).Value = "INVESTIGATE"
                    .Cells(auditRow, 10).Interior.Color = RGB(255, 199, 206)
                End If

                auditRow = auditRow + 1
            End If
        Next i

        .Range(.Cells(amortStart, 2), .Cells(auditRow - 1, 3)).NumberFormat = "mm/dd/yyyy"
        .Range(.Cells(amortStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(amortStart, 7), .Cells(auditRow - 1, 9)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 3: ROLLFORWARD
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 3: PREPAID ROLLFORWARD"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Description"
        .Cells(auditRow, 2).Value = "PY Balance"
        .Cells(auditRow, 3).Value = "Additions"
        .Cells(auditRow, 4).Value = "Amortization"
        .Cells(auditRow, 5).Value = "Expected"
        .Cells(auditRow, 6).Value = "CY Balance"
        .Cells(auditRow, 7).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
        auditRow = auditRow + 1

        Dim rollStart As Long
        rollStart = auditRow

        Dim totalPYBal As Double, totalAdd As Double, totalAmort As Double, totalCYBal As Double

        lastRow = wsPrepaids.Cells(wsPrepaids.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim pyBal As Double
            Dim cyAdd As Double
            Dim cyAmort As Double
            Dim cyBal As Double
            Dim expectedBal As Double

            pyBal = wsPrepaids.Cells(i, 7).Value
            cyAdd = wsPrepaids.Cells(i, 8).Value
            cyAmort = wsPrepaids.Cells(i, 9).Value
            cyBal = wsPrepaids.Cells(i, 10).Value
            expectedBal = pyBal + cyAdd - cyAmort

            .Cells(auditRow, 1).Value = wsPrepaids.Cells(i, 2).Value
            .Cells(auditRow, 2).Value = pyBal
            .Cells(auditRow, 3).Value = cyAdd
            .Cells(auditRow, 4).Value = cyAmort
            .Cells(auditRow, 5).Value = expectedBal
            .Cells(auditRow, 6).Value = cyBal

            If Abs(expectedBal - cyBal) < 1 Then
                .Cells(auditRow, 7).Value = "TIES"
                .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
            Else
                .Cells(auditRow, 7).Value = "DIFFERENCE"
                .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
            End If

            totalPYBal = totalPYBal + pyBal
            totalAdd = totalAdd + cyAdd
            totalAmort = totalAmort + cyAmort
            totalCYBal = totalCYBal + cyBal

            auditRow = auditRow + 1
        Next i

        ' Totals
        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = totalPYBal
        .Cells(auditRow, 3).Value = totalAdd
        .Cells(auditRow, 4).Value = totalAmort
        .Cells(auditRow, 5).Value = totalPYBal + totalAdd - totalAmort
        .Cells(auditRow, 6).Value = totalCYBal
        .Range(.Cells(auditRow, 2), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        .Range(.Cells(rollStart, 2), .Cells(auditRow - 1, 6)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 4: EXISTENCE TESTING
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: EXISTENCE - POLICY/CONTRACT VERIFICATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Description"
        .Cells(auditRow, 2).Value = "Vendor"
        .Cells(auditRow, 3).Value = "Amount"
        .Cells(auditRow, 4).Value = "Doc Obtained"
        .Cells(auditRow, 5).Value = "Terms Verified"
        .Cells(auditRow, 6).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        lastRow = wsPrepaids.Cells(wsPrepaids.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            .Cells(auditRow, 1).Value = wsPrepaids.Cells(i, 2).Value
            .Cells(auditRow, 2).Value = wsPrepaids.Cells(i, 3).Value
            .Cells(auditRow, 3).Value = wsPrepaids.Cells(i, 10).Value
            .Cells(auditRow, 3).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Cells(auditRow, 4).Value = "[Y/N]"
            .Cells(auditRow, 4).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 5).Value = "[Y/N]"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 6).Value = "[Complete]"
            auditRow = auditRow + 1
        Next i

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

        .Cells(auditRow, 1).Value = "Total Prepaid Expenses:"
        .Cells(auditRow, 2).Value = slPrepaids
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Schedule to GL reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Amortization recalculation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Rollforward testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Policy/contract verification (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Cutoff testing (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 25
        .Columns("B:J").ColumnWidth = 13

    End With

    Application.ScreenUpdating = True

    MsgBox "Prepaids Audit Complete!" & vbCrLf & _
           "Total Prepaids: " & Format(slPrepaids, "$#,##0"), vbInformation

End Sub
```

---

## Common Prepaid Types

| Prepaid Type | Amortization Method | Documentation |
|-------------|---------------------|---------------|
| **Insurance** | Straight-line over policy period | Policy declarations |
| **Rent** | Straight-line over lease term | Lease agreement |
| **Software Subscriptions** | Straight-line over subscription | Invoice/contract |
| **Maintenance Contracts** | Straight-line over contract | Service agreement |
| **Advertising** | When service received | Contract, invoices |
| **Deposits** | N/A until refund/applied | Deposit agreement |
| **Dues & Memberships** | Over membership period | Invoice, certificate |

---

## Search for Unrecorded Prepaids

```vba
Sub SearchUnrecordedPrepaids()
    '================================================
    ' SEARCH FOR UNRECORDED PREPAIDS
    '
    ' Scans expense accounts for large payments that
    ' may need to be capitalized as prepaids
    '================================================

    Dim wsGL As Worksheet
    Dim wsResults As Worksheet
    Dim lastRow As Long, i As Long
    Dim resultRow As Long

    Const MIN_AMOUNT As Double = 10000
    Const PREPAID_KEYWORDS As String = "insurance,rent,subscription,maintenance,deposit,annual,membership,retainer"

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Sheets("Prepaid_Search").Delete
    On Error GoTo 0

    Set wsResults = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResults.Name = "Prepaid_Search"

    With wsResults
        .Range("A1").Value = "SEARCH FOR UNRECORDED PREPAIDS"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Minimum Amount: " & Format(MIN_AMOUNT, "$#,##0")

        resultRow = 4

        .Cells(resultRow, 1).Value = "Date"
        .Cells(resultRow, 2).Value = "Account"
        .Cells(resultRow, 3).Value = "Description"
        .Cells(resultRow, 4).Value = "Amount"
        .Cells(resultRow, 5).Value = "Keyword"
        .Cells(resultRow, 6).Value = "Action"
        .Range(.Cells(resultRow, 1), .Cells(resultRow, 6)).Font.Bold = True
        .Range(.Cells(resultRow, 1), .Cells(resultRow, 6)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(resultRow, 1), .Cells(resultRow, 6)).Font.Color = RGB(255, 255, 255)
        resultRow = resultRow + 1

        Dim keywords() As String
        keywords = Split(PREPAID_KEYWORDS, ",")

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            Dim desc As String
            Dim amt As Double
            Dim isExpense As Boolean

            acctNum = CStr(wsGL.Cells(i, 3).Value)
            desc = LCase(wsGL.Cells(i, 5).Value)
            amt = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value  ' Debit balance

            ' Check if expense account (6xxx, 7xxx)
            isExpense = (Left(acctNum, 1) = "6" Or Left(acctNum, 1) = "7")

            If isExpense And amt >= MIN_AMOUNT Then
                ' Check for prepaid keywords
                Dim kw As Variant
                For Each kw In keywords
                    If InStr(1, desc, kw, vbTextCompare) > 0 Then
                        .Cells(resultRow, 1).Value = wsGL.Cells(i, 1).Value
                        .Cells(resultRow, 2).Value = wsGL.Cells(i, 4).Value
                        .Cells(resultRow, 3).Value = wsGL.Cells(i, 5).Value
                        .Cells(resultRow, 4).Value = amt
                        .Cells(resultRow, 5).Value = UCase(kw)
                        .Cells(resultRow, 6).Value = "[Review]"
                        .Cells(resultRow, 5).Interior.Color = RGB(255, 235, 156)
                        resultRow = resultRow + 1
                        Exit For
                    End If
                Next kw
            End If
        Next i

        .Range(.Cells(5, 4), .Cells(resultRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 40
        .Columns("D:F").ColumnWidth = 14

    End With

    MsgBox "Prepaid Search Complete!" & vbCrLf & _
           (resultRow - 5) & " potential items identified.", vbInformation

End Sub
```

---

## Output Examples

### Prepaids_Audit Worksheet

The `AuditPrepaids` procedure generates a comprehensive worksheet:

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ PREPAID EXPENSES - AUDIT WORKPAPER                                                  │
│ Period: December 31, 2024                                                           │
│ Prepared: AUDITOR on 12/15/2024 3:00:00 PM                                         │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ TEST 1: SCHEDULE TO GL RECONCILIATION                                               │
├───────────────────────────────────┬─────────────────────────────────────────────────┤
│ GL Balance                        │ $186,500                                        │
│ Schedule Balance                  │ $186,500                                        │
│ Difference                        │ $0.00                       ✓ RECONCILED        │
└───────────────────────────────────┴─────────────────────────────────────────────────┘
```

**TEST 2: AMORTIZATION RECALCULATION**
```
┌───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 2: AMORTIZATION RECALCULATION                                                                                        │
├────────────────────────┬────────────┬────────────┬───────────┬─────────┬─────────┬───────────┬──────────┬───────┬────────┤
│ Description            │ Start      │ End        │ Total     │ Term(Mo)│ Mo in CY│ Calc Amort│ Actual   │ Diff  │ Status │
├────────────────────────┼────────────┼────────────┼───────────┼─────────┼─────────┼───────────┼──────────┼───────┼────────┤
│ Prepaid Insurance      │ 07/01/2024 │ 06/30/2025 │ $36,000   │ 12      │ 6       │ $18,000   │ $18,000  │ $0    │ ✓      │
│ Prepaid Rent           │ 01/01/2024 │ 12/31/2024 │ $120,000  │ 12      │ 12      │ $120,000  │ $120,000 │ $0    │ ✓      │
│ Software Subscription  │ 04/01/2024 │ 03/31/2025 │ $24,000   │ 12      │ 9       │ $18,000   │ $17,500  │ ($500)│ ✓      │
│ Maintenance Contract   │ 10/01/2024 │ 09/30/2025 │ $18,000   │ 12      │ 3       │ $4,500    │ $4,500   │ $0    │ ✓      │
└────────────────────────┴────────────┴────────────┴───────────┴─────────┴─────────┴───────────┴──────────┴───────┴────────┘
  Status: ✓ REASONABLE = Difference within 5% or $500
```

**TEST 3: PREPAID ROLLFORWARD**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 3: PREPAID ROLLFORWARD                                                                             │
├────────────────────────┬────────────┬────────────┬────────────┬────────────┬────────────┬───────────────┤
│ Description            │ PY Balance │ Additions  │ Amortization│ Expected   │ CY Balance │ Status        │
├────────────────────────┼────────────┼────────────┼────────────┼────────────┼────────────┼───────────────┤
│ Prepaid Insurance      │ $18,000    │ $36,000    │ $18,000    │ $36,000    │ $36,000    │ ✓ TIES        │
│ Prepaid Rent           │ $10,000    │ $120,000   │ $120,000   │ $10,000    │ $10,000    │ ✓ TIES        │
│ Software Subscription  │ $6,000     │ $24,000    │ $17,500    │ $12,500    │ $12,500    │ ✓ TIES        │
│ Maintenance Contract   │ $0         │ $18,000    │ $4,500     │ $13,500    │ $13,500    │ ✓ TIES        │
├────────────────────────┼────────────┼────────────┼────────────┼────────────┼────────────┼───────────────┤
│ TOTAL                  │ $34,000    │ $198,000   │ $160,000   │ $72,000    │ $72,000    │               │
└────────────────────────┴────────────┴────────────┴────────────┴────────────┴────────────┴───────────────┘
```

**TEST 4: EXISTENCE - POLICY/CONTRACT VERIFICATION**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 4: EXISTENCE - POLICY/CONTRACT VERIFICATION                                                    │
├────────────────────────┬─────────────────────┬────────────┬────────────┬──────────────┬─────────────┤
│ Description            │ Vendor              │ Amount     │ Doc Obtained│ Terms Verified│ Status     │
├────────────────────────┼─────────────────────┼────────────┼────────────┼──────────────┼─────────────┤
│ Prepaid Insurance      │ ABC Insurance Co    │ $36,000    │ [Y/N]  ▓▓▓ │ [Y/N]    ▓▓▓ │ [Complete]  │
│ Prepaid Rent           │ XYZ Properties      │ $10,000    │ [Y/N]  ▓▓▓ │ [Y/N]    ▓▓▓ │ [Complete]  │
│ Software Subscription  │ Software Corp       │ $12,500    │ [Y/N]  ▓▓▓ │ [Y/N]    ▓▓▓ │ [Complete]  │
│ Maintenance Contract   │ Service Provider    │ $13,500    │ [Y/N]  ▓▓▓ │ [Y/N]    ▓▓▓ │ [Complete]  │
└────────────────────────┴─────────────────────┴────────────┴────────────┴──────────────┴─────────────┘
  (▓▓▓ = Yellow highlight for input cells)
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                               │
├─────────────────────────────────────────────────────────────────────────────┤
│ Total Prepaid Expenses: $72,000                                             │
├─────────────────────────────────────────────────────────────────────────────┤
│ Procedures Performed:                                                       │
│   ✓ Schedule to GL reconciliation                                           │
│   ✓ Amortization recalculation                                              │
│   ✓ Rollforward testing                                                     │
│   ☐ Policy/contract verification (manual)                                   │
│   ☐ Cutoff testing (manual)                                                 │
├─────────────────────────────────────────────────────────────────────────────┤
│ CONCLUSION: [Document conclusion]                                           │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Prepaid_Search Worksheet

```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ SEARCH FOR UNRECORDED PREPAIDS                                                                              │
│ Minimum Amount: $10,000                                                                                     │
├────────────┬────────────────────────┬──────────────────────────────────────────┬───────────┬────────┬───────┤
│ Date       │ Account                │ Description                              │ Amount    │ Keyword│ Action│
├────────────┼────────────────────────┼──────────────────────────────────────────┼───────────┼────────┼───────┤
│ 11/15/2024 │ Professional Services  │ Annual retainer - legal services         │ $24,000   │ ANNUAL │[Review]│
│ 12/01/2024 │ Insurance Expense      │ Liability insurance premium - 2025       │ $18,500   │INSURANCE│[Review]│
│ 12/15/2024 │ Software Expense       │ Annual subscription renewal              │ $15,000   │SUBSCRIPT│[Review]│
└────────────┴────────────────────────┴──────────────────────────────────────────┴───────────┴────────┴───────┘
  3 potential items identified for review
```

---

## Key Audit Considerations

| Consideration | Audit Response |
|--------------|----------------|
| **Large New Prepaids** | Obtain and inspect supporting documents |
| **Expired Policies** | Ensure fully amortized |
| **Refundable Deposits** | Confirm still refundable |
| **Multi-Year Contracts** | Verify amortization schedule |
| **Prepaid vs. Expense** | Test cutoff for year-end payments |
| **Classification** | Current vs. non-current |

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Inventory](./inventory.md) | [➡️ PP&E](./ppe.md)
