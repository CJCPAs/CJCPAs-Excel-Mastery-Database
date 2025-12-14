# Income Tax Expense Audit VBA

> **Tax Provision Testing** - Complete VBA for auditing income tax expense/benefit per GAAS/GAAP (ASC 740)

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 8xxx (Income Tax Expense), 2xxx (Tax Liability), 1xxx (Deferred Tax Asset) |
| **Assertions** | Valuation, Accuracy, Completeness, Presentation |
| **Risk Level** | High (complex estimates, GAAP vs. tax differences) |
| **Key Standards** | ASC 740 (Income Taxes) |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for tax accounts

### Input Sheet 2: `Tax_Provision`
Income tax provision calculation

| Column | Header | Example |
|--------|--------|---------|
| A | `Component` | Current Federal |
| B | `PY_Balance` | 450000 |
| C | `CY_Balance` | 525000 |
| D | `Change` | 75000 |

### Input Sheet 3: `Book_Tax_Differences`
Schedule of book-tax differences

| Column | Header | Example |
|--------|--------|---------|
| A | `Description` | Depreciation |
| B | `Type` | Temporary |
| C | `Book_Basis` | 2500000 |
| D | `Tax_Basis` | 1800000 |
| E | `Difference` | 700000 |
| F | `DTA_DTL` | DTL |
| G | `Deferred_Tax` | 147000 |

---

## Audit Procedures

```vba
Sub AuditIncomeTax()
    '================================================
    ' INCOME TAX EXPENSE - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with tax transactions
    '   - Sheet "Tax_Provision" with provision detail
    '   - Sheet "Book_Tax_Differences" with temp/perm diffs
    '
    ' OUTPUTS:
    '   - Creates "Tax_Audit" worksheet
    '   - Analyzes effective tax rate
    '   - Tests current/deferred components
    '   - Reviews book-tax differences
    '
    ' ASSERTIONS TESTED:
    '   - Valuation (provision accurate)
    '   - Accuracy (calculations correct)
    '   - Completeness (all items considered)
    '================================================

    Dim wsGL As Worksheet
    Dim wsProv As Worksheet
    Dim wsDiff As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const FED_RATE As Double = 0.21  ' 21% federal rate
    Const STATE_RATE As Double = 0.05  ' Assumed state rate

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsProv = ThisWorkbook.Sheets("Tax_Provision")
    Set wsDiff = ThisWorkbook.Sheets("Book_Tax_Differences")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Tax_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Tax_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "INCOME TAX EXPENSE - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now
        .Range("A4").Value = "Reference: ASC 740 - Income Taxes"

        auditRow = 6

        ' ========================================
        ' TEST 1: TAX PROVISION SUMMARY
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: TAX PROVISION SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        ' Calculate GL tax expense
        Dim glCurrentTax As Double
        Dim glDeferredTax As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            Dim acctName As String
            acctNum = CStr(wsGL.Cells(i, 3).Value)
            acctName = LCase(wsGL.Cells(i, 4).Value)

            ' Tax expense accounts (8xxx typically)
            If Left(acctNum, 1) = "8" Then
                Dim taxAmt As Double
                taxAmt = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value

                If InStr(acctName, "current") > 0 Then
                    glCurrentTax = glCurrentTax + taxAmt
                ElseIf InStr(acctName, "deferred") > 0 Then
                    glDeferredTax = glDeferredTax + taxAmt
                Else
                    glCurrentTax = glCurrentTax + taxAmt  ' Default to current
                End If
            End If
        Next i

        .Cells(auditRow, 1).Value = "Component"
        .Cells(auditRow, 2).Value = "Per GL"
        .Cells(auditRow, 3).Value = "Per Schedule"
        .Cells(auditRow, 4).Value = "Difference"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim provStart As Long
        provStart = auditRow

        ' Get provision detail if available
        Dim provCurrent As Double, provDeferred As Double

        If Not wsProv Is Nothing Then
            lastRow = wsProv.Cells(wsProv.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim comp As String
                comp = LCase(wsProv.Cells(i, 1).Value)
                If InStr(comp, "current") > 0 Then
                    provCurrent = provCurrent + wsProv.Cells(i, 3).Value
                ElseIf InStr(comp, "deferred") > 0 Then
                    provDeferred = provDeferred + wsProv.Cells(i, 3).Value
                End If
            Next i
        End If

        ' Current tax
        .Cells(auditRow, 1).Value = "Current Tax Expense"
        .Cells(auditRow, 2).Value = glCurrentTax
        .Cells(auditRow, 3).Value = provCurrent
        .Cells(auditRow, 4).Value = glCurrentTax - provCurrent

        If Abs(glCurrentTax - provCurrent) < 100 Then
            .Cells(auditRow, 5).Value = "RECONCILED"
            .Cells(auditRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 5).Value = "DIFFERENCE"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Deferred tax
        .Cells(auditRow, 1).Value = "Deferred Tax Expense/(Benefit)"
        .Cells(auditRow, 2).Value = glDeferredTax
        .Cells(auditRow, 3).Value = provDeferred
        .Cells(auditRow, 4).Value = glDeferredTax - provDeferred

        If Abs(glDeferredTax - provDeferred) < 100 Then
            .Cells(auditRow, 5).Value = "RECONCILED"
            .Cells(auditRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 5).Value = "DIFFERENCE"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Total
        .Cells(auditRow, 1).Value = "TOTAL TAX EXPENSE"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glCurrentTax + glDeferredTax
        .Cells(auditRow, 2).Font.Bold = True
        .Cells(auditRow, 3).Value = provCurrent + provDeferred
        .Cells(auditRow, 3).Font.Bold = True
        .Cells(auditRow, 4).Value = (glCurrentTax + glDeferredTax) - (provCurrent + provDeferred)
        auditRow = auditRow + 1

        .Range(.Cells(provStart, 2), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 2: EFFECTIVE TAX RATE ANALYSIS
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 2: EFFECTIVE TAX RATE ANALYSIS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Pre-tax Income (Book)"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Total Tax Expense"
        .Cells(auditRow, 2).Value = glCurrentTax + glDeferredTax
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Effective Tax Rate"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Calculate]"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Statutory Rate (Federal)"
        .Cells(auditRow, 2).Value = FED_RATE
        .Cells(auditRow, 2).NumberFormat = "0.0%"
        auditRow = auditRow + 2

        ' Rate reconciliation
        .Cells(auditRow, 1).Value = "RATE RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1

        Dim rateItems As Variant
        rateItems = Array( _
            Array("Federal statutory rate", FED_RATE), _
            Array("State taxes, net of federal", "[Input]"), _
            Array("Permanent differences", "[Input]"), _
            Array("Tax credits", "[Input]"), _
            Array("Change in valuation allowance", "[Input]"), _
            Array("Prior year adjustments", "[Input]"), _
            Array("Other", "[Input]"))

        .Cells(auditRow, 1).Value = "Item"
        .Cells(auditRow, 2).Value = "Rate Impact"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 2)).Font.Bold = True
        auditRow = auditRow + 1

        Dim rateStart As Long
        rateStart = auditRow

        Dim r As Variant
        For Each r In rateItems
            .Cells(auditRow, 1).Value = r(0)
            If r(0) = "Federal statutory rate" Then
                .Cells(auditRow, 2).Value = r(1)
            Else
                .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            End If
            auditRow = auditRow + 1
        Next r

        .Cells(auditRow, 1).Value = "Effective Rate"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Sum]"
        auditRow = auditRow + 1

        .Range(.Cells(rateStart, 2), .Cells(auditRow - 1, 2)).NumberFormat = "0.0%"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 3: DEFERRED TAX ANALYSIS
        ' ========================================
        If Not wsDiff Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 3: DEFERRED TAX ANALYSIS"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Description"
            .Cells(auditRow, 2).Value = "Type"
            .Cells(auditRow, 3).Value = "Book Basis"
            .Cells(auditRow, 4).Value = "Tax Basis"
            .Cells(auditRow, 5).Value = "Difference"
            .Cells(auditRow, 6).Value = "DTA/DTL"
            .Cells(auditRow, 7).Value = "Deferred Tax"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
            auditRow = auditRow + 1

            Dim diffStart As Long
            diffStart = auditRow

            Dim totalDTA As Double, totalDTL As Double

            lastRow = wsDiff.Cells(wsDiff.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                .Cells(auditRow, 1).Value = wsDiff.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsDiff.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = wsDiff.Cells(i, 3).Value
                .Cells(auditRow, 4).Value = wsDiff.Cells(i, 4).Value
                .Cells(auditRow, 5).Value = wsDiff.Cells(i, 5).Value
                .Cells(auditRow, 6).Value = wsDiff.Cells(i, 6).Value
                .Cells(auditRow, 7).Value = wsDiff.Cells(i, 7).Value

                If UCase(wsDiff.Cells(i, 6).Value) = "DTA" Then
                    totalDTA = totalDTA + wsDiff.Cells(i, 7).Value
                Else
                    totalDTL = totalDTL + wsDiff.Cells(i, 7).Value
                End If

                auditRow = auditRow + 1
            Next i

            .Range(.Cells(diffStart, 3), .Cells(auditRow - 1, 5)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Range(.Cells(diffStart, 7), .Cells(auditRow - 1, 7)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Total Deferred Tax Assets:"
            .Cells(auditRow, 7).Value = totalDTA
            .Cells(auditRow, 7).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Total Deferred Tax Liabilities:"
            .Cells(auditRow, 7).Value = totalDTL
            .Cells(auditRow, 7).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Net Deferred Tax:"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 7).Value = totalDTA - totalDTL
            .Cells(auditRow, 7).Font.Bold = True
            .Cells(auditRow, 7).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            auditRow = auditRow + 1

            auditRow = auditRow + 2
        End If

        ' ========================================
        ' TEST 4: DEFERRED TAX ROLLFORWARD
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: DEFERRED TAX ROLLFORWARD"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Beginning Net Deferred Tax"
        .Cells(auditRow, 2).Value = "[Input PY]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Deferred Tax Expense/(Benefit)"
        .Cells(auditRow, 2).Value = glDeferredTax
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "OCI Items"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Other Adjustments"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Calculated Ending"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Sum]"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Per Balance Sheet"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Difference"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Calc]"
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

        .Cells(auditRow, 1).Value = "Current Tax Expense:"
        .Cells(auditRow, 2).Value = glCurrentTax
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Deferred Tax Expense/(Benefit):"
        .Cells(auditRow, 2).Value = glDeferredTax
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "TOTAL TAX EXPENSE:"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glCurrentTax + glDeferredTax
        .Cells(auditRow, 2).Font.Bold = True
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Tax provision summary"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Effective tax rate analysis (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Deferred tax analysis (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Deferred tax rollforward (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Valuation allowance assessment (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Uncertain tax positions (ASC 740-10) (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 35
        .Columns("B:G").ColumnWidth = 14

    End With

    Application.ScreenUpdating = True

    MsgBox "Income Tax Audit Complete!" & vbCrLf & _
           "Total Tax Expense: " & Format(glCurrentTax + glDeferredTax, "$#,##0"), vbInformation

End Sub
```

---

## Output Examples

### Tax_Audit Worksheet

The `AuditIncomeTax` procedure generates a comprehensive worksheet:

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ INCOME TAX EXPENSE - AUDIT WORKPAPER                                                │
│ Period: December 31, 2024                                                           │
│ Prepared: AUDITOR on 12/15/2024 5:30:00 PM                                         │
│ Reference: ASC 740 - Income Taxes                                                   │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

**TEST 1: TAX PROVISION SUMMARY**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 1: TAX PROVISION SUMMARY                                                       │
├────────────────────────────────────────┬────────────┬────────────┬────────┬─────────┤
│ Component                              │ Per GL     │ Per Schedule│ Diff   │ Status  │
├────────────────────────────────────────┼────────────┼────────────┼────────┼─────────┤
│ Current Tax Expense                    │ $525,000   │ $525,000   │ $0     │ ✓       │
│ Deferred Tax Expense/(Benefit)         │ $75,000    │ $75,000    │ $0     │ ✓       │
├────────────────────────────────────────┼────────────┼────────────┼────────┼─────────┤
│ TOTAL TAX EXPENSE                      │ $600,000   │ $600,000   │ $0     │         │
└────────────────────────────────────────┴────────────┴────────────┴────────┴─────────┘
```

**TEST 2: EFFECTIVE TAX RATE ANALYSIS**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ TEST 2: EFFECTIVE TAX RATE ANALYSIS                                         │
├─────────────────────────────────────────┬───────────────────────────────────┤
│ Pre-tax Income (Book)                   │ $2,600,000                        │
│ Total Tax Expense                       │ $600,000                          │
│ Effective Tax Rate                      │ 23.1%                             │
│ Statutory Rate (Federal)                │ 21.0%                             │
└─────────────────────────────────────────┴───────────────────────────────────┘

RATE RECONCILIATION
┌─────────────────────────────────────────────────────────────────────────────┐
│ Item                                    │ Rate Impact                       │
├─────────────────────────────────────────┼───────────────────────────────────┤
│ Federal statutory rate                  │ 21.0%                             │
│ State taxes, net of federal             │ 3.9%                          ▓▓▓ │
│ Permanent differences                   │ (0.8%)                        ▓▓▓ │
│ Tax credits                             │ (1.2%)                        ▓▓▓ │
│ Change in valuation allowance           │ 0.0%                          ▓▓▓ │
│ Prior year adjustments                  │ 0.2%                          ▓▓▓ │
│ Other                                   │ 0.0%                          ▓▓▓ │
├─────────────────────────────────────────┼───────────────────────────────────┤
│ Effective Rate                          │ 23.1%                             │
└─────────────────────────────────────────┴───────────────────────────────────┘
  (▓▓▓ = Yellow highlight for input cells)
```

**TEST 3: DEFERRED TAX ANALYSIS**
```
┌───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 3: DEFERRED TAX ANALYSIS                                                                                             │
├────────────────────────────┬───────────┬────────────────┬────────────────┬────────────────┬─────────┬─────────────────────┤
│ Description                │ Type      │ Book Basis     │ Tax Basis      │ Difference     │ DTA/DTL │ Deferred Tax        │
├────────────────────────────┼───────────┼────────────────┼────────────────┼────────────────┼─────────┼─────────────────────┤
│ Depreciation               │ Temporary │ $2,500,000     │ $1,800,000     │ $700,000       │ DTL     │ $175,000            │
│ Allowance for Doubtful Acct│ Temporary │ $150,000       │ $0             │ $150,000       │ DTA     │ ($37,500)           │
│ Inventory Reserve          │ Temporary │ $85,000        │ $0             │ $85,000        │ DTA     │ ($21,250)           │
│ Warranty Accrual           │ Temporary │ $200,000       │ $0             │ $200,000       │ DTA     │ ($50,000)           │
│ Accrued Compensation       │ Temporary │ $125,000       │ $0             │ $125,000       │ DTA     │ ($31,250)           │
│ Stock Compensation         │ Temporary │ $180,000       │ $60,000        │ $120,000       │ DTA     │ ($30,000)           │
│ NOL Carryforward           │ Temporary │ $0             │ $200,000       │ ($200,000)     │ DTA     │ ($50,000)           │
├────────────────────────────┴───────────┴────────────────┴────────────────┴────────────────┼─────────┼─────────────────────┤
│                                                                                           │         │                     │
│ Total Deferred Tax Assets:                                                                │         │ $220,000            │
│ Total Deferred Tax Liabilities:                                                           │         │ $175,000            │
│ Net Deferred Tax:                                                                         │         │ $45,000 (DTA)       │
└───────────────────────────────────────────────────────────────────────────────────────────┴─────────┴─────────────────────┘
```

**TEST 4: DEFERRED TAX ROLLFORWARD**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ TEST 4: DEFERRED TAX ROLLFORWARD                                            │
├─────────────────────────────────────────┬───────────────────────────────────┤
│ Beginning Net Deferred Tax              │ $120,000 (DTA)              ▓▓▓   │
│ Deferred Tax Expense/(Benefit)          │ ($75,000)                         │
│ OCI Items                               │ $0                          ▓▓▓   │
│ Other Adjustments                       │ $0                          ▓▓▓   │
├─────────────────────────────────────────┼───────────────────────────────────┤
│ Calculated Ending                       │ $45,000 (DTA)                     │
│ Per Balance Sheet                       │ $45,000 (DTA)               ▓▓▓   │
│ Difference                              │ $0                                │
└─────────────────────────────────────────┴───────────────────────────────────┘
  (▓▓▓ = Yellow highlight for input cells)
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                               │
├─────────────────────────────────────────────────────────────────────────────┤
│ Current Tax Expense:                  $525,000                              │
│ Deferred Tax Expense/(Benefit):       $75,000                               │
│ TOTAL TAX EXPENSE:                    $600,000                              │
├─────────────────────────────────────────────────────────────────────────────┤
│ Procedures Performed:                                                       │
│   ✓ Tax provision summary                                                   │
│   ☐ Effective tax rate analysis (manual)                                    │
│   ☐ Deferred tax analysis (manual)                                          │
│   ☐ Deferred tax rollforward (manual)                                       │
│   ☐ Valuation allowance assessment (manual)                                 │
│   ☐ Uncertain tax positions (ASC 740-10) (manual)                           │
├─────────────────────────────────────────────────────────────────────────────┤
│ CONCLUSION: [Document conclusion]                                           │
└─────────────────────────────────────────────────────────────────────────────┘
```

### VA_Assessment Worksheet

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ VALUATION ALLOWANCE ASSESSMENT                                                      │
│ Per ASC 740-10-30-16 through 30-25                                                  │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ POSITIVE EVIDENCE (More Likely Than Not to Realize)                    [GREEN]      │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ Evidence                                                        │ Present? │ Weight │
├─────────────────────────────────────────────────────────────────┼──────────┼────────┤
│ Existing contracts/backlog generating future taxable income     │ [Y/N]▓▓▓ │[H/M/L]▓│
│ History of profitable operations                                │ [Y/N]▓▓▓ │[H/M/L]▓│
│ Excess asset value over tax basis (built-in gains)              │ [Y/N]▓▓▓ │[H/M/L]▓│
│ Carryback availability                                          │ [Y/N]▓▓▓ │[H/M/L]▓│
│ Tax planning strategies that would be implemented               │ [Y/N]▓▓▓ │[H/M/L]▓│
├─────────────────────────────────────────────────────────────────────────────────────┤
│ NEGATIVE EVIDENCE (Less Likely to Realize)                         [RED]            │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ Evidence                                                        │ Present? │ Weight │
├─────────────────────────────────────────────────────────────────┼──────────┼────────┤
│ Cumulative losses in recent years                               │ [Y/N]▓▓▓ │[H/M/L]▓│
│ History of NOL/credit carryforwards expiring unused             │ [Y/N]▓▓▓ │[H/M/L]▓│
│ Losses expected in early future years                           │ [Y/N]▓▓▓ │[H/M/L]▓│
│ Unsettled circumstances that could adversely affect operations  │ [Y/N]▓▓▓ │[H/M/L]▓│
│ Brief carryforward period                                       │ [Y/N]▓▓▓ │[H/M/L]▓│
├─────────────────────────────────────────────────────────────────────────────────────┤
│ CONCLUSION                                                                          │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ Gross Deferred Tax Assets:                    $220,000                          ▓▓▓ │
│ Valuation Allowance Required:                 $0                                ▓▓▓ │
│ Net Deferred Tax Asset:                       $220,000                              │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ Management's Assessment:                                                            │
│ [Document management's position]                                                    │
└─────────────────────────────────────────────────────────────────────────────────────┘
  (▓▓▓ = Yellow highlight for input cells)
```

---

## Valuation Allowance Assessment

```vba
Sub AssessValuationAllowance()
    '================================================
    ' VALUATION ALLOWANCE ASSESSMENT (ASC 740-10-30)
    '
    ' Evaluates whether a valuation allowance is needed
    ' for deferred tax assets
    '================================================

    Dim wsAudit As Worksheet
    Dim auditRow As Long

    On Error Resume Next
    ThisWorkbook.Sheets("VA_Assessment").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "VA_Assessment"

    With wsAudit
        .Range("A1").Value = "VALUATION ALLOWANCE ASSESSMENT"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Per ASC 740-10-30-16 through 30-25"

        auditRow = 4

        ' Positive evidence
        .Cells(auditRow, 1).Value = "POSITIVE EVIDENCE (More Likely Than Not to Realize)"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(198, 239, 206)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 3)).Merge
        auditRow = auditRow + 2

        Dim posEvidence As Variant
        posEvidence = Array( _
            "Existing contracts/backlog generating future taxable income", _
            "History of profitable operations", _
            "Excess asset value over tax basis (built-in gains)", _
            "Carryback availability", _
            "Tax planning strategies that would be implemented")

        .Cells(auditRow, 1).Value = "Evidence"
        .Cells(auditRow, 2).Value = "Present?"
        .Cells(auditRow, 3).Value = "Weight"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 3)).Font.Bold = True
        auditRow = auditRow + 1

        Dim ev As Variant
        For Each ev In posEvidence
            .Cells(auditRow, 1).Value = ev
            .Cells(auditRow, 2).Value = "[Y/N]"
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 3).Value = "[H/M/L]"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next ev

        auditRow = auditRow + 1

        ' Negative evidence
        .Cells(auditRow, 1).Value = "NEGATIVE EVIDENCE (Less Likely to Realize)"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(255, 199, 206)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 3)).Merge
        auditRow = auditRow + 2

        Dim negEvidence As Variant
        negEvidence = Array( _
            "Cumulative losses in recent years", _
            "History of NOL/credit carryforwards expiring unused", _
            "Losses expected in early future years", _
            "Unsettled circumstances that could adversely affect operations", _
            "Brief carryforward period")

        .Cells(auditRow, 1).Value = "Evidence"
        .Cells(auditRow, 2).Value = "Present?"
        .Cells(auditRow, 3).Value = "Weight"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 3)).Font.Bold = True
        auditRow = auditRow + 1

        For Each ev In negEvidence
            .Cells(auditRow, 1).Value = ev
            .Cells(auditRow, 2).Value = "[Y/N]"
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 3).Value = "[H/M/L]"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next ev

        auditRow = auditRow + 2

        ' Conclusion
        .Cells(auditRow, 1).Value = "CONCLUSION"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Gross Deferred Tax Assets:"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Valuation Allowance Required:"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Net Deferred Tax Asset:"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Calc]"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Management's Assessment:"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document management's position]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 55
        .Columns("B:C").ColumnWidth = 12

    End With

    MsgBox "Valuation Allowance Assessment Template Created!", vbInformation

End Sub
```

---

## Key ASC 740 Considerations

| Topic | Guidance |
|-------|----------|
| **Current vs. Deferred** | Current = amounts payable/refundable for current year; Deferred = future tax consequences |
| **Temporary Differences** | Differences that will reverse (depreciation, allowances, accruals) |
| **Permanent Differences** | Differences that will never reverse (meals 50%, fines, municipal interest) |
| **Valuation Allowance** | Required when "more likely than not" DTA won't be realized |
| **Uncertain Tax Positions** | Two-step process: recognition then measurement |
| **Deferred Tax Rate** | Use enacted rate expected when temporary difference reverses |

---

## Common Book-Tax Differences

| Item | Book | Tax | Type | Effect |
|------|------|-----|------|--------|
| **Depreciation** | Straight-line | MACRS | Temporary | DTL |
| **Bad Debt** | Allowance | Direct write-off | Temporary | DTA |
| **Warranty** | Accrual | When paid | Temporary | DTA |
| **Prepaid Rent** | Expense | Deduction | Temporary | Either |
| **Meals** | 100% | 50% | Permanent | N/A |
| **Fines/Penalties** | Expense | Non-deductible | Permanent | N/A |
| **Stock Comp** | FV expense | When exercised | Temporary | DTA |

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Payroll](./payroll.md)
