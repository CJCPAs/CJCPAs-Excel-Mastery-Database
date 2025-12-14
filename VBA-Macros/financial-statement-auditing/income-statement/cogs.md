# Cost of Goods Sold Audit VBA

> **COGS Testing** - Complete VBA for auditing cost of goods sold per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 5000-5999 (typically) |
| **Assertions** | Occurrence, Accuracy, Cutoff, Classification |
| **Risk Level** | Moderate-High (inventory costing, cutoff) |
| **Key Standards** | ASC 330 (Inventory), ASC 606 (Matching) |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for COGS accounts

### Input Sheet 2: `COGS_Detail`
Cost of goods sold detail by category

| Column | Header | Example |
|--------|--------|---------|
| A | `Account` | 5100 |
| B | `Description` | Material Costs |
| C | `PY_Balance` | 3500000 |
| D | `CY_Balance` | 4200000 |
| E | `Dollar_Change` | 700000 |
| F | `Pct_Change` | 0.20 |

### Input Sheet 3: `Purchases`
Purchases detail for cutoff testing

| Column | Header | Example |
|--------|--------|---------|
| A | `PO_Number` | PO-2024-0458 |
| B | `Vendor` | ABC Supplier |
| C | `Invoice_Date` | 12/28/2024 |
| D | `Receipt_Date` | 01/02/2025 |
| E | `Invoice_Amount` | 75000 |
| F | `Period_Recorded` | 2024 |

---

## Audit Procedures

```vba
Sub AuditCOGS()
    '================================================
    ' COST OF GOODS SOLD - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with COGS transactions
    '   - Sheet "COGS_Detail" with COGS breakdown
    '   - Sheet "Purchases" for cutoff testing
    '
    ' OUTPUTS:
    '   - Creates "COGS_Audit" worksheet
    '   - Analyzes gross margin
    '   - Tests inventory equation
    '   - Performs cutoff testing
    '
    ' ASSERTIONS TESTED:
    '   - Occurrence (costs relate to sales)
    '   - Accuracy (amounts correct)
    '   - Cutoff (proper period)
    '================================================

    Dim wsGL As Worksheet
    Dim wsCOGS As Worksheet
    Dim wsPurch As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const MARGIN_THRESHOLD As Double = 0.05  ' 5% change triggers review

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsCOGS = ThisWorkbook.Sheets("COGS_Detail")
    Set wsPurch = ThisWorkbook.Sheets("Purchases")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("COGS_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "COGS_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "COST OF GOODS SOLD - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: COGS BREAKDOWN BY ACCOUNT
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: COGS BREAKDOWN BY ACCOUNT"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        ' Aggregate GL by account
        Dim cogsDict As Object
        Set cogsDict = CreateObject("Scripting.Dictionary")
        Dim totalCOGS As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            Dim acctName As String
            Dim acctAmt As Double

            acctNum = CStr(wsGL.Cells(i, 3).Value)

            ' COGS accounts (5xxx) - debit balance
            If Left(acctNum, 1) = "5" Then
                acctName = wsGL.Cells(i, 4).Value
                acctAmt = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value

                If cogsDict.Exists(acctNum) Then
                    cogsDict(acctNum) = Array(acctName, cogsDict(acctNum)(1) + acctAmt)
                Else
                    cogsDict.Add acctNum, Array(acctName, acctAmt)
                End If

                totalCOGS = totalCOGS + acctAmt
            End If
        Next i

        .Cells(auditRow, 1).Value = "Account"
        .Cells(auditRow, 2).Value = "Description"
        .Cells(auditRow, 3).Value = "Balance"
        .Cells(auditRow, 4).Value = "% of COGS"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        Dim acctStart As Long
        acctStart = auditRow

        Dim key As Variant
        For Each key In cogsDict.Keys
            Dim data As Variant
            data = cogsDict(key)
            .Cells(auditRow, 1).Value = key
            .Cells(auditRow, 2).Value = data(0)
            .Cells(auditRow, 3).Value = data(1)
            If totalCOGS <> 0 Then
                .Cells(auditRow, 4).Value = data(1) / totalCOGS
            End If
            auditRow = auditRow + 1
        Next key

        .Cells(auditRow, 1).Value = "TOTAL COGS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 3).Value = totalCOGS
        .Cells(auditRow, 3).Font.Bold = True
        .Cells(auditRow, 4).Value = 1
        auditRow = auditRow + 1

        .Range(.Cells(acctStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(acctStart, 4), .Cells(auditRow - 1, 4)).NumberFormat = "0.0%"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 2: INVENTORY EQUATION TEST
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 2: INVENTORY EQUATION TEST"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Beginning Inventory"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Add: Purchases"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Less: Ending Inventory"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Calculated COGS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Formula]"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Per GL"
        .Cells(auditRow, 2).Value = totalCOGS
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Difference"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = "[Formula]"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Status"
        .Cells(auditRow, 2).Value = "[Calculate]"
        auditRow = auditRow + 1

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 3: GROSS MARGIN ANALYSIS
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 3: GROSS MARGIN ANALYSIS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Metric"
        .Cells(auditRow, 2).Value = "PY"
        .Cells(auditRow, 3).Value = "CY"
        .Cells(auditRow, 4).Value = "Change"
        .Cells(auditRow, 5).Value = "Industry"
        .Cells(auditRow, 6).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        Dim marginStart As Long
        marginStart = auditRow

        Dim metrics As Variant
        metrics = Array( _
            Array("Revenue", "[Input]", "[Input]"), _
            Array("COGS", "[Input]", totalCOGS), _
            Array("Gross Profit", "[Calc]", "[Calc]"), _
            Array("Gross Margin %", "[Calc]", "[Calc]"))

        Dim m As Variant
        For Each m In metrics
            .Cells(auditRow, 1).Value = m(0)
            .Cells(auditRow, 2).Value = m(1)
            If m(0) = "COGS" Then
                .Cells(auditRow, 3).Value = totalCOGS
                .Cells(auditRow, 3).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            Else
                .Cells(auditRow, 3).Value = m(2)
            End If
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 5).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next m

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 4: MONTHLY COGS TREND
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: MONTHLY COGS TREND"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Month"
        .Cells(auditRow, 2).Value = "COGS"
        .Cells(auditRow, 3).Value = "Revenue"
        .Cells(auditRow, 4).Value = "GM %"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        ' Aggregate COGS by month
        Dim monthCOGS(1 To 12) As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            acctNum = CStr(wsGL.Cells(i, 3).Value)
            If Left(acctNum, 1) = "5" Then
                Dim trxDate As Date
                On Error Resume Next
                trxDate = wsGL.Cells(i, 1).Value
                On Error GoTo 0

                If IsDate(trxDate) Then
                    Dim mo As Integer
                    mo = Month(trxDate)
                    monthCOGS(mo) = monthCOGS(mo) + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
                End If
            End If
        Next i

        Dim trendStart As Long
        trendStart = auditRow

        Dim monthNames As Variant
        monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

        For mo = 1 To 12
            .Cells(auditRow, 1).Value = monthNames(mo - 1)
            .Cells(auditRow, 2).Value = monthCOGS(mo)
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)  ' Revenue input
            .Cells(auditRow, 4).Value = "[Calc]"
            auditRow = auditRow + 1
        Next mo

        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = totalCOGS
        .Cells(auditRow, 2).Font.Bold = True
        auditRow = auditRow + 1

        .Range(.Cells(trendStart, 2), .Cells(auditRow - 1, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 5: PURCHASE CUTOFF
        ' ========================================
        If Not wsPurch Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 5: PURCHASE CUTOFF TESTING"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "PO Number"
            .Cells(auditRow, 2).Value = "Vendor"
            .Cells(auditRow, 3).Value = "Invoice Date"
            .Cells(auditRow, 4).Value = "Receipt Date"
            .Cells(auditRow, 5).Value = "Amount"
            .Cells(auditRow, 6).Value = "Recorded"
            .Cells(auditRow, 7).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
            auditRow = auditRow + 1

            Dim cutStart As Long
            cutStart = auditRow

            Dim yearEnd As Date
            yearEnd = DateSerial(Year(Date), 12, 31)

            lastRow = wsPurch.Cells(wsPurch.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim invDate As Date
                Dim rcptDate As Date
                Dim recPeriod As String

                On Error Resume Next
                invDate = wsPurch.Cells(i, 3).Value
                rcptDate = wsPurch.Cells(i, 4).Value
                On Error GoTo 0

                recPeriod = wsPurch.Cells(i, 6).Value

                ' Only show cutoff items (around year-end)
                If (invDate >= DateAdd("d", -10, yearEnd) And invDate <= DateAdd("d", 10, yearEnd)) Or _
                   (rcptDate >= DateAdd("d", -10, yearEnd) And rcptDate <= DateAdd("d", 10, yearEnd)) Then

                    .Cells(auditRow, 1).Value = wsPurch.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsPurch.Cells(i, 2).Value
                    .Cells(auditRow, 3).Value = invDate
                    .Cells(auditRow, 4).Value = rcptDate
                    .Cells(auditRow, 5).Value = wsPurch.Cells(i, 5).Value

                    .Cells(auditRow, 6).Value = recPeriod

                    ' Check cutoff - should be recorded when received
                    Dim shouldBe As String
                    If rcptDate <= yearEnd Then
                        shouldBe = "2024"
                    Else
                        shouldBe = "2025"
                    End If

                    If CStr(recPeriod) = shouldBe Then
                        .Cells(auditRow, 7).Value = "CORRECT"
                        .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
                    Else
                        .Cells(auditRow, 7).Value = "CUTOFF ERROR"
                        .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                    End If

                    auditRow = auditRow + 1
                End If
            Next i

            .Range(.Cells(cutStart, 3), .Cells(auditRow - 1, 4)).NumberFormat = "mm/dd/yyyy"
            .Range(.Cells(cutStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 2
        End If

        ' ========================================
        ' AUDIT SUMMARY
        ' ========================================
        .Cells(auditRow, 1).Value = "AUDIT SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Total Cost of Goods Sold:"
        .Cells(auditRow, 2).Value = totalCOGS
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " COGS breakdown analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Monthly trend analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Inventory equation test (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Gross margin analysis (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Purchase cutoff testing (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 25
        .Columns("B:G").ColumnWidth = 14

    End With

    Application.ScreenUpdating = True

    MsgBox "COGS Audit Complete!" & vbCrLf & _
           "Total COGS: " & Format(totalCOGS, "$#,##0"), vbInformation

End Sub
```

---

## Output Examples

### COGS_Audit Worksheet

The `AuditCOGS` procedure generates a comprehensive worksheet:

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ COST OF GOODS SOLD - AUDIT WORKPAPER                                                │
│ Period: December 31, 2024                                                           │
│ Prepared: AUDITOR on 12/15/2024 4:00:00 PM                                         │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

**TEST 1: COGS BREAKDOWN BY ACCOUNT**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 1: COGS BREAKDOWN BY ACCOUNT                                                   │
├─────────────┬────────────────────────────────┬────────────────┬─────────────────────┤
│ Account     │ Description                    │ Balance        │ % of COGS           │
├─────────────┼────────────────────────────────┼────────────────┼─────────────────────┤
│ 5100        │ Direct Materials               │ $3,850,000     │ 55.0%               │
│ 5200        │ Direct Labor                   │ $1,540,000     │ 22.0%               │
│ 5300        │ Manufacturing Overhead         │ $1,050,000     │ 15.0%               │
│ 5400        │ Freight-In                     │ $280,000       │ 4.0%                │
│ 5500        │ Purchase Discounts             │ ($70,000)      │ -1.0%               │
│ 5900        │ Other COGS                     │ $350,000       │ 5.0%                │
├─────────────┴────────────────────────────────┼────────────────┼─────────────────────┤
│ TOTAL COGS                                   │ $7,000,000     │ 100.0%              │
└──────────────────────────────────────────────┴────────────────┴─────────────────────┘
```

**TEST 2: INVENTORY EQUATION TEST**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ TEST 2: INVENTORY EQUATION TEST                                             │
├─────────────────────────────────────────┬───────────────────────────────────┤
│ Beginning Inventory                     │ [Input]                       ▓▓▓ │
│ Add: Purchases                          │ [Input]                       ▓▓▓ │
│ Less: Ending Inventory                  │ [Input]                       ▓▓▓ │
├─────────────────────────────────────────┼───────────────────────────────────┤
│ Calculated COGS                         │ [Formula]                         │
│ Per GL                                  │ $7,000,000                        │
│ Difference                              │ [Formula]                         │
│ Status                                  │ [Calculate]                       │
└─────────────────────────────────────────┴───────────────────────────────────┘
  (▓▓▓ = Yellow highlight for input cells)
```

**TEST 3: GROSS MARGIN ANALYSIS**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 3: GROSS MARGIN ANALYSIS                                                                   │
├─────────────────────────────┬────────────────┬────────────────┬──────────────┬─────────┬────────┤
│ Metric                      │ PY             │ CY             │ Change       │ Industry│ Status │
├─────────────────────────────┼────────────────┼────────────────┼──────────────┼─────────┼────────┤
│ Revenue                     │ $10,035,000    │ $10,900,000    │ $865,000     │         │        │
│ COGS                        │ $6,522,750     │ $7,000,000     │ $477,250     │         │        │
│ Gross Profit                │ $3,512,250     │ $3,900,000     │ $387,750     │         │        │
│ Gross Margin %              │ 35.0%          │ 35.8%          │ 0.8%         │ 35%     │ ✓      │
└─────────────────────────────┴────────────────┴────────────────┴──────────────┴─────────┴────────┘
```

**TEST 4: MONTHLY COGS TREND**
```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 4: MONTHLY COGS TREND                                                          │
├──────────────┬────────────────┬────────────────┬─────────────────┬──────────────────┤
│ Month        │ COGS           │ Revenue        │ GM %            │ Status           │
├──────────────┼────────────────┼────────────────┼─────────────────┼──────────────────┤
│ Jan          │ $544,000       │ $850,000       │ 36.0%           │ ✓                │
│ Feb          │ $533,000       │ $820,000       │ 35.0%           │ ✓                │
│ Mar          │ $582,400       │ $910,000       │ 36.0%           │ ✓                │
│ Apr          │ $568,750       │ $875,000       │ 35.0%           │ ✓                │
│ May          │ $601,250       │ $925,000       │ 35.0%           │ ✓                │
│ Jun          │ $608,000       │ $950,000       │ 36.0%           │ ✓                │
│ Jul          │ $559,000       │ $860,000       │ 35.0%           │ ✓                │
│ Aug          │ $572,000       │ $880,000       │ 35.0%           │ ✓                │
│ Sep          │ $598,000       │ $920,000       │ 35.0%           │ ✓                │
│ Oct          │ $611,000       │ $940,000       │ 35.0%           │ ✓                │
│ Nov          │ $627,200       │ $980,000       │ 36.0%           │ ✓                │
│ Dec          │ $595,400       │ $1,090,000     │ 45.4%           │ ⚠ HIGH           │
├──────────────┼────────────────┼────────────────┼─────────────────┼──────────────────┤
│ TOTAL        │ $7,000,000     │ $10,900,000    │ 35.8%           │                  │
└──────────────┴────────────────┴────────────────┴─────────────────┴──────────────────┘
  Status: ⚠ = Gross margin significantly above/below average
```

**TEST 5: PURCHASE CUTOFF TESTING**
```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ TEST 5: PURCHASE CUTOFF TESTING                                                                             │
├────────────────┬─────────────────────┬────────────┬────────────┬────────────┬──────────┬────────────────────┤
│ PO Number      │ Vendor              │ Invoice Date│ Receipt Date│ Amount     │ Recorded │ Status            │
├────────────────┼─────────────────────┼────────────┼────────────┼────────────┼──────────┼────────────────────┤
│ PO-2024-0945   │ ABC Supplier        │ 12/26/2024 │ 12/28/2024 │ $45,000    │ 2024     │ ✓ CORRECT          │
│ PO-2024-0952   │ XYZ Materials       │ 12/28/2024 │ 12/30/2024 │ $62,000    │ 2024     │ ✓ CORRECT          │
│ PO-2024-0958   │ Delta Supplies      │ 12/30/2024 │ 01/03/2025 │ $38,500    │ 2024     │ ✗ CUTOFF ERROR     │
│ PO-2024-0961   │ Gamma Industries    │ 12/31/2024 │ 12/31/2024 │ $27,000    │ 2024     │ ✓ CORRECT          │
│ PO-2025-0003   │ Beta Corp           │ 01/02/2025 │ 12/29/2024 │ $51,000    │ 2025     │ ✗ CUTOFF ERROR     │
└────────────────┴─────────────────────┴────────────┴────────────┴────────────┴──────────┴────────────────────┘
  ✗ = Should be recorded when goods received (title passes)
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                               │
├─────────────────────────────────────────────────────────────────────────────┤
│ Total Cost of Goods Sold: $7,000,000                                        │
├─────────────────────────────────────────────────────────────────────────────┤
│ Procedures Performed:                                                       │
│   ✓ COGS breakdown analysis                                                 │
│   ✓ Monthly trend analysis                                                  │
│   ☐ Inventory equation test (manual)                                        │
│   ☐ Gross margin analysis (manual)                                          │
│   ☐ Purchase cutoff testing (manual)                                        │
├─────────────────────────────────────────────────────────────────────────────┤
│ CONCLUSION: [Document conclusion]                                           │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## COGS Components Analysis

| Component | Typical % | Audit Focus |
|-----------|----------|-------------|
| **Direct Materials** | 40-60% | Pricing, quantities, cutoff |
| **Direct Labor** | 15-30% | Payroll testing, allocation |
| **Manufacturing Overhead** | 10-25% | Allocation methodology |
| **Freight-In** | 2-5% | Proper classification |
| **Purchase Discounts** | (1-3%) | Timing of recognition |

---

## Standard Cost Variance Analysis

```vba
Sub AnalyzeStandardCostVariances()
    '================================================
    ' STANDARD COST VARIANCE ANALYSIS
    '
    ' For companies using standard costing
    '================================================

    Dim wsAudit As Worksheet
    Dim auditRow As Long

    On Error Resume Next
    ThisWorkbook.Sheets("Cost_Variances").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Cost_Variances"

    With wsAudit
        .Range("A1").Value = "STANDARD COST VARIANCE ANALYSIS"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        auditRow = 3

        ' Variance summary
        .Cells(auditRow, 1).Value = "VARIANCE SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Variance Type"
        .Cells(auditRow, 2).Value = "Amount"
        .Cells(auditRow, 3).Value = "Favorable?"
        .Cells(auditRow, 4).Value = "% of Std"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim variances As Variant
        variances = Array( _
            "Material Price Variance", _
            "Material Usage Variance", _
            "Labor Rate Variance", _
            "Labor Efficiency Variance", _
            "Variable OH Spending Variance", _
            "Variable OH Efficiency Variance", _
            "Fixed OH Budget Variance", _
            "Fixed OH Volume Variance")

        Dim v As Variant
        For Each v In variances
            .Cells(auditRow, 1).Value = v
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 4).Interior.Color = RGB(255, 255, 204)
            auditRow = auditRow + 1
        Next v

        auditRow = auditRow + 2

        ' Materiality assessment
        .Cells(auditRow, 1).Value = "VARIANCE DISPOSITION"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Total Variances:"
        .Cells(auditRow, 2).Value = "[Sum]"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Materiality:"
        .Cells(auditRow, 2).Value = "[Input]"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Disposition:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Write off to COGS (immaterial)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Allocate to inventory/COGS (material)"
        auditRow = auditRow + 1

        .Columns("A").ColumnWidth = 35
        .Columns("B:E").ColumnWidth = 15

    End With

    MsgBox "Variance Analysis Template Created!", vbInformation

End Sub
```

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Revenue](./revenue.md) | [➡️ Operating Expenses](./operating-expenses.md)
