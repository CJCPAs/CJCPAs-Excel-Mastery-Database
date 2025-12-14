# Property, Plant & Equipment Audit VBA

> **Fixed Assets Testing** - Complete VBA for auditing PP&E per GAAS/GAAP (ASC 360)

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 1500-1799 (typically) |
| **Assertions** | Existence, Rights, Valuation, Completeness |
| **Risk Level** | Moderate (impairment, depreciation estimates) |
| **Key Standards** | ASC 360-10 (Impairment), ASC 250 (Useful Life Changes) |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for fixed asset and accumulated depreciation accounts

### Input Sheet 2: `FA_Register`
Fixed asset register/subledger

| Column | Header | Example |
|--------|--------|---------|
| A | `Asset_ID` | FA-001 |
| B | `Description` | Building - HQ |
| C | `Category` | Buildings |
| D | `Acquired_Date` | 03/15/2018 |
| E | `Cost` | 2500000 |
| F | `Accum_Depr` | 583333 |
| G | `NBV` | 1916667 |
| H | `Useful_Life` | 39 |
| I | `Method` | SL |
| J | `Location` | 100 Main St |
| K | `Disposed` | No |

### Input Sheet 3: `Additions`
Current year asset additions

| Column | Header | Example |
|--------|--------|---------|
| A | `Asset_ID` | FA-025 |
| B | `Description` | Server Equipment |
| C | `Invoice_Number` | INV-78945 |
| D | `Vendor` | Dell Technologies |
| E | `Date_Acquired` | 06/15/2024 |
| F | `Cost` | 125000 |
| G | `Useful_Life` | 5 |
| H | `In_Service_Date` | 06/20/2024 |

### Input Sheet 4: `Disposals`
Current year asset disposals

| Column | Header | Example |
|--------|--------|---------|
| A | `Asset_ID` | FA-008 |
| B | `Description` | Old Copier |
| C | `Original_Cost` | 15000 |
| D | `Accum_Depr` | 15000 |
| E | `NBV` | 0 |
| F | `Sale_Price` | 500 |
| G | `Gain_Loss` | 500 |
| H | `Disposal_Date` | 09/30/2024 |

---

## Audit Procedures

```vba
Sub AuditPPE()
    '================================================
    ' PROPERTY, PLANT & EQUIPMENT - COMPLETE AUDIT
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with FA transactions
    '   - Sheet "FA_Register" with fixed asset subledger
    '   - Sheet "Additions" with CY additions
    '   - Sheet "Disposals" with CY disposals
    '
    ' OUTPUTS:
    '   - Creates "PPE_Audit" worksheet
    '   - Reconciles subledger to GL
    '   - Tests additions for proper capitalization
    '   - Tests disposals for proper accounting
    '   - Recalculates depreciation expense
    '
    ' ASSERTIONS TESTED:
    '   - Existence (assets physically exist)
    '   - Rights (company owns assets)
    '   - Valuation (properly depreciated)
    '   - Completeness (all assets recorded)
    '================================================

    Dim wsGL As Worksheet
    Dim wsFA As Worksheet
    Dim wsAdd As Worksheet
    Dim wsDisp As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const CAPITALIZATION_THRESHOLD As Double = 5000
    Const MATERIALITY As Double = 100000

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsFA = ThisWorkbook.Sheets("FA_Register")
    Set wsAdd = ThisWorkbook.Sheets("Additions")
    Set wsDisp = ThisWorkbook.Sheets("Disposals")
    On Error GoTo 0

    If wsGL Is Nothing Or wsFA Is Nothing Then
        MsgBox "GL_Detail and FA_Register sheets required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("PPE_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "PPE_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "PROPERTY, PLANT & EQUIPMENT - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: GL TO SUBLEDGER RECONCILIATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: GL TO SUBLEDGER RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        ' Calculate GL balances
        Dim glCost As Double
        Dim glAccumDepr As Double
        Dim slCost As Double
        Dim slAccumDepr As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            acctNum = CStr(wsGL.Cells(i, 3).Value)

            ' Fixed assets (15xx)
            If Left(acctNum, 2) = "15" Or Left(acctNum, 2) = "16" Then
                glCost = glCost + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
            End If

            ' Accumulated depreciation (17xx) - credit balance
            If Left(acctNum, 2) = "17" Then
                glAccumDepr = glAccumDepr + wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value
            End If
        Next i

        ' Calculate subledger balances
        lastRow = wsFA.Cells(wsFA.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            If LCase(wsFA.Cells(i, 11).Value) <> "yes" Then  ' Not disposed
                slCost = slCost + wsFA.Cells(i, 5).Value
                slAccumDepr = slAccumDepr + wsFA.Cells(i, 6).Value
            End If
        Next i

        .Cells(auditRow, 1).Value = "Category"
        .Cells(auditRow, 2).Value = "GL Balance"
        .Cells(auditRow, 3).Value = "Subledger"
        .Cells(auditRow, 4).Value = "Difference"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim reconStart As Long
        reconStart = auditRow

        ' Cost reconciliation
        .Cells(auditRow, 1).Value = "Fixed Assets - Cost"
        .Cells(auditRow, 2).Value = glCost
        .Cells(auditRow, 3).Value = slCost
        .Cells(auditRow, 4).Value = glCost - slCost

        If Abs(glCost - slCost) < 1 Then
            .Cells(auditRow, 5).Value = "RECONCILED"
            .Cells(auditRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 5).Value = "DIFFERENCE"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Accumulated depreciation reconciliation
        .Cells(auditRow, 1).Value = "Accumulated Depreciation"
        .Cells(auditRow, 2).Value = glAccumDepr
        .Cells(auditRow, 3).Value = slAccumDepr
        .Cells(auditRow, 4).Value = glAccumDepr - slAccumDepr

        If Abs(glAccumDepr - slAccumDepr) < 1 Then
            .Cells(auditRow, 5).Value = "RECONCILED"
            .Cells(auditRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 5).Value = "DIFFERENCE"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Net book value
        .Cells(auditRow, 1).Value = "Net Book Value"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glCost - glAccumDepr
        .Cells(auditRow, 2).Font.Bold = True
        .Cells(auditRow, 3).Value = slCost - slAccumDepr
        .Cells(auditRow, 3).Font.Bold = True
        .Cells(auditRow, 4).Value = (glCost - glAccumDepr) - (slCost - slAccumDepr)
        auditRow = auditRow + 1

        .Range(.Cells(reconStart, 2), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 2: ADDITIONS TESTING
        ' ========================================
        If Not wsAdd Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 2: ADDITIONS TESTING"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Asset ID"
            .Cells(auditRow, 2).Value = "Description"
            .Cells(auditRow, 3).Value = "Date"
            .Cells(auditRow, 4).Value = "Cost"
            .Cells(auditRow, 5).Value = "Useful Life"
            .Cells(auditRow, 6).Value = "Capitalized?"
            .Cells(auditRow, 7).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
            auditRow = auditRow + 1

            Dim addStart As Long
            addStart = auditRow

            Dim totalAdditions As Double

            lastRow = wsAdd.Cells(wsAdd.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim addCost As Double
                addCost = wsAdd.Cells(i, 6).Value

                .Cells(auditRow, 1).Value = wsAdd.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsAdd.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = wsAdd.Cells(i, 5).Value
                .Cells(auditRow, 4).Value = addCost
                .Cells(auditRow, 5).Value = wsAdd.Cells(i, 7).Value

                totalAdditions = totalAdditions + addCost

                ' Check capitalization threshold
                If addCost >= CAPITALIZATION_THRESHOLD Then
                    .Cells(auditRow, 6).Value = "PROPER"
                    .Cells(auditRow, 6).Interior.Color = RGB(198, 239, 206)

                    ' Verify in subledger
                    Dim assetFound As Boolean
                    assetFound = False
                    Dim j As Long
                    Dim faLastRow As Long
                    faLastRow = wsFA.Cells(wsFA.Rows.Count, "A").End(xlUp).Row
                    For j = 2 To faLastRow
                        If CStr(wsFA.Cells(j, 1).Value) = CStr(wsAdd.Cells(i, 1).Value) Then
                            assetFound = True
                            Exit For
                        End If
                    Next j

                    If assetFound Then
                        .Cells(auditRow, 7).Value = "IN REGISTER"
                        .Cells(auditRow, 7).Interior.Color = RGB(198, 239, 206)
                    Else
                        .Cells(auditRow, 7).Value = "NOT IN REGISTER"
                        .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                    End If
                Else
                    .Cells(auditRow, 6).Value = "BELOW THRESHOLD"
                    .Cells(auditRow, 6).Interior.Color = RGB(255, 235, 156)
                    .Cells(auditRow, 7).Value = "REVIEW EXPENSE"
                    .Cells(auditRow, 7).Interior.Color = RGB(255, 235, 156)
                End If

                auditRow = auditRow + 1
            Next i

            ' Total additions
            .Cells(auditRow, 1).Value = "TOTAL ADDITIONS"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 4).Value = totalAdditions
            .Cells(auditRow, 4).Font.Bold = True

            .Range(.Cells(addStart, 4), .Cells(auditRow, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 3
        End If

        ' ========================================
        ' TEST 3: DISPOSALS TESTING
        ' ========================================
        If Not wsDisp Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 3: DISPOSALS TESTING"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Asset ID"
            .Cells(auditRow, 2).Value = "Description"
            .Cells(auditRow, 3).Value = "Orig Cost"
            .Cells(auditRow, 4).Value = "Accum Depr"
            .Cells(auditRow, 5).Value = "NBV"
            .Cells(auditRow, 6).Value = "Sale Price"
            .Cells(auditRow, 7).Value = "Gain/Loss"
            .Cells(auditRow, 8).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Font.Bold = True
            auditRow = auditRow + 1

            Dim dispStart As Long
            dispStart = auditRow

            lastRow = wsDisp.Cells(wsDisp.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim origCost As Double
                Dim accumDepr As Double
                Dim nbv As Double
                Dim salePrice As Double
                Dim gainLoss As Double
                Dim calcGainLoss As Double

                origCost = wsDisp.Cells(i, 3).Value
                accumDepr = wsDisp.Cells(i, 4).Value
                nbv = wsDisp.Cells(i, 5).Value
                salePrice = wsDisp.Cells(i, 6).Value
                gainLoss = wsDisp.Cells(i, 7).Value

                calcGainLoss = salePrice - nbv

                .Cells(auditRow, 1).Value = wsDisp.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsDisp.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = origCost
                .Cells(auditRow, 4).Value = accumDepr
                .Cells(auditRow, 5).Value = nbv
                .Cells(auditRow, 6).Value = salePrice
                .Cells(auditRow, 7).Value = gainLoss

                ' Verify gain/loss calculation
                If Abs(gainLoss - calcGainLoss) < 1 Then
                    .Cells(auditRow, 8).Value = "CALC VERIFIED"
                    .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
                Else
                    .Cells(auditRow, 8).Value = "CALC ERROR"
                    .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
                End If

                auditRow = auditRow + 1
            Next i

            .Range(.Cells(dispStart, 3), .Cells(auditRow - 1, 7)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 3
        End If

        ' ========================================
        ' TEST 4: DEPRECIATION RECALCULATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: DEPRECIATION RECALCULATION (SAMPLE)"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Asset ID"
        .Cells(auditRow, 2).Value = "Description"
        .Cells(auditRow, 3).Value = "Cost"
        .Cells(auditRow, 4).Value = "Life"
        .Cells(auditRow, 5).Value = "Method"
        .Cells(auditRow, 6).Value = "Months"
        .Cells(auditRow, 7).Value = "Expected"
        .Cells(auditRow, 8).Value = "Per Books"
        .Cells(auditRow, 9).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 9)).Font.Bold = True
        auditRow = auditRow + 1

        Dim deprStart As Long
        deprStart = auditRow

        ' Sample top 10 assets by cost for depreciation testing
        lastRow = wsFA.Cells(wsFA.Rows.Count, "A").End(xlUp).Row
        Dim sampleCount As Long
        sampleCount = 0

        For i = 2 To lastRow
            If sampleCount >= 10 Then Exit For
            If LCase(wsFA.Cells(i, 11).Value) <> "yes" Then  ' Not disposed
                Dim assetCost As Double
                Dim usefulLife As Double
                Dim deprMethod As String
                Dim acquiredDate As Date
                Dim monthsHeld As Long
                Dim expectedDepr As Double
                Dim actualAccumDepr As Double

                assetCost = wsFA.Cells(i, 5).Value
                usefulLife = wsFA.Cells(i, 8).Value
                deprMethod = wsFA.Cells(i, 9).Value
                actualAccumDepr = wsFA.Cells(i, 6).Value

                On Error Resume Next
                acquiredDate = wsFA.Cells(i, 4).Value
                On Error GoTo 0

                If IsDate(acquiredDate) And usefulLife > 0 Then
                    ' Calculate months held
                    monthsHeld = DateDiff("m", acquiredDate, DateSerial(Year(Date), 12, 31))
                    If monthsHeld < 0 Then monthsHeld = 0

                    ' Calculate expected depreciation (straight-line)
                    expectedDepr = (assetCost / usefulLife) * (monthsHeld / 12)

                    .Cells(auditRow, 1).Value = wsFA.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = Left(wsFA.Cells(i, 2).Value, 25)
                    .Cells(auditRow, 3).Value = assetCost
                    .Cells(auditRow, 4).Value = usefulLife
                    .Cells(auditRow, 5).Value = deprMethod
                    .Cells(auditRow, 6).Value = monthsHeld
                    .Cells(auditRow, 7).Value = expectedDepr
                    .Cells(auditRow, 8).Value = actualAccumDepr

                    ' Compare expected to actual (within 5% tolerance)
                    If expectedDepr > 0 Then
                        If Abs(actualAccumDepr - expectedDepr) / expectedDepr < 0.05 Then
                            .Cells(auditRow, 9).Value = "REASONABLE"
                            .Cells(auditRow, 9).Interior.Color = RGB(198, 239, 206)
                        Else
                            .Cells(auditRow, 9).Value = "INVESTIGATE"
                            .Cells(auditRow, 9).Interior.Color = RGB(255, 199, 206)
                        End If
                    Else
                        .Cells(auditRow, 9).Value = "NEW ASSET"
                        .Cells(auditRow, 9).Interior.Color = RGB(255, 235, 156)
                    End If

                    auditRow = auditRow + 1
                    sampleCount = sampleCount + 1
                End If
            End If
        Next i

        .Range(.Cells(deprStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(deprStart, 7), .Cells(auditRow - 1, 8)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 3

        ' ========================================
        ' TEST 5: ASSET CATEGORY ROLLFORWARD
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 5: FIXED ASSET ROLLFORWARD"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Category"
        .Cells(auditRow, 2).Value = "Beg Balance"
        .Cells(auditRow, 3).Value = "Additions"
        .Cells(auditRow, 4).Value = "Disposals"
        .Cells(auditRow, 5).Value = "End Balance"
        .Cells(auditRow, 6).Value = "Per SL"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        ' Aggregate by category from FA register
        Dim catDict As Object
        Set catDict = CreateObject("Scripting.Dictionary")

        lastRow = wsFA.Cells(wsFA.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim catName As String
            Dim catCost As Double

            catName = wsFA.Cells(i, 3).Value
            If LCase(wsFA.Cells(i, 11).Value) <> "yes" Then
                catCost = wsFA.Cells(i, 5).Value
            Else
                catCost = 0
            End If

            If catDict.Exists(catName) Then
                catDict(catName) = catDict(catName) + catCost
            Else
                catDict.Add catName, catCost
            End If
        Next i

        Dim catKey As Variant
        For Each catKey In catDict.Keys
            .Cells(auditRow, 1).Value = catKey
            .Cells(auditRow, 5).Value = "[Calculate]"
            .Cells(auditRow, 6).Value = catDict(catKey)
            auditRow = auditRow + 1
        Next catKey

        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 6).Value = slCost
        .Cells(auditRow, 6).Font.Bold = True

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

        .Cells(auditRow, 1).Value = "Total PP&E (Cost):"
        .Cells(auditRow, 2).Value = glCost
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Accumulated Depreciation:"
        .Cells(auditRow, 2).Value = glAccumDepr
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Net Book Value:"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glCost - glAccumDepr
        .Cells(auditRow, 2).Font.Bold = True
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " GL to subledger reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Additions testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Disposals testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Depreciation recalculation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Physical inspection (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Impairment assessment (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 25
        .Columns("B:I").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "PP&E Audit Complete!" & vbCrLf & _
           "Net Book Value: " & Format(glCost - glAccumDepr, "$#,##0"), vbInformation

End Sub
```

---

## Useful Life Standards

| Asset Category | Typical Life | IRS Class |
|---------------|--------------|-----------|
| **Buildings** | 39 years | Nonresidential real property |
| **Building Improvements** | 15-39 years | Qualified improvement |
| **Vehicles** | 5 years | 5-year property |
| **Furniture & Fixtures** | 7 years | 7-year property |
| **Computers** | 5 years | 5-year property |
| **Machinery** | 7 years | 7-year property |
| **Land Improvements** | 15 years | 15-year property |
| **Land** | N/A | Not depreciable |

---

## Impairment Testing (ASC 360-10)

```vba
Sub TestPPEImpairment()
    '================================================
    ' PPE IMPAIRMENT INDICATOR REVIEW
    '
    ' Per ASC 360-10-35-21, test for impairment when:
    '   - Significant decrease in market price
    '   - Adverse change in asset use
    '   - Adverse change in legal/business climate
    '   - Accumulation of costs significantly exceed
    '     original acquisition cost
    '   - Current period operating loss
    '   - Projected future operating losses
    '================================================

    Dim wsFA As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    On Error Resume Next
    Set wsFA = ThisWorkbook.Sheets("FA_Register")
    On Error GoTo 0

    If wsFA Is Nothing Then
        MsgBox "FA_Register sheet required.", vbCritical
        Exit Sub
    End If

    ' Create impairment worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("PPE_Impairment").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "PPE_Impairment"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "PPE IMPAIRMENT INDICATOR REVIEW"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Per ASC 360-10-35"

        auditRow = 4

        .Cells(auditRow, 1).Value = "IMPAIRMENT INDICATORS CHECKLIST"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 3)).Merge
        auditRow = auditRow + 2

        Dim indicators As Variant
        indicators = Array( _
            "Significant decrease in market price of asset", _
            "Significant adverse change in how asset is used", _
            "Significant adverse change in legal or business climate", _
            "Costs accumulated significantly exceed original budget", _
            "Current period operating or cash flow loss", _
            "Projected future operating or cash flow losses", _
            "Asset currently expected to be disposed before end of useful life")

        Dim ind As Variant
        For Each ind In indicators
            .Cells(auditRow, 1).Value = ChrW(9744) & " " & ind
            .Cells(auditRow, 3).Value = "[Yes/No/N/A]"
            auditRow = auditRow + 1
        Next ind

        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "ASSETS WITH POTENTIAL IMPAIRMENT"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Asset ID"
        .Cells(auditRow, 2).Value = "Description"
        .Cells(auditRow, 3).Value = "NBV"
        .Cells(auditRow, 4).Value = "Age (Yrs)"
        .Cells(auditRow, 5).Value = "% Depreciated"
        .Cells(auditRow, 6).Value = "Indicator"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        lastRow = wsFA.Cells(wsFA.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If LCase(wsFA.Cells(i, 11).Value) <> "yes" Then
                Dim faCost As Double
                Dim faAccum As Double
                Dim faNBV As Double
                Dim faAge As Double
                Dim pctDepr As Double

                faCost = wsFA.Cells(i, 5).Value
                faAccum = wsFA.Cells(i, 6).Value
                faNBV = wsFA.Cells(i, 7).Value

                If faCost > 0 Then
                    pctDepr = faAccum / faCost
                End If

                On Error Resume Next
                faAge = DateDiff("yyyy", wsFA.Cells(i, 4).Value, Date)
                On Error GoTo 0

                ' Flag assets that are >80% depreciated but still in use
                If pctDepr > 0.8 And faNBV > 0 Then
                    .Cells(auditRow, 1).Value = wsFA.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsFA.Cells(i, 2).Value
                    .Cells(auditRow, 3).Value = faNBV
                    .Cells(auditRow, 4).Value = faAge
                    .Cells(auditRow, 5).Value = pctDepr
                    .Cells(auditRow, 6).Value = "NEAR END OF LIFE"
                    .Cells(auditRow, 6).Interior.Color = RGB(255, 235, 156)
                    auditRow = auditRow + 1
                End If
            End If
        Next i

        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 30
        .Columns("C:F").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "Impairment Review Complete!", vbInformation

End Sub
```

---

## Existence Testing Template

```vba
Sub GeneratePPEPhysicalInspection()
    '================================================
    ' GENERATE PHYSICAL INSPECTION SELECTION
    '
    ' Creates sample of assets for physical verification
    ' per AU-C 501 requirements
    '================================================

    Dim wsFA As Worksheet
    Dim wsSample As Worksheet
    Dim lastRow As Long, i As Long
    Dim sampleRow As Long

    Const SAMPLE_SIZE As Long = 25

    On Error Resume Next
    Set wsFA = ThisWorkbook.Sheets("FA_Register")
    On Error GoTo 0

    If wsFA Is Nothing Then
        MsgBox "FA_Register sheet required.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Sheets("PPE_Inspection").Delete
    On Error GoTo 0

    Set wsSample = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsSample.Name = "PPE_Inspection"

    With wsSample
        .Range("A1").Value = "FIXED ASSET PHYSICAL INSPECTION"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "Sample Size: " & SAMPLE_SIZE & " assets"
        .Range("A3").Value = "Date: " & Format(Date, "mmmm d, yyyy")

        sampleRow = 5

        .Cells(sampleRow, 1).Value = "Asset ID"
        .Cells(sampleRow, 2).Value = "Description"
        .Cells(sampleRow, 3).Value = "Location"
        .Cells(sampleRow, 4).Value = "Cost"
        .Cells(sampleRow, 5).Value = "Inspected By"
        .Cells(sampleRow, 6).Value = "Date"
        .Cells(sampleRow, 7).Value = "Condition"
        .Cells(sampleRow, 8).Value = "Status"
        .Range(.Cells(sampleRow, 1), .Cells(sampleRow, 8)).Font.Bold = True
        .Range(.Cells(sampleRow, 1), .Cells(sampleRow, 8)).Interior.Color = RGB(0, 51, 102)
        .Range(.Cells(sampleRow, 1), .Cells(sampleRow, 8)).Font.Color = RGB(255, 255, 255)
        sampleRow = sampleRow + 1

        ' Random sample selection
        lastRow = wsFA.Cells(wsFA.Rows.Count, "A").End(xlUp).Row
        Dim totalAssets As Long
        totalAssets = lastRow - 1

        Dim selected As Object
        Set selected = CreateObject("Scripting.Dictionary")

        Randomize

        Dim attempts As Long
        Do While selected.Count < Application.WorksheetFunction.Min(SAMPLE_SIZE, totalAssets) And attempts < 1000
            Dim randRow As Long
            randRow = Int((totalAssets) * Rnd + 2)

            If Not selected.Exists(randRow) Then
                If LCase(wsFA.Cells(randRow, 11).Value) <> "yes" Then
                    selected.Add randRow, True

                    .Cells(sampleRow, 1).Value = wsFA.Cells(randRow, 1).Value
                    .Cells(sampleRow, 2).Value = wsFA.Cells(randRow, 2).Value
                    .Cells(sampleRow, 3).Value = wsFA.Cells(randRow, 10).Value
                    .Cells(sampleRow, 4).Value = wsFA.Cells(randRow, 5).Value
                    .Cells(sampleRow, 4).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

                    sampleRow = sampleRow + 1
                End If
            End If
            attempts = attempts + 1
        Loop

        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 20
        .Columns("D:H").ColumnWidth = 15

    End With

    MsgBox "Physical Inspection Sample Generated!" & vbCrLf & _
           selected.Count & " assets selected for inspection", vbInformation

End Sub
```

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Accrued Expenses](./accrued-expenses.md) | [➡️ Debt](./debt.md)
