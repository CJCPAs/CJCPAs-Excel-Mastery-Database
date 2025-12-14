# Inventory Audit VBA

> **Inventory Audit Automation** - Complete VBA for auditing inventory per GAAS/GAAP (ASC 330)

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 1200-1299 (typically) |
| **Assertions** | Existence, Valuation, Completeness, Rights |
| **Risk Level** | HIGH (valuation, obsolescence, fraud risk) |
| **Key Standards** | ASC 330 (Inventory), Lower of Cost or NRV |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for inventory accounts

| Column | Header | Example |
|--------|--------|---------|
| A | `Date` | 12/31/2024 |
| B | `JE_Number` | JE-2024-3333 |
| C | `Account` | 1200 |
| D | `Account_Name` | Raw Materials |
| E | `Description` | Inventory adjustment |
| F | `Debit` | 25000 |
| G | `Credit` | 0 |

### Input Sheet 2: `Inventory_Listing`
Perpetual inventory listing at year-end

| Column | Header | Example |
|--------|--------|---------|
| A | `Item_Number` | SKU-001 |
| B | `Description` | Widget A |
| C | `Location` | Warehouse 1 |
| D | `Quantity` | 500 |
| E | `Unit_Cost` | 25.00 |
| F | `Extended_Cost` | 12500 |
| G | `Last_Sale_Date` | 11/15/2024 |
| H | `Last_Purchase_Date` | 12/01/2024 |

### Input Sheet 3: `Test_Counts`
Physical inventory test counts

| Column | Header | Example |
|--------|--------|---------|
| A | `Item_Number` | SKU-001 |
| B | `Description` | Widget A |
| C | `Book_Quantity` | 500 |
| D | `Count_Quantity` | 498 |
| E | `Unit_Cost` | 25.00 |

---

## Audit Procedures

### 1. Complete Inventory Audit Module

```vba
Sub AuditInventory()
    '================================================
    ' INVENTORY - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with inventory transactions
    '   - Sheet "Inventory_Listing" with perpetual inventory
    '   - Sheet "Test_Counts" with physical count results
    '
    ' OUTPUTS:
    '   - Creates "Inventory_Audit" worksheet
    '   - Reconciles GL to perpetual
    '   - Tests count accuracy
    '   - Analyzes obsolescence
    '   - Tests cost accuracy
    '
    ' ASSERTIONS TESTED:
    '   - Existence (test counts)
    '   - Valuation (cost testing, NRV, obsolescence)
    '   - Completeness (GL to perpetual tie)
    '================================================

    Dim wsGL As Worksheet
    Dim wsInv As Worksheet
    Dim wsCounts As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    ' Materiality thresholds
    Const MATERIALITY As Double = 50000
    Const TRIVIAL As Double = 2500
    Const OBSOLETE_DAYS As Long = 365  ' No sales in 1 year = obsolete
    Const COUNT_VARIANCE_PCT As Double = 0.02  ' 2% count variance threshold

    ' Validate required sheets
    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsInv = ThisWorkbook.Sheets("Inventory_Listing")
    Set wsCounts = ThisWorkbook.Sheets("Test_Counts")
    On Error GoTo 0

    If wsGL Is Nothing Or wsInv Is Nothing Then
        MsgBox "Required sheets not found." & vbCrLf & _
               "Please ensure GL_Detail and Inventory_Listing sheets exist.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Inventory_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Inventory_Audit"

    Application.ScreenUpdating = False

    ' ========================================
    ' HEADER
    ' ========================================
    With wsAudit
        .Range("A1").Value = "INVENTORY - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now
        .Range("A4").Value = "Materiality: " & Format(MATERIALITY, "$#,##0")

        auditRow = 6
    End With

    ' ========================================
    ' TEST 1: GL TO PERPETUAL RECONCILIATION
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 1: GL TO PERPETUAL INVENTORY RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        ' Calculate GL Balance
        Dim glBalance As Double
        glBalance = 0
        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 2) = "12" Then  ' Inventory accounts
                glBalance = glBalance + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
            End If
        Next i

        ' Calculate Perpetual Balance
        Dim perpBalance As Double
        perpBalance = 0
        lastRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            perpBalance = perpBalance + wsInv.Cells(i, 6).Value  ' Extended cost
        Next i

        Dim reconDiff As Double
        reconDiff = glBalance - perpBalance

        .Cells(auditRow, 1).Value = "GL Balance (1200-1299):"
        .Cells(auditRow, 2).Value = glBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Perpetual Inventory Balance:"
        .Cells(auditRow, 2).Value = perpBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "DIFFERENCE:"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = reconDiff

        If Abs(reconDiff) < 1 Then
            .Cells(auditRow, 3).Value = "RECONCILED"
            .Cells(auditRow, 3).Interior.Color = RGB(198, 239, 206)
        ElseIf Abs(reconDiff) < TRIVIAL Then
            .Cells(auditRow, 3).Value = "TRIVIAL DIFF"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 235, 156)
        Else
            .Cells(auditRow, 3).Value = "EXCEPTION - INVESTIGATE"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 199, 206)
        End If

        .Range(.Cells(auditRow - 2, 2), .Cells(auditRow, 2)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 2: INVENTORY TEST COUNTS
    ' ========================================
    If Not wsCounts Is Nothing Then
        With wsAudit
            .Cells(auditRow, 1).Value = "TEST 2: PHYSICAL COUNT TEST RESULTS"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Item #"
            .Cells(auditRow, 2).Value = "Description"
            .Cells(auditRow, 3).Value = "Book Qty"
            .Cells(auditRow, 4).Value = "Count Qty"
            .Cells(auditRow, 5).Value = "Variance"
            .Cells(auditRow, 6).Value = "Unit Cost"
            .Cells(auditRow, 7).Value = "$ Variance"
            .Cells(auditRow, 8).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Font.Bold = True
            auditRow = auditRow + 1

            Dim countStart As Long
            countStart = auditRow

            Dim totalVariance As Double
            Dim varianceCount As Long

            lastRow = wsCounts.Cells(wsCounts.Rows.Count, "A").End(xlUp).Row

            For i = 2 To lastRow
                Dim bookQty As Double
                Dim countQty As Double
                Dim qtyVar As Double
                Dim unitCost As Double
                Dim dollarVar As Double

                bookQty = wsCounts.Cells(i, 3).Value
                countQty = wsCounts.Cells(i, 4).Value
                unitCost = wsCounts.Cells(i, 5).Value
                qtyVar = countQty - bookQty
                dollarVar = qtyVar * unitCost

                .Cells(auditRow, 1).Value = wsCounts.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsCounts.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = bookQty
                .Cells(auditRow, 4).Value = countQty
                .Cells(auditRow, 5).Value = qtyVar
                .Cells(auditRow, 6).Value = unitCost
                .Cells(auditRow, 7).Value = dollarVar

                totalVariance = totalVariance + dollarVar

                ' Status based on variance
                If qtyVar = 0 Then
                    .Cells(auditRow, 8).Value = "EXACT MATCH"
                    .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
                ElseIf Abs(qtyVar) / bookQty <= COUNT_VARIANCE_PCT Then
                    .Cells(auditRow, 8).Value = "Within tolerance"
                    .Cells(auditRow, 8).Interior.Color = RGB(255, 235, 156)
                Else
                    .Cells(auditRow, 8).Value = "VARIANCE - INVESTIGATE"
                    .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
                    varianceCount = varianceCount + 1
                End If

                auditRow = auditRow + 1
            Next i

            ' Summary
            auditRow = auditRow + 1
            .Cells(auditRow, 1).Value = "COUNT SUMMARY:"
            .Cells(auditRow, 1).Font.Bold = True
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Items Tested:"
            .Cells(auditRow, 2).Value = lastRow - 1
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Items with Variance:"
            .Cells(auditRow, 2).Value = varianceCount
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Net $ Variance:"
            .Cells(auditRow, 2).Value = totalVariance
            .Cells(auditRow, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            If Abs(totalVariance) >= MATERIALITY Then
                .Cells(auditRow, 3).Value = "MATERIAL - PROPOSE ADJUSTMENT"
                .Cells(auditRow, 3).Interior.Color = RGB(255, 199, 206)
            End If

            .Range(.Cells(countStart, 6), .Cells(auditRow - 4, 7)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 3
        End With
    End If

    ' ========================================
    ' TEST 3: OBSOLESCENCE / SLOW-MOVING ANALYSIS
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 3: OBSOLESCENCE & SLOW-MOVING INVENTORY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Item #"
        .Cells(auditRow, 2).Value = "Description"
        .Cells(auditRow, 3).Value = "Extended Cost"
        .Cells(auditRow, 4).Value = "Last Sale Date"
        .Cells(auditRow, 5).Value = "Days Since Sale"
        .Cells(auditRow, 6).Value = "Obsolete Risk"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
        auditRow = auditRow + 1

        Dim obsStart As Long
        obsStart = auditRow

        Dim obsoleteTotal As Double
        Dim slowTotal As Double

        lastRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            Dim lastSaleDate As Date
            Dim daysSinceSale As Long
            Dim extCost As Double

            extCost = wsInv.Cells(i, 6).Value

            If IsDate(wsInv.Cells(i, 7).Value) Then
                lastSaleDate = wsInv.Cells(i, 7).Value
                daysSinceSale = Date - lastSaleDate

                ' Flag items with no sales in specified period
                If daysSinceSale > OBSOLETE_DAYS / 2 Then  ' Over 6 months
                    .Cells(auditRow, 1).Value = wsInv.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsInv.Cells(i, 2).Value
                    .Cells(auditRow, 3).Value = extCost
                    .Cells(auditRow, 4).Value = lastSaleDate
                    .Cells(auditRow, 5).Value = daysSinceSale

                    If daysSinceSale > OBSOLETE_DAYS Then
                        .Cells(auditRow, 6).Value = "OBSOLETE"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                        obsoleteTotal = obsoleteTotal + extCost
                    Else
                        .Cells(auditRow, 6).Value = "Slow-Moving"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 235, 156)
                        slowTotal = slowTotal + extCost
                    End If

                    auditRow = auditRow + 1
                End If
            End If
        Next i

        ' Summary
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "OBSOLESCENCE SUMMARY:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Obsolete Inventory (>" & OBSOLETE_DAYS & " days):"
        .Cells(auditRow, 2).Value = obsoleteTotal
        .Cells(auditRow, 3).Value = obsoleteTotal / perpBalance
        .Cells(auditRow, 3).NumberFormat = "0.0%"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Slow-Moving Inventory:"
        .Cells(auditRow, 2).Value = slowTotal
        .Cells(auditRow, 3).Value = slowTotal / perpBalance
        .Cells(auditRow, 3).NumberFormat = "0.0%"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Recommended Reserve:"
        .Cells(auditRow, 2).Value = obsoleteTotal + (slowTotal * 0.25)  ' 100% obsolete + 25% slow
        .Cells(auditRow, 2).Font.Bold = True

        .Range(.Cells(obsStart, 3), .Cells(auditRow, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 4: INVENTORY BY CATEGORY
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 4: INVENTORY COMPOSITION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Location"
        .Cells(auditRow, 2).Value = "Value"
        .Cells(auditRow, 3).Value = "% of Total"
        .Cells(auditRow, 4).Value = "Item Count"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        Dim locStart As Long
        locStart = auditRow

        ' Aggregate by location
        Dim locDict As Object
        Set locDict = CreateObject("Scripting.Dictionary")

        lastRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            Dim loc As String
            Dim locAmt As Double

            loc = wsInv.Cells(i, 3).Value
            If loc = "" Then loc = "Unspecified"
            locAmt = wsInv.Cells(i, 6).Value

            If locDict.Exists(loc) Then
                locDict(loc) = Array(locDict(loc)(0) + locAmt, locDict(loc)(1) + 1)
            Else
                locDict.Add loc, Array(locAmt, 1)
            End If
        Next i

        Dim key As Variant
        Dim locData As Variant

        For Each key In locDict.Keys
            locData = locDict(key)
            .Cells(auditRow, 1).Value = key
            .Cells(auditRow, 2).Value = locData(0)
            .Cells(auditRow, 3).Value = locData(0) / perpBalance
            .Cells(auditRow, 4).Value = locData(1)
            auditRow = auditRow + 1
        Next key

        .Range(.Cells(locStart, 2), .Cells(auditRow - 1, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(locStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "0.0%"

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

        .Cells(auditRow, 1).Value = "Total Inventory Balance:"
        .Cells(auditRow, 2).Value = perpBalance
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " GL to perpetual reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Physical count testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Obsolescence analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Location analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Cost testing (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " NRV testing (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Cutoff testing (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion based on procedures performed]"
        .Cells(auditRow, 1).Font.Italic = True

        ' Column widths
        .Columns("A").ColumnWidth = 30
        .Columns("B:H").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "Inventory Audit Complete!" & vbCrLf & vbCrLf & _
           "Total Inventory: " & Format(perpBalance, "$#,##0") & vbCrLf & _
           "Review the Inventory_Audit worksheet.", vbInformation

End Sub
```

---

## Lower of Cost or NRV Testing

```vba
Sub TestInventoryNRV()
    '================================================
    ' Test Inventory at Lower of Cost or NRV
    ' Per ASC 330-10-35
    '================================================

    Dim wsInv As Worksheet
    Dim wsNRV As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nrvRow As Long
    Dim writeDownTotal As Double

    On Error Resume Next
    Set wsInv = ThisWorkbook.Sheets("Inventory_Listing")
    On Error GoTo 0

    If wsInv Is Nothing Then
        MsgBox "Inventory_Listing sheet required.", vbExclamation
        Exit Sub
    End If

    ' Create NRV testing worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("NRV_Testing").Delete
    On Error GoTo 0

    Set wsNRV = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsNRV.Name = "NRV_Testing"

    Application.ScreenUpdating = False

    With wsNRV
        .Range("A1").Value = "LOWER OF COST OR NET REALIZABLE VALUE TEST"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Per ASC 330-10-35: Inventory measured at lower of cost or NRV"

        .Range("A4").Value = "Item #"
        .Range("B4").Value = "Description"
        .Range("C4").Value = "Quantity"
        .Range("D4").Value = "Unit Cost"
        .Range("E4").Value = "Selling Price"
        .Range("F4").Value = "Est. Costs to Sell"
        .Range("G4").Value = "NRV"
        .Range("H4").Value = "Lower of C/NRV"
        .Range("I4").Value = "Write-Down"
        .Range("J4").Value = "Status"
        .Range("A4:J4").Font.Bold = True
        .Range("A4:J4").Interior.Color = RGB(0, 51, 102)
        .Range("A4:J4").Font.Color = RGB(255, 255, 255)

        nrvRow = 5

        ' Prompt for selling prices (in real audit, would come from pricing data)
        MsgBox "For each item, enter the Selling Price in column E and Est. Costs to Sell in column F." & vbCrLf & _
               "The macro will calculate NRV and any required write-downs.", vbInformation

        lastRow = wsInv.Cells(wsInv.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            .Cells(nrvRow, 1).Value = wsInv.Cells(i, 1).Value
            .Cells(nrvRow, 2).Value = wsInv.Cells(i, 2).Value
            .Cells(nrvRow, 3).Value = wsInv.Cells(i, 4).Value  ' Quantity
            .Cells(nrvRow, 4).Value = wsInv.Cells(i, 5).Value  ' Unit cost

            ' Columns E and F are for user input
            .Cells(nrvRow, 5).Interior.Color = RGB(255, 255, 200)  ' Yellow for input
            .Cells(nrvRow, 6).Interior.Color = RGB(255, 255, 200)

            ' NRV = Selling Price - Costs to Sell
            .Cells(nrvRow, 7).Formula = "=E" & nrvRow & "-F" & nrvRow

            ' Lower of Cost or NRV
            .Cells(nrvRow, 8).Formula = "=MIN(D" & nrvRow & ",G" & nrvRow & ")"

            ' Write-down needed
            .Cells(nrvRow, 9).Formula = "=IF(G" & nrvRow & "<D" & nrvRow & ",(D" & nrvRow & "-G" & nrvRow & ")*C" & nrvRow & ",0)"

            ' Status
            .Cells(nrvRow, 10).Formula = "=IF(G" & nrvRow & "<D" & nrvRow & ",""WRITE-DOWN NEEDED"",""OK"")"

            nrvRow = nrvRow + 1
        Next i

        ' Total write-down
        nrvRow = nrvRow + 1
        .Cells(nrvRow, 1).Value = "TOTAL WRITE-DOWN REQUIRED:"
        .Cells(nrvRow, 1).Font.Bold = True
        .Cells(nrvRow, 9).Formula = "=SUM(I5:I" & (nrvRow - 2) & ")"
        .Cells(nrvRow, 9).Font.Bold = True

        ' Conditional formatting
        .Range("J5:J" & nrvRow - 2).FormatConditions.Add Type:=xlTextString, String:="WRITE-DOWN", TextOperator:=xlContains
        .Range("J5:J" & nrvRow - 2).FormatConditions(1).Interior.Color = RGB(255, 199, 206)

        ' Format
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 25
        .Columns("C:J").ColumnWidth = 15
        .Range("D:I").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    Application.ScreenUpdating = True

    MsgBox "NRV testing template created!" & vbCrLf & _
           "Enter selling prices and costs to sell for each item.", vbInformation

End Sub
```

---

## Assertions Tested

| Assertion | Test | Pass Criteria |
|-----------|------|---------------|
| **Existence** | Physical counts | Count = Book |
| **Valuation** | Cost testing, NRV, obsolescence | Lower of cost/NRV |
| **Completeness** | GL to perpetual tie | Balances agree |
| **Rights** | Title review | Company owns inventory |

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ AR](./accounts-receivable.md) | [➡️ PP&E](./ppe.md)
