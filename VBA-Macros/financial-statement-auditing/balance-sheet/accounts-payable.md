# Accounts Payable Audit VBA

> **AP Audit - Search for Unrecorded Liabilities** - Complete VBA for auditing payables per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 2000-2099 (typically) |
| **Assertions** | **Completeness** (primary), Existence, Accuracy, Cutoff |
| **Risk Level** | HIGH (understatement risk) |
| **Key Documents** | AP aging, vendor invoices, subsequent disbursements, statements |

**Key Audit Focus:** Unlike AR, AP risk is primarily **understatement** (completeness) - companies may not record all liabilities.

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for AP accounts

| Column | Header | Example |
|--------|--------|---------|
| A | `Date` | 12/31/2024 |
| B | `JE_Number` | JE-2024-5678 |
| C | `Account` | 2000 |
| D | `Account_Name` | Accounts Payable |
| E | `Description` | Vendor invoice |
| F | `Debit` | 0 |
| G | `Credit` | 15000 |
| H | `Source` | AP |

### Input Sheet 2: `AP_Aging`
Accounts payable aging report

| Column | Header | Example |
|--------|--------|---------|
| A | `Vendor_ID` | VEND001 |
| B | `Vendor_Name` | ABC Supplies |
| C | `Invoice_Number` | INV-9876 |
| D | `Invoice_Date` | 12/15/2024 |
| E | `Due_Date` | 01/15/2025 |
| F | `Amount` | 15000 |
| G | `Current` | 15000 |
| H | `1_30_Days` | 0 |
| I | `31_60_Days` | 0 |
| J | `Over_60_Days` | 0 |

### Input Sheet 3: `Subsequent_Disbursements`
Cash disbursements after year-end

| Column | Header | Example |
|--------|--------|---------|
| A | `Check_Date` | 01/10/2025 |
| B | `Check_Number` | 5001 |
| C | `Vendor_ID` | VEND001 |
| D | `Vendor_Name` | ABC Supplies |
| E | `Invoice_Number` | INV-9876 |
| F | `Invoice_Date` | 12/15/2024 |
| G | `Amount` | 15000 |

---

## Audit Procedures

### 1. Complete AP Audit Module

```vba
Sub AuditAccountsPayable()
    '================================================
    ' ACCOUNTS PAYABLE - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with AP transactions
    '   - Sheet "AP_Aging" with aged payables
    '   - Sheet "Subsequent_Disbursements" with post Y/E payments
    '
    ' OUTPUTS:
    '   - Creates "AP_Audit" worksheet with all test results
    '   - Performs search for unrecorded liabilities
    '   - Tests cutoff
    '   - Analyzes vendor concentrations
    '
    ' ASSERTIONS TESTED:
    '   - COMPLETENESS (primary - unrecorded liabilities)
    '   - Existence (vendor confirmation)
    '   - Accuracy (invoice agreement)
    '   - Cutoff (goods/services received)
    '================================================

    Dim wsGL As Worksheet
    Dim wsAP As Worksheet
    Dim wsDisb As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    ' Materiality thresholds
    Const MATERIALITY As Double = 50000
    Const TRIVIAL As Double = 2500
    Const SEARCH_THRESHOLD As Double = 5000  ' Search for items over this

    ' Validate required sheets
    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsAP = ThisWorkbook.Sheets("AP_Aging")
    Set wsDisb = ThisWorkbook.Sheets("Subsequent_Disbursements")
    On Error GoTo 0

    If wsGL Is Nothing Or wsAP Is Nothing Then
        MsgBox "Required sheets not found." & vbCrLf & _
               "Please ensure GL_Detail and AP_Aging sheets exist.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("AP_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "AP_Audit"

    Application.ScreenUpdating = False

    ' ========================================
    ' HEADER
    ' ========================================
    With wsAudit
        .Range("A1").Value = "ACCOUNTS PAYABLE - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now
        .Range("A4").Value = "Materiality: " & Format(MATERIALITY, "$#,##0")

        auditRow = 6
    End With

    ' ========================================
    ' TEST 1: GL TO SUB-LEDGER RECONCILIATION
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 1: GL TO SUB-LEDGER RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        ' Calculate GL Balance (credit balance = positive)
        Dim glBalance As Double
        glBalance = 0
        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 2) = "20" Then  ' AP accounts
                glBalance = glBalance + wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value
            End If
        Next i

        ' Calculate Sub-ledger Balance
        Dim subBalance As Double
        subBalance = 0
        lastRow = wsAP.Cells(wsAP.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            subBalance = subBalance + wsAP.Cells(i, 6).Value
        Next i

        Dim reconDiff As Double
        reconDiff = glBalance - subBalance

        .Cells(auditRow, 1).Value = "GL Balance (2000-2099):"
        .Cells(auditRow, 2).Value = glBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Sub-Ledger Balance:"
        .Cells(auditRow, 2).Value = subBalance
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
            .Cells(auditRow, 3).Value = "EXCEPTION"
            .Cells(auditRow, 3).Interior.Color = RGB(255, 199, 206)
        End If

        .Range(.Cells(auditRow - 2, 2), .Cells(auditRow, 2)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 2: SEARCH FOR UNRECORDED LIABILITIES
    ' ========================================
    If Not wsDisb Is Nothing Then
        With wsAudit
            .Cells(auditRow, 1).Value = "TEST 2: SEARCH FOR UNRECORDED LIABILITIES"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Examining subsequent disbursements for items that should have been accrued at Y/E"
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Check Date"
            .Cells(auditRow, 2).Value = "Check #"
            .Cells(auditRow, 3).Value = "Vendor"
            .Cells(auditRow, 4).Value = "Invoice #"
            .Cells(auditRow, 5).Value = "Invoice Date"
            .Cells(auditRow, 6).Value = "Amount"
            .Cells(auditRow, 7).Value = "In Y/E AP?"
            .Cells(auditRow, 8).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Font.Bold = True
            auditRow = auditRow + 1

            Dim searchStart As Long
            searchStart = auditRow

            Dim unrecordedCount As Long
            Dim unrecordedAmount As Double

            lastRow = wsDisb.Cells(wsDisb.Rows.Count, "A").End(xlUp).Row
            Dim apLastRow As Long
            apLastRow = wsAP.Cells(wsAP.Rows.Count, "A").End(xlUp).Row

            For i = 2 To lastRow
                Dim disbAmt As Double
                Dim invDate As Date
                Dim invNum As String
                Dim vendorName As String
                Dim foundInAP As Boolean
                Dim yearEnd As Date

                yearEnd = DateSerial(Year(Date), 12, 31)

                disbAmt = wsDisb.Cells(i, 7).Value

                ' Only test items over search threshold
                If disbAmt >= SEARCH_THRESHOLD Then
                    invNum = wsDisb.Cells(i, 5).Value
                    vendorName = wsDisb.Cells(i, 4).Value

                    If IsDate(wsDisb.Cells(i, 6).Value) Then
                        invDate = wsDisb.Cells(i, 6).Value
                    Else
                        invDate = Date  ' Default to current if not a date
                    End If

                    ' Check if invoice was in Y/E AP
                    foundInAP = False
                    Dim j As Long
                    For j = 2 To apLastRow
                        If wsAP.Cells(j, 3).Value = invNum Then
                            foundInAP = True
                            Exit For
                        End If
                    Next j

                    .Cells(auditRow, 1).Value = wsDisb.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsDisb.Cells(i, 2).Value
                    .Cells(auditRow, 3).Value = vendorName
                    .Cells(auditRow, 4).Value = invNum
                    .Cells(auditRow, 5).Value = invDate
                    .Cells(auditRow, 6).Value = disbAmt

                    If foundInAP Then
                        .Cells(auditRow, 7).Value = "Yes"
                        .Cells(auditRow, 8).Value = "Properly recorded"
                        .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
                    Else
                        ' Invoice dated before Y/E but not in AP?
                        If invDate <= yearEnd Then
                            .Cells(auditRow, 7).Value = "NO"
                            .Cells(auditRow, 8).Value = "UNRECORDED LIABILITY"
                            .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
                            unrecordedCount = unrecordedCount + 1
                            unrecordedAmount = unrecordedAmount + disbAmt
                        Else
                            .Cells(auditRow, 7).Value = "N/A"
                            .Cells(auditRow, 8).Value = "2025 invoice - OK"
                            .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
                        End If
                    End If

                    auditRow = auditRow + 1
                End If
            Next i

            ' Format
            .Range(.Cells(searchStart, 6), .Cells(auditRow - 1, 6)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Range(.Cells(searchStart, 1), .Cells(auditRow - 1, 1)).NumberFormat = "mm/dd/yyyy"
            .Range(.Cells(searchStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "mm/dd/yyyy"

            ' Summary
            auditRow = auditRow + 1
            .Cells(auditRow, 1).Value = "UNRECORDED LIABILITIES IDENTIFIED:"
            .Cells(auditRow, 1).Font.Bold = True
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Count:"
            .Cells(auditRow, 2).Value = unrecordedCount
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Total Amount:"
            .Cells(auditRow, 2).Value = unrecordedAmount
            .Cells(auditRow, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            If unrecordedAmount >= MATERIALITY Then
                .Cells(auditRow, 3).Value = "MATERIAL - PROPOSE ADJUSTMENT"
                .Cells(auditRow, 3).Interior.Color = RGB(255, 199, 206)
            ElseIf unrecordedAmount >= TRIVIAL Then
                .Cells(auditRow, 3).Value = "Above trivial - document"
                .Cells(auditRow, 3).Interior.Color = RGB(255, 235, 156)
            End If

            auditRow = auditRow + 3
        End With
    End If

    ' ========================================
    ' TEST 3: VENDOR CONCENTRATION ANALYSIS
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 3: VENDOR CONCENTRATION ANALYSIS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Vendor"
        .Cells(auditRow, 2).Value = "Balance"
        .Cells(auditRow, 3).Value = "% of Total AP"
        .Cells(auditRow, 4).Value = "Risk Level"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        Dim concStart As Long
        concStart = auditRow

        ' Aggregate by vendor
        Dim vendDict As Object
        Set vendDict = CreateObject("Scripting.Dictionary")

        lastRow = wsAP.Cells(wsAP.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            Dim vendID As String
            Dim vendName As String
            Dim vendAmt As Double

            vendID = wsAP.Cells(i, 1).Value
            vendName = wsAP.Cells(i, 2).Value
            vendAmt = wsAP.Cells(i, 6).Value

            If vendDict.Exists(vendID) Then
                vendDict(vendID) = Array(vendName, vendDict(vendID)(1) + vendAmt)
            Else
                vendDict.Add vendID, Array(vendName, vendAmt)
            End If
        Next i

        ' Output top vendors
        Dim key As Variant
        Dim vendData As Variant

        For Each key In vendDict.Keys
            vendData = vendDict(key)

            If vendData(1) >= SEARCH_THRESHOLD Then
                .Cells(auditRow, 1).Value = vendData(0)
                .Cells(auditRow, 2).Value = vendData(1)
                .Cells(auditRow, 3).Value = vendData(1) / subBalance

                If vendData(1) / subBalance > 0.25 Then
                    .Cells(auditRow, 4).Value = "HIGH CONCENTRATION"
                    .Cells(auditRow, 4).Interior.Color = RGB(255, 199, 206)
                ElseIf vendData(1) / subBalance > 0.1 Then
                    .Cells(auditRow, 4).Value = "Moderate"
                    .Cells(auditRow, 4).Interior.Color = RGB(255, 235, 156)
                Else
                    .Cells(auditRow, 4).Value = "Normal"
                End If

                auditRow = auditRow + 1
            End If
        Next key

        .Range(.Cells(concStart, 2), .Cells(auditRow - 1, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(concStart, 3), .Cells(auditRow - 1, 3)).NumberFormat = "0.0%"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 4: AP AGING ANALYSIS
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 4: AGING ANALYSIS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        ' Sum aging buckets
        Dim bucketCurrent As Double, bucket30 As Double
        Dim bucket60 As Double, bucketOver60 As Double

        lastRow = wsAP.Cells(wsAP.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            bucketCurrent = bucketCurrent + Val(wsAP.Cells(i, 7).Value)
            bucket30 = bucket30 + Val(wsAP.Cells(i, 8).Value)
            bucket60 = bucket60 + Val(wsAP.Cells(i, 9).Value)
            bucketOver60 = bucketOver60 + Val(wsAP.Cells(i, 10).Value)
        Next i

        .Cells(auditRow, 1).Value = "Current:"
        .Cells(auditRow, 2).Value = bucketCurrent
        .Cells(auditRow, 3).Value = bucketCurrent / subBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "1-30 Days:"
        .Cells(auditRow, 2).Value = bucket30
        .Cells(auditRow, 3).Value = bucket30 / subBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "31-60 Days:"
        .Cells(auditRow, 2).Value = bucket60
        .Cells(auditRow, 3).Value = bucket60 / subBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Over 60 Days:"
        .Cells(auditRow, 2).Value = bucketOver60
        .Cells(auditRow, 3).Value = bucketOver60 / subBalance

        If bucketOver60 / subBalance > 0.1 Then
            .Cells(auditRow, 4).Value = "HIGH - INVESTIGATE"
            .Cells(auditRow, 4).Interior.Color = RGB(255, 199, 206)
        End If

        .Range(.Cells(auditRow - 3, 2), .Cells(auditRow, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(auditRow - 3, 3), .Cells(auditRow, 3)).NumberFormat = "0.0%"

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

        .Cells(auditRow, 1).Value = "Total AP Balance:"
        .Cells(auditRow, 2).Value = subBalance
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " GL to sub-ledger reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Search for unrecorded liabilities"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Vendor concentration analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Aging analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Vendor confirmations (manual)"
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

    MsgBox "Accounts Payable Audit Complete!" & vbCrLf & vbCrLf & _
           "Total AP: " & Format(subBalance, "$#,##0") & vbCrLf & _
           "Review the AP_Audit worksheet.", vbInformation

End Sub
```

---

### 2. Purchase Cutoff Testing

```vba
Sub TestPurchaseCutoff()
    '================================================
    ' Purchase Cutoff Testing
    ' Tests purchases/receipts around year-end
    '================================================

    Dim wsGL As Worksheet
    Dim wsCutoff As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cutoffRow As Long

    Const CUTOFF_DAYS As Long = 5

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbExclamation
        Exit Sub
    End If

    ' Create cutoff worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Purchase_Cutoff").Delete
    On Error GoTo 0

    Set wsCutoff = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsCutoff.Name = "Purchase_Cutoff"

    Application.ScreenUpdating = False

    With wsCutoff
        .Range("A1").Value = "PURCHASE CUTOFF TESTING"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Testing transactions within " & CUTOFF_DAYS & " days of year-end"

        .Range("A4").Value = "Date"
        .Range("B4").Value = "JE Number"
        .Range("C4").Value = "Account"
        .Range("D4").Value = "Description"
        .Range("E4").Value = "Amount"
        .Range("F4").Value = "Days from Y/E"
        .Range("G4").Value = "Receipt Date"
        .Range("H4").Value = "Cutoff Status"
        .Range("A4:H4").Font.Bold = True
        .Range("A4:H4").Interior.Color = RGB(0, 51, 102)
        .Range("A4:H4").Font.Color = RGB(255, 255, 255)

        cutoffRow = 5

        Dim yearEnd As Date
        yearEnd = DateSerial(Year(Date), 12, 31)

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            ' Filter for AP accounts (2000s) and expense accounts (5000-7000s)
            If Left(wsGL.Cells(i, 3).Value, 1) = "2" Or _
               (Val(Left(wsGL.Cells(i, 3).Value, 1)) >= 5 And Val(Left(wsGL.Cells(i, 3).Value, 1)) <= 7) Then

                If IsDate(wsGL.Cells(i, 1).Value) Then
                    Dim transDate As Date
                    Dim daysFromYE As Long

                    transDate = wsGL.Cells(i, 1).Value
                    daysFromYE = transDate - yearEnd

                    If Abs(daysFromYE) <= CUTOFF_DAYS Then
                        .Cells(cutoffRow, 1).Value = transDate
                        .Cells(cutoffRow, 2).Value = wsGL.Cells(i, 2).Value
                        .Cells(cutoffRow, 3).Value = wsGL.Cells(i, 3).Value
                        .Cells(cutoffRow, 4).Value = wsGL.Cells(i, 5).Value
                        .Cells(cutoffRow, 5).Value = wsGL.Cells(i, 7).Value - wsGL.Cells(i, 6).Value
                        .Cells(cutoffRow, 6).Value = daysFromYE

                        If daysFromYE > 0 Then
                            .Cells(cutoffRow, 8).Value = "AFTER Y/E - VERIFY RECEIPT"
                            .Cells(cutoffRow, 8).Interior.Color = RGB(255, 199, 206)
                        Else
                            .Cells(cutoffRow, 8).Value = "Before Y/E"
                            .Cells(cutoffRow, 8).Interior.Color = RGB(198, 239, 206)
                        End If

                        cutoffRow = cutoffRow + 1
                    End If
                End If
            End If
        Next i

        ' Format
        .Columns("A").ColumnWidth = 12
        .Columns("B:H").ColumnWidth = 15
        .Range("E:E").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    Application.ScreenUpdating = True

    MsgBox "Purchase cutoff testing complete!", vbInformation

End Sub
```

---

## Assertions Tested

| Assertion | Test | Pass Criteria |
|-----------|------|---------------|
| **Completeness** | Search for unrecorded liabilities | All period invoices recorded |
| **Existence** | Vendor confirmations | Balance confirmed |
| **Accuracy** | Invoice agreement | Amounts match support |
| **Cutoff** | Goods receipt timing | Proper period recorded |

---

## Common Exceptions

| Exception | Cause | Resolution |
|-----------|-------|------------|
| Unrecorded liability | Invoice not processed | Propose adjustment |
| Vendor concentration | Single source risk | Disclose if significant |
| Old payables | Disputes, errors | Investigate, write off |
| Cutoff error | Wrong period | Reclassify |

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Inventory](./inventory.md) | [➡️ Accrued Expenses](./accrued-expenses.md)
