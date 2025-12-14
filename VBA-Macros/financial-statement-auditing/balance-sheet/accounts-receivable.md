# Accounts Receivable Audit VBA

> **AR Audit Automation** - Complete VBA procedures for auditing receivables per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 1100-1199 (typically) |
| **Assertions** | Existence, Valuation, Completeness, Rights, Cutoff |
| **Risk Level** | HIGH (revenue fraud risk, collectibility) |
| **Key Documents** | AR aging, customer invoices, confirmations, subsequent receipts |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for AR accounts

| Column | Header | Example |
|--------|--------|---------|
| A | `Date` | 12/31/2024 |
| B | `JE_Number` | JE-2024-1234 |
| C | `Account` | 1100 |
| D | `Account_Name` | Accounts Receivable |
| E | `Description` | Invoice #5001 |
| F | `Debit` | 10000 |
| G | `Credit` | 0 |
| H | `Source` | AR |

### Input Sheet 2: `AR_Aging`
Accounts receivable aging report

| Column | Header | Example |
|--------|--------|---------|
| A | `Customer_ID` | CUST001 |
| B | `Customer_Name` | Acme Corp |
| C | `Invoice_Number` | INV-5001 |
| D | `Invoice_Date` | 11/15/2024 |
| E | `Due_Date` | 12/15/2024 |
| F | `Amount` | 10000 |
| G | `Current` | 0 |
| H | `1_30_Days` | 0 |
| I | `31_60_Days` | 10000 |
| J | `61_90_Days` | 0 |
| K | `Over_90_Days` | 0 |

### Input Sheet 3: `Subsequent_Receipts`
Cash receipts after year-end

| Column | Header | Example |
|--------|--------|---------|
| A | `Receipt_Date` | 01/15/2025 |
| B | `Customer_ID` | CUST001 |
| C | `Customer_Name` | Acme Corp |
| D | `Invoice_Number` | INV-5001 |
| E | `Amount` | 10000 |

---

## Audit Procedures

### 1. Complete AR Audit Module

```vba
Sub AuditAccountsReceivable()
    '================================================
    ' ACCOUNTS RECEIVABLE - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with AR transactions
    '   - Sheet "AR_Aging" with aged receivables
    '   - Sheet "Subsequent_Receipts" with post Y/E collections
    '
    ' OUTPUTS:
    '   - Creates "AR_Audit" worksheet with all test results
    '   - Generates confirmation selection
    '   - Analyzes allowance adequacy
    '   - Tests subsequent receipts
    '
    ' ASSERTIONS TESTED:
    '   - Existence (confirmations, subsequent receipts)
    '   - Valuation (allowance adequacy, aging analysis)
    '   - Completeness (GL to sub-ledger tie)
    '   - Cutoff (sales around year-end)
    '================================================

    Dim wsGL As Worksheet
    Dim wsAR As Worksheet
    Dim wsReceipts As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    ' Materiality and thresholds
    Const MATERIALITY As Double = 50000
    Const TRIVIAL As Double = 2500
    Const CONFIRM_THRESHOLD As Double = 10000  ' Confirm balances over this
    Const ALLOWANCE_RATE_CURRENT As Double = 0.01      ' 1% current
    Const ALLOWANCE_RATE_30 As Double = 0.05           ' 5% 1-30 days
    Const ALLOWANCE_RATE_60 As Double = 0.1            ' 10% 31-60 days
    Const ALLOWANCE_RATE_90 As Double = 0.25           ' 25% 61-90 days
    Const ALLOWANCE_RATE_OVER90 As Double = 0.5        ' 50% over 90 days

    ' Validate required sheets
    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsAR = ThisWorkbook.Sheets("AR_Aging")
    Set wsReceipts = ThisWorkbook.Sheets("Subsequent_Receipts")
    On Error GoTo 0

    If wsGL Is Nothing Or wsAR Is Nothing Then
        MsgBox "Required sheets not found." & vbCrLf & _
               "Please ensure GL_Detail and AR_Aging sheets exist.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("AR_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "AR_Audit"

    Application.ScreenUpdating = False

    ' ========================================
    ' HEADER
    ' ========================================
    With wsAudit
        .Range("A1").Value = "ACCOUNTS RECEIVABLE - AUDIT WORKPAPER"
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

        ' Calculate GL Balance
        Dim glBalance As Double
        glBalance = 0
        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 2) = "11" Then  ' AR accounts
                glBalance = glBalance + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
            End If
        Next i

        ' Calculate Sub-ledger Balance
        Dim subBalance As Double
        subBalance = 0
        lastRow = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            subBalance = subBalance + wsAR.Cells(i, 6).Value  ' Amount column
        Next i

        Dim reconDiff As Double
        reconDiff = glBalance - subBalance

        .Cells(auditRow, 1).Value = "GL Balance (1100-1199):"
        .Cells(auditRow, 2).Value = glBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Sub-Ledger Balance:"
        .Cells(auditRow, 2).Value = subBalance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "DIFFERENCE:"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = reconDiff
        .Cells(auditRow, 2).Font.Bold = True

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
    ' TEST 2: AGING ANALYSIS & ALLOWANCE TEST
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 2: AGING ANALYSIS & ALLOWANCE ADEQUACY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
        auditRow = auditRow + 2

        ' Headers
        .Cells(auditRow, 1).Value = "Aging Bucket"
        .Cells(auditRow, 2).Value = "Balance"
        .Cells(auditRow, 3).Value = "% of Total"
        .Cells(auditRow, 4).Value = "Reserve Rate"
        .Cells(auditRow, 5).Value = "Calculated Reserve"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim agingStart As Long
        agingStart = auditRow

        ' Sum aging buckets
        Dim bucketCurrent As Double, bucket30 As Double, bucket60 As Double
        Dim bucket90 As Double, bucketOver90 As Double
        Dim totalAR As Double

        lastRow = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            bucketCurrent = bucketCurrent + Val(wsAR.Cells(i, 7).Value)
            bucket30 = bucket30 + Val(wsAR.Cells(i, 8).Value)
            bucket60 = bucket60 + Val(wsAR.Cells(i, 9).Value)
            bucket90 = bucket90 + Val(wsAR.Cells(i, 10).Value)
            bucketOver90 = bucketOver90 + Val(wsAR.Cells(i, 11).Value)
        Next i

        totalAR = bucketCurrent + bucket30 + bucket60 + bucket90 + bucketOver90

        ' Output aging
        .Cells(auditRow, 1).Value = "Current"
        .Cells(auditRow, 2).Value = bucketCurrent
        .Cells(auditRow, 3).Value = bucketCurrent / totalAR
        .Cells(auditRow, 4).Value = ALLOWANCE_RATE_CURRENT
        .Cells(auditRow, 5).Value = bucketCurrent * ALLOWANCE_RATE_CURRENT
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "1-30 Days"
        .Cells(auditRow, 2).Value = bucket30
        .Cells(auditRow, 3).Value = bucket30 / totalAR
        .Cells(auditRow, 4).Value = ALLOWANCE_RATE_30
        .Cells(auditRow, 5).Value = bucket30 * ALLOWANCE_RATE_30
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "31-60 Days"
        .Cells(auditRow, 2).Value = bucket60
        .Cells(auditRow, 3).Value = bucket60 / totalAR
        .Cells(auditRow, 4).Value = ALLOWANCE_RATE_60
        .Cells(auditRow, 5).Value = bucket60 * ALLOWANCE_RATE_60
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "61-90 Days"
        .Cells(auditRow, 2).Value = bucket90
        .Cells(auditRow, 3).Value = bucket90 / totalAR
        .Cells(auditRow, 4).Value = ALLOWANCE_RATE_90
        .Cells(auditRow, 5).Value = bucket90 * ALLOWANCE_RATE_90
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Over 90 Days"
        .Cells(auditRow, 2).Value = bucketOver90
        .Cells(auditRow, 3).Value = bucketOver90 / totalAR
        .Cells(auditRow, 4).Value = ALLOWANCE_RATE_OVER90
        .Cells(auditRow, 5).Value = bucketOver90 * ALLOWANCE_RATE_OVER90

        If bucketOver90 > totalAR * 0.1 Then
            .Cells(auditRow, 6).Value = "HIGH RISK"
            .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Totals
        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = totalAR
        .Cells(auditRow, 2).Font.Bold = True
        .Cells(auditRow, 5).Formula = "=SUM(E" & agingStart & ":E" & (auditRow - 1) & ")"
        .Cells(auditRow, 5).Font.Bold = True

        Dim calculatedAllowance As Double
        calculatedAllowance = .Cells(auditRow, 5).Value

        ' Format
        .Range(.Cells(agingStart, 2), .Cells(auditRow, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(agingStart, 3), .Cells(auditRow - 1, 4)).NumberFormat = "0.0%"
        .Range(.Cells(agingStart, 5), .Cells(auditRow, 5)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' Compare to recorded allowance
        .Cells(auditRow, 1).Value = "ALLOWANCE COMPARISON:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Calculated Allowance:"
        .Cells(auditRow, 2).Value = calculatedAllowance
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Recorded Allowance (enter):"
        .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 200)  ' Yellow for input
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Difference:"
        .Cells(auditRow, 2).Formula = "=B" & (auditRow - 1) & "-B" & (auditRow - 2)
        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 3: CONFIRMATION SELECTION
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 3: CONFIRMATION SELECTION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Customer"
        .Cells(auditRow, 2).Value = "Balance"
        .Cells(auditRow, 3).Value = "% of Total AR"
        .Cells(auditRow, 4).Value = "Confirm?"
        .Cells(auditRow, 5).Value = "Sent Date"
        .Cells(auditRow, 6).Value = "Response"
        .Cells(auditRow, 7).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
        auditRow = auditRow + 1

        Dim confStart As Long
        confStart = auditRow

        ' Get unique customers with balances over threshold
        Dim custDict As Object
        Set custDict = CreateObject("Scripting.Dictionary")

        lastRow = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            Dim custID As String
            Dim custName As String
            Dim invAmt As Double

            custID = wsAR.Cells(i, 1).Value
            custName = wsAR.Cells(i, 2).Value
            invAmt = wsAR.Cells(i, 6).Value

            If custDict.Exists(custID) Then
                custDict(custID) = custDict(custID) + invAmt
            Else
                custDict.Add custID, invAmt
            End If
        Next i

        ' Output customers over threshold for confirmation
        Dim key As Variant
        Dim custBalance As Double
        Dim confirmCount As Long
        Dim confirmTotal As Double

        For Each key In custDict.Keys
            custBalance = custDict(key)
            If custBalance >= CONFIRM_THRESHOLD Then
                ' Get customer name
                For i = 2 To lastRow
                    If wsAR.Cells(i, 1).Value = key Then
                        custName = wsAR.Cells(i, 2).Value
                        Exit For
                    End If
                Next i

                .Cells(auditRow, 1).Value = custName
                .Cells(auditRow, 2).Value = custBalance
                .Cells(auditRow, 3).Value = custBalance / totalAR
                .Cells(auditRow, 4).Value = "YES"
                .Cells(auditRow, 4).Interior.Color = RGB(255, 235, 156)

                confirmCount = confirmCount + 1
                confirmTotal = confirmTotal + custBalance
                auditRow = auditRow + 1
            End If
        Next key

        ' Confirmation coverage summary
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "Confirmation Coverage:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Customers to Confirm:"
        .Cells(auditRow, 2).Value = confirmCount
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Dollar Coverage:"
        .Cells(auditRow, 2).Value = confirmTotal
        .Cells(auditRow, 3).Value = confirmTotal / totalAR
        .Cells(auditRow, 3).NumberFormat = "0.0%"

        ' Format
        .Range(.Cells(confStart, 2), .Cells(auditRow, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        .Range(.Cells(confStart, 3), .Cells(auditRow - 3, 3)).NumberFormat = "0.0%"

        auditRow = auditRow + 3
    End With

    ' ========================================
    ' TEST 4: SUBSEQUENT RECEIPTS
    ' ========================================
    If Not wsReceipts Is Nothing Then
        With wsAudit
            .Cells(auditRow, 1).Value = "TEST 4: SUBSEQUENT RECEIPTS TESTING"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Customer"
            .Cells(auditRow, 2).Value = "Y/E Balance"
            .Cells(auditRow, 3).Value = "Collected"
            .Cells(auditRow, 4).Value = "% Collected"
            .Cells(auditRow, 5).Value = "Collection Date"
            .Cells(auditRow, 6).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
            auditRow = auditRow + 1

            Dim subStart As Long
            subStart = auditRow

            ' Match subsequent receipts to Y/E balances
            Dim recLastRow As Long
            recLastRow = wsReceipts.Cells(wsReceipts.Rows.Count, "A").End(xlUp).Row

            ' Create collection dictionary
            Dim collDict As Object
            Set collDict = CreateObject("Scripting.Dictionary")

            For i = 2 To recLastRow
                custID = wsReceipts.Cells(i, 2).Value
                Dim collAmt As Double
                collAmt = wsReceipts.Cells(i, 5).Value

                If collDict.Exists(custID) Then
                    collDict(custID) = collDict(custID) + collAmt
                Else
                    collDict.Add custID, collAmt
                End If
            Next i

            ' Output results for customers over threshold
            Dim collected As Double
            Dim collectedTotal As Double
            Dim testedTotal As Double

            For Each key In custDict.Keys
                custBalance = custDict(key)

                If custBalance >= CONFIRM_THRESHOLD Then
                    ' Get customer name and collection
                    For i = 2 To lastRow
                        If wsAR.Cells(i, 1).Value = key Then
                            custName = wsAR.Cells(i, 2).Value
                            Exit For
                        End If
                    Next i

                    If collDict.Exists(key) Then
                        collected = collDict(key)
                    Else
                        collected = 0
                    End If

                    .Cells(auditRow, 1).Value = custName
                    .Cells(auditRow, 2).Value = custBalance
                    .Cells(auditRow, 3).Value = collected
                    .Cells(auditRow, 4).Value = collected / custBalance

                    collectedTotal = collectedTotal + collected
                    testedTotal = testedTotal + custBalance

                    If collected >= custBalance Then
                        .Cells(auditRow, 6).Value = "COLLECTED 100%"
                        .Cells(auditRow, 6).Interior.Color = RGB(198, 239, 206)
                    ElseIf collected >= custBalance * 0.5 Then
                        .Cells(auditRow, 6).Value = "Partial Collection"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 235, 156)
                    ElseIf collected > 0 Then
                        .Cells(auditRow, 6).Value = "Low Collection"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                    Else
                        .Cells(auditRow, 6).Value = "NOT COLLECTED"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                    End If

                    auditRow = auditRow + 1
                End If
            Next key

            ' Summary
            auditRow = auditRow + 1
            .Cells(auditRow, 1).Value = "Collection Summary:"
            .Cells(auditRow, 1).Font.Bold = True
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Total Tested:"
            .Cells(auditRow, 2).Value = testedTotal
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Total Collected:"
            .Cells(auditRow, 2).Value = collectedTotal
            auditRow = auditRow + 1

            .Cells(auditRow, 1).Value = "Collection Rate:"
            .Cells(auditRow, 2).Value = collectedTotal / testedTotal
            .Cells(auditRow, 2).NumberFormat = "0.0%"

            ' Format
            .Range(.Cells(subStart, 2), .Cells(auditRow - 2, 3)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            .Range(.Cells(subStart, 4), .Cells(auditRow - 4, 4)).NumberFormat = "0.0%"

            auditRow = auditRow + 3
        End With
    End If

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

        .Cells(auditRow, 1).Value = "Total AR Balance:"
        .Cells(auditRow, 2).Value = totalAR
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " GL to sub-ledger reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Aging analysis and allowance test"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Confirmation selection"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Subsequent receipts testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Send confirmations (manual)"
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
        .Columns("B:G").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "Accounts Receivable Audit Complete!" & vbCrLf & vbCrLf & _
           "Total AR: " & Format(totalAR, "$#,##0") & vbCrLf & _
           "Confirmations Selected: " & confirmCount & vbCrLf & _
           "Review the AR_Audit worksheet.", vbInformation

End Sub
```

---

### 2. AR Confirmation Letter Generator

```vba
Sub GenerateARConfirmations()
    '================================================
    ' Generate AR Confirmation Letters
    ' Creates individual confirmation requests
    '================================================

    Dim wsAR As Worksheet
    Dim wsConf As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim confNum As Long

    Const CONFIRM_THRESHOLD As Double = 10000

    On Error Resume Next
    Set wsAR = ThisWorkbook.Sheets("AR_Aging")
    On Error GoTo 0

    If wsAR Is Nothing Then
        MsgBox "AR_Aging sheet required.", vbExclamation
        Exit Sub
    End If

    ' Get unique customers
    Dim custDict As Object
    Set custDict = CreateObject("Scripting.Dictionary")

    lastRow = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        Dim custID As String
        Dim custName As String
        Dim invAmt As Double

        custID = wsAR.Cells(i, 1).Value
        custName = wsAR.Cells(i, 2).Value
        invAmt = wsAR.Cells(i, 6).Value

        If Not custDict.Exists(custID) Then
            custDict.Add custID, Array(custName, invAmt)
        Else
            Dim arr As Variant
            arr = custDict(custID)
            arr(1) = arr(1) + invAmt
            custDict(custID) = arr
        End If
    Next i

    Application.ScreenUpdating = False

    ' Create confirmations for customers over threshold
    Dim key As Variant
    Dim custData As Variant

    For Each key In custDict.Keys
        custData = custDict(key)

        If custData(1) >= CONFIRM_THRESHOLD Then
            confNum = confNum + 1

            Set wsConf = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsConf.Name = "ARConf-" & confNum

            With wsConf
                .Range("A1").Value = "[YOUR FIRM NAME]"
                .Range("A2").Value = "[Your Address]"

                .Range("A4").Value = Format(Date, "mmmm d, yyyy")

                .Range("A6").Value = custData(0)  ' Customer name
                .Range("A7").Value = "[Customer Address]"

                .Range("A10").Value = "RE: Confirmation of Account Balance"
                .Range("A10").Font.Bold = True

                .Range("A12").Value = "Dear Sir or Madam:"

                .Range("A14").Value = "Our auditors are conducting an audit of our financial statements. Please confirm"
                .Range("A15").Value = "directly to them the amount you owed us as of December 31, " & Year(Date) & "."

                .Range("A17").Value = "According to our records, your balance was:"
                .Range("A19").Value = Format(custData(1), "$#,##0.00")
                .Range("A19").Font.Bold = True
                .Range("A19").Font.Size = 14

                .Range("A22").Value = "Please indicate whether this agrees with your records:"
                .Range("A24").Value = "___ The balance is CORRECT"
                .Range("A25").Value = "___ The balance is INCORRECT (explain below)"

                .Range("A28").Value = "Explanation of difference:"
                .Range("A29:A32").Borders(xlEdgeBottom).LineStyle = xlContinuous

                .Range("A35").Value = "Signature: _______________________"
                .Range("A36").Value = "Title: _______________________"
                .Range("A37").Value = "Date: _______________________"

                .Range("A40").Value = "Return directly to our auditors at:"
                .Range("A41").Value = "[Auditor Address]"
            End With
        End If
    Next key

    Application.ScreenUpdating = True

    MsgBox "Generated " & confNum & " AR confirmations.", vbInformation

End Sub
```

---

### 3. Sales Cutoff Testing

```vba
Sub TestSalesCutoff()
    '================================================
    ' Sales Cutoff Testing
    ' Tests revenue transactions around year-end
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
    ThisWorkbook.Sheets("Sales_Cutoff").Delete
    On Error GoTo 0

    Set wsCutoff = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsCutoff.Name = "Sales_Cutoff"

    Application.ScreenUpdating = False

    With wsCutoff
        .Range("A1").Value = "SALES CUTOFF TESTING"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Testing transactions within " & CUTOFF_DAYS & " days of year-end"

        .Range("A4").Value = "Date"
        .Range("B4").Value = "JE Number"
        .Range("C4").Value = "Account"
        .Range("D4").Value = "Description"
        .Range("E4").Value = "Amount"
        .Range("F4").Value = "Days from Y/E"
        .Range("G4").Value = "Ship Date"
        .Range("H4").Value = "Cutoff Status"
        .Range("A4:H4").Font.Bold = True
        .Range("A4:H4").Interior.Color = RGB(0, 51, 102)
        .Range("A4:H4").Font.Color = RGB(255, 255, 255)

        cutoffRow = 5

        Dim yearEnd As Date
        yearEnd = DateSerial(Year(Date), 12, 31)

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            ' Filter for revenue accounts (4000s) and AR (1100s)
            If Left(wsGL.Cells(i, 3).Value, 1) = "4" Or Left(wsGL.Cells(i, 3).Value, 2) = "11" Then
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
                        .Cells(cutoffRow, 5).Value = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
                        .Cells(cutoffRow, 6).Value = daysFromYE

                        ' Status based on timing
                        If daysFromYE > 0 Then
                            .Cells(cutoffRow, 8).Value = "AFTER Y/E - VERIFY SHIP DATE"
                            .Cells(cutoffRow, 8).Interior.Color = RGB(255, 199, 206)
                        Else
                            .Cells(cutoffRow, 8).Value = "Before Y/E - OK if shipped"
                            .Cells(cutoffRow, 8).Interior.Color = RGB(255, 235, 156)
                        End If

                        cutoffRow = cutoffRow + 1
                    End If
                End If
            End If
        Next i

        ' Format
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 35
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 12
        .Columns("H").ColumnWidth = 25

        .Range("E:E").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    Application.ScreenUpdating = True

    MsgBox "Sales cutoff testing complete!" & vbCrLf & _
           "Review the Sales_Cutoff worksheet.", vbInformation

End Sub
```

---

## Assertions Tested

| Assertion | Test | Pass Criteria |
|-----------|------|---------------|
| **Existence** | Confirmations, subsequent receipts | Balance confirmed/collected |
| **Valuation** | Allowance analysis | Adequate reserve |
| **Completeness** | GL to sub-ledger tie | Balances agree |
| **Cutoff** | Sales around Y/E | Revenue in correct period |
| **Rights** | Customer agreements | Company owns receivables |

---

## Common Exceptions

| Exception | Cause | Resolution |
|-----------|-------|------------|
| Confirmation difference | Timing, disputes | Reconcile, alternative procedures |
| High over-90 | Collection issues | Increase allowance |
| GL/sub-ledger difference | Posting errors | Investigate, adjust |
| Cutoff error | Early/late recording | Propose adjustment |

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Cash](./cash.md) | [➡️ Inventory](./inventory.md)
