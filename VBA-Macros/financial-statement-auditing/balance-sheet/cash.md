# Cash & Cash Equivalents Audit VBA

> **Audit Cash Like a Pro** - Complete VBA procedures for auditing cash accounts per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 1000-1099 (typically) |
| **Assertions** | Existence, Completeness, Valuation, Rights |
| **Risk Level** | Moderate to High (fraud risk) |
| **Key Documents** | Bank statements, bank reconciliations, cutoff statements |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for cash accounts

| Column | Header | Example |
|--------|--------|---------|
| A | `Date` | 12/31/2024 |
| B | `JE_Number` | JE-2024-1234 |
| C | `Account` | 1010 |
| D | `Account_Name` | Operating Cash |
| E | `Description` | Wire transfer |
| F | `Debit` | 50000 |
| G | `Credit` | 0 |
| H | `Source` | AR |

### Input Sheet 2: `Bank_Recs`
Bank reconciliation detail

| Column | Header | Example |
|--------|--------|---------|
| A | `Account` | 1010 |
| B | `Bank_Name` | First National Bank |
| C | `Statement_Date` | 12/31/2024 |
| D | `Bank_Balance` | 125000 |
| E | `Book_Balance` | 118500 |
| F | `Deposits_In_Transit` | 8500 |
| G | `Outstanding_Checks` | 15000 |
| H | `Other_Reconciling` | 0 |
| I | `Reconciled_Balance` | 118500 |

### Input Sheet 3: `Bank_Statements`
Bank statement ending balances

| Column | Header | Example |
|--------|--------|---------|
| A | `Account` | 1010 |
| B | `Bank_Name` | First National Bank |
| C | `Statement_Balance` | 125000 |
| D | `Statement_Date` | 12/31/2024 |

---

## Audit Procedures

### 1. Bank Reconciliation Testing

```vba
Sub AuditCash_BankReconciliation()
    '================================================
    ' CASH AUDIT - Bank Reconciliation Testing
    ' Tests mathematical accuracy and reconciling items
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with cash transactions
    '   - Sheet "Bank_Recs" with bank reconciliation data
    '   - Sheet "Bank_Statements" with confirmed bank balances
    '
    ' OUTPUTS:
    '   - Creates "Cash_Audit" worksheet with test results
    '   - Flags exceptions and reconciling differences
    '
    ' ASSERTIONS TESTED:
    '   - Existence (bank balance confirmed)
    '   - Valuation (amounts reconcile)
    '   - Completeness (all accounts included)
    '================================================

    Dim wsGL As Worksheet
    Dim wsBankRec As Worksheet
    Dim wsBankStmt As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim auditRow As Long

    ' Materiality thresholds
    Const MATERIALITY As Double = 50000
    Const TRIVIAL As Double = 2500
    Const RECON_TOLERANCE As Double = 0.01  ' Penny tolerance

    ' Validate required sheets exist
    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsBankRec = ThisWorkbook.Sheets("Bank_Recs")
    Set wsBankStmt = ThisWorkbook.Sheets("Bank_Statements")
    On Error GoTo 0

    If wsGL Is Nothing Or wsBankRec Is Nothing Then
        MsgBox "Required sheets not found." & vbCrLf & _
               "Please ensure GL_Detail and Bank_Recs sheets exist.", vbCritical
        Exit Sub
    End If

    ' Create or clear audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Cash_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Cash_Audit"

    Application.ScreenUpdating = False

    ' ========================================
    ' SECTION 1: AUDIT HEADER
    ' ========================================
    With wsAudit
        .Range("A1").Value = "CASH & CASH EQUIVALENTS - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(Date, "12/31/yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        .Range("A4").Value = "Materiality: " & Format(MATERIALITY, "$#,##0")
        .Range("B4").Value = "Trivial: " & Format(TRIVIAL, "$#,##0")

        auditRow = 6
    End With

    ' ========================================
    ' SECTION 2: GL TO BANK RECONCILIATION TIE-OUT
    ' ========================================
    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 1: GL TO BANK RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 10)).Merge
        auditRow = auditRow + 1

        ' Headers
        .Cells(auditRow, 1).Value = "Account"
        .Cells(auditRow, 2).Value = "Bank Name"
        .Cells(auditRow, 3).Value = "GL Balance"
        .Cells(auditRow, 4).Value = "Bank Rec Balance"
        .Cells(auditRow, 5).Value = "Difference"
        .Cells(auditRow, 6).Value = "Bank Statement"
        .Cells(auditRow, 7).Value = "Deposits in Transit"
        .Cells(auditRow, 8).Value = "Outstanding Checks"
        .Cells(auditRow, 9).Value = "Recon Math Check"
        .Cells(auditRow, 10).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 10)).Font.Bold = True
        auditRow = auditRow + 1

        Dim recStartRow As Long
        recStartRow = auditRow

        ' Process each bank reconciliation
        lastRow = wsBankRec.Cells(wsBankRec.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            Dim acctNum As String
            Dim glBalance As Double
            Dim bankRecBalance As Double
            Dim bankStmtBalance As Double
            Dim dit As Double
            Dim oc As Double
            Dim otherRecon As Double
            Dim reconCalc As Double
            Dim diff As Double

            acctNum = wsBankRec.Cells(i, 1).Value

            ' Get GL balance for this account
            glBalance = GetGLBalance(wsGL, acctNum)

            ' Get bank rec data
            bankRecBalance = wsBankRec.Cells(i, 5).Value  ' Book balance from rec
            bankStmtBalance = wsBankRec.Cells(i, 4).Value ' Bank statement balance
            dit = wsBankRec.Cells(i, 6).Value             ' Deposits in transit
            oc = wsBankRec.Cells(i, 7).Value              ' Outstanding checks
            otherRecon = wsBankRec.Cells(i, 8).Value      ' Other reconciling

            ' Calculate what reconciled balance should be
            reconCalc = bankStmtBalance + dit - oc + otherRecon

            ' Difference between GL and bank rec
            diff = glBalance - bankRecBalance

            ' Output results
            .Cells(auditRow, 1).Value = acctNum
            .Cells(auditRow, 2).Value = wsBankRec.Cells(i, 2).Value
            .Cells(auditRow, 3).Value = glBalance
            .Cells(auditRow, 4).Value = bankRecBalance
            .Cells(auditRow, 5).Value = diff
            .Cells(auditRow, 6).Value = bankStmtBalance
            .Cells(auditRow, 7).Value = dit
            .Cells(auditRow, 8).Value = oc
            .Cells(auditRow, 9).Value = reconCalc - bankRecBalance  ' Math check

            ' Status determination
            If Abs(diff) < RECON_TOLERANCE And Abs(reconCalc - bankRecBalance) < RECON_TOLERANCE Then
                .Cells(auditRow, 10).Value = "PASS"
                .Cells(auditRow, 10).Interior.Color = RGB(198, 239, 206)
            ElseIf Abs(diff) < TRIVIAL Then
                .Cells(auditRow, 10).Value = "TRIVIAL DIFF"
                .Cells(auditRow, 10).Interior.Color = RGB(255, 235, 156)
            Else
                .Cells(auditRow, 10).Value = "EXCEPTION"
                .Cells(auditRow, 10).Interior.Color = RGB(255, 199, 206)
            End If

            auditRow = auditRow + 1
        Next i

        ' Format amounts
        .Range(.Cells(recStartRow, 3), .Cells(auditRow - 1, 9)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    ' ========================================
    ' SECTION 3: CUTOFF TESTING
    ' ========================================
    auditRow = auditRow + 2

    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 2: CASH CUTOFF TESTING"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Date"
        .Cells(auditRow, 2).Value = "JE Number"
        .Cells(auditRow, 3).Value = "Account"
        .Cells(auditRow, 4).Value = "Description"
        .Cells(auditRow, 5).Value = "Debit"
        .Cells(auditRow, 6).Value = "Credit"
        .Cells(auditRow, 7).Value = "Days from YE"
        .Cells(auditRow, 8).Value = "Cutoff Flag"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Font.Bold = True
        auditRow = auditRow + 1

        Dim cutoffStart As Long
        cutoffStart = auditRow

        ' Test transactions within 5 days of year-end
        Dim yearEnd As Date
        yearEnd = DateSerial(Year(Date), 12, 31)

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            ' Filter for cash accounts (1000-1099)
            If Left(wsGL.Cells(i, 3).Value, 2) = "10" Then
                Dim transDate As Date
                Dim daysFromYE As Long

                If IsDate(wsGL.Cells(i, 1).Value) Then
                    transDate = wsGL.Cells(i, 1).Value
                    daysFromYE = transDate - yearEnd

                    ' Flag transactions within 5 days before or after year-end
                    If Abs(daysFromYE) <= 5 Then
                        .Cells(auditRow, 1).Value = transDate
                        .Cells(auditRow, 2).Value = wsGL.Cells(i, 2).Value
                        .Cells(auditRow, 3).Value = wsGL.Cells(i, 3).Value
                        .Cells(auditRow, 4).Value = wsGL.Cells(i, 5).Value
                        .Cells(auditRow, 5).Value = wsGL.Cells(i, 6).Value
                        .Cells(auditRow, 6).Value = wsGL.Cells(i, 7).Value
                        .Cells(auditRow, 7).Value = daysFromYE

                        If daysFromYE > 0 Then
                            .Cells(auditRow, 8).Value = "AFTER Y/E - REVIEW"
                            .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
                        Else
                            .Cells(auditRow, 8).Value = "Before Y/E"
                            .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
                        End If

                        auditRow = auditRow + 1
                    End If
                End If
            End If
        Next i

        ' Format
        .Range(.Cells(cutoffStart, 5), .Cells(auditRow - 1, 6)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    ' ========================================
    ' SECTION 4: LARGE/UNUSUAL ITEMS
    ' ========================================
    auditRow = auditRow + 2

    With wsAudit
        .Cells(auditRow, 1).Value = "TEST 3: LARGE & UNUSUAL TRANSACTIONS"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Merge
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Date"
        .Cells(auditRow, 2).Value = "JE Number"
        .Cells(auditRow, 3).Value = "Account"
        .Cells(auditRow, 4).Value = "Description"
        .Cells(auditRow, 5).Value = "Amount"
        .Cells(auditRow, 6).Value = "% of Materiality"
        .Cells(auditRow, 7).Value = "Review Required"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 7)).Font.Bold = True
        auditRow = auditRow + 1

        Dim largeStart As Long
        largeStart = auditRow

        ' Threshold for large items (25% of materiality)
        Dim largeThreshold As Double
        largeThreshold = MATERIALITY * 0.25

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If Left(wsGL.Cells(i, 3).Value, 2) = "10" Then
                Dim transAmt As Double
                transAmt = wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value  ' Net amount

                If Abs(transAmt) >= largeThreshold Then
                    .Cells(auditRow, 1).Value = wsGL.Cells(i, 1).Value
                    .Cells(auditRow, 2).Value = wsGL.Cells(i, 2).Value
                    .Cells(auditRow, 3).Value = wsGL.Cells(i, 3).Value
                    .Cells(auditRow, 4).Value = wsGL.Cells(i, 5).Value
                    .Cells(auditRow, 5).Value = transAmt
                    .Cells(auditRow, 6).Value = Abs(transAmt) / MATERIALITY
                    .Cells(auditRow, 6).NumberFormat = "0%"

                    If Abs(transAmt) >= MATERIALITY Then
                        .Cells(auditRow, 7).Value = "MATERIAL - VOUCH"
                        .Cells(auditRow, 7).Interior.Color = RGB(255, 199, 206)
                    Else
                        .Cells(auditRow, 7).Value = "Review"
                        .Cells(auditRow, 7).Interior.Color = RGB(255, 235, 156)
                    End If

                    auditRow = auditRow + 1
                End If
            End If
        Next i

        .Range(.Cells(largeStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    ' ========================================
    ' SECTION 5: AUDIT SUMMARY
    ' ========================================
    auditRow = auditRow + 2

    With wsAudit
        .Cells(auditRow, 1).Value = "AUDIT SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Total GL Cash Balance:"
        .Cells(auditRow, 2).Value = GetTotalCashBalance(wsGL)
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Bank reconciliation tie-out"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Cutoff testing (5 days)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Large/unusual transaction review"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Bank confirmation (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " Subsequent disbursements (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion based on procedures performed and exceptions noted]"
        .Cells(auditRow, 1).Font.Italic = True

        ' Column widths
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 20
        .Columns("C:J").ColumnWidth = 15

    End With

    Application.ScreenUpdating = True

    MsgBox "Cash Audit Complete!" & vbCrLf & vbCrLf & _
           "Review the Cash_Audit worksheet for results.", vbInformation

End Sub

' Helper function to get GL balance for specific account
Private Function GetGLBalance(wsGL As Worksheet, acctNum As String) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim balance As Double

    lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
    balance = 0

    For i = 2 To lastRow
        If CStr(wsGL.Cells(i, 3).Value) = acctNum Then
            balance = balance + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
        End If
    Next i

    GetGLBalance = balance
End Function

' Helper function to get total cash balance
Private Function GetTotalCashBalance(wsGL As Worksheet) As Double
    Dim lastRow As Long
    Dim i As Long
    Dim balance As Double

    lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
    balance = 0

    For i = 2 To lastRow
        If Left(wsGL.Cells(i, 3).Value, 2) = "10" Then
            balance = balance + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
        End If
    Next i

    GetTotalCashBalance = balance
End Function
```

---

### 2. Bank Confirmation Tracking

```vba
Sub AuditCash_ConfirmationTracker()
    '================================================
    ' Track Bank Confirmation Status
    ' Creates confirmation log and tracks responses
    '================================================

    Dim wsConf As Worksheet
    Dim wsBankRec As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim confRow As Long

    On Error Resume Next
    Set wsBankRec = ThisWorkbook.Sheets("Bank_Recs")
    On Error GoTo 0

    If wsBankRec Is Nothing Then
        MsgBox "Bank_Recs sheet required.", vbExclamation
        Exit Sub
    End If

    ' Create confirmation tracker
    On Error Resume Next
    ThisWorkbook.Sheets("Cash_Confirmations").Delete
    On Error GoTo 0

    Set wsConf = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsConf.Name = "Cash_Confirmations"

    With wsConf
        .Range("A1").Value = "BANK CONFIRMATION TRACKER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        ' Headers
        .Range("A3").Value = "Bank Name"
        .Range("B3").Value = "Account Number"
        .Range("C3").Value = "Confirmation Date"
        .Range("D3").Value = "Sent Date"
        .Range("E3").Value = "Response Date"
        .Range("F3").Value = "Confirmed Balance"
        .Range("G3").Value = "Book Balance"
        .Range("H3").Value = "Difference"
        .Range("I3").Value = "Status"
        .Range("A3:I3").Font.Bold = True
        .Range("A3:I3").Interior.Color = RGB(0, 51, 102)
        .Range("A3:I3").Font.Color = RGB(255, 255, 255)

        confRow = 4
        lastRow = wsBankRec.Cells(wsBankRec.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            .Cells(confRow, 1).Value = wsBankRec.Cells(i, 2).Value  ' Bank name
            .Cells(confRow, 2).Value = wsBankRec.Cells(i, 1).Value  ' Account
            .Cells(confRow, 3).Value = "12/31/" & Year(Date)        ' Confirmation date
            .Cells(confRow, 7).Value = wsBankRec.Cells(i, 5).Value  ' Book balance
            .Cells(confRow, 8).Formula = "=F" & confRow & "-G" & confRow  ' Difference
            .Cells(confRow, 9).Value = "Pending"
            confRow = confRow + 1
        Next i

        ' Add data validation for status
        .Range("I4:I" & confRow - 1).Validation.Add _
            Type:=xlValidateList, _
            Formula1:="Pending,Sent,Received-Agrees,Received-Exception,No Response"

        ' Conditional formatting
        .Range("I4:I100").FormatConditions.Add Type:=xlTextString, String:="Agrees", TextOperator:=xlContains
        .Range("I4:I100").FormatConditions(1).Interior.Color = RGB(198, 239, 206)

        .Range("I4:I100").FormatConditions.Add Type:=xlTextString, String:="Exception", TextOperator:=xlContains
        .Range("I4:I100").FormatConditions(2).Interior.Color = RGB(255, 199, 206)

        ' Format
        .Columns("A").ColumnWidth = 25
        .Columns("B:I").ColumnWidth = 15
        .Range("F:H").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    End With

    MsgBox "Confirmation tracker created!", vbInformation

End Sub
```

---

### 3. Outstanding Check Analysis

```vba
Sub AuditCash_OutstandingChecks()
    '================================================
    ' Analyze Outstanding Checks
    ' Tests for stale-dated checks and unusual items
    '================================================

    Dim wsOC As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim auditRow As Long
    Dim staleCount As Long
    Dim staleAmount As Double

    ' Assumes Outstanding Checks are on a sheet called "Outstanding_Checks"
    ' Columns: A=Check#, B=Date, C=Payee, D=Amount, E=Cleared Date

    On Error Resume Next
    Set wsOC = ThisWorkbook.Sheets("Outstanding_Checks")
    On Error GoTo 0

    If wsOC Is Nothing Then
        MsgBox "Outstanding_Checks sheet required." & vbCrLf & vbCrLf & _
               "Format: Check#, Date, Payee, Amount, Cleared Date", vbExclamation
        Exit Sub
    End If

    ' Create audit results
    On Error Resume Next
    ThisWorkbook.Sheets("OC_Analysis").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "OC_Analysis"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "OUTSTANDING CHECK ANALYSIS"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A3").Value = "Check #"
        .Range("B3").Value = "Date"
        .Range("C3").Value = "Payee"
        .Range("D3").Value = "Amount"
        .Range("E3").Value = "Days Outstanding"
        .Range("F3").Value = "Status"
        .Range("A3:F3").Font.Bold = True
        .Range("A3:F3").Interior.Color = RGB(0, 51, 102)
        .Range("A3:F3").Font.Color = RGB(255, 255, 255)

        auditRow = 4
        lastRow = wsOC.Cells(wsOC.Rows.Count, "A").End(xlUp).Row

        For i = 2 To lastRow
            If IsEmpty(wsOC.Cells(i, 5).Value) Or wsOC.Cells(i, 5).Value = "" Then  ' Not cleared
                Dim checkDate As Date
                Dim daysOut As Long

                .Cells(auditRow, 1).Value = wsOC.Cells(i, 1).Value
                .Cells(auditRow, 2).Value = wsOC.Cells(i, 2).Value
                .Cells(auditRow, 3).Value = wsOC.Cells(i, 3).Value
                .Cells(auditRow, 4).Value = wsOC.Cells(i, 4).Value

                If IsDate(wsOC.Cells(i, 2).Value) Then
                    checkDate = wsOC.Cells(i, 2).Value
                    daysOut = Date - checkDate
                    .Cells(auditRow, 5).Value = daysOut

                    ' Stale dated = over 180 days (6 months)
                    If daysOut > 180 Then
                        .Cells(auditRow, 6).Value = "STALE - INVESTIGATE"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                        staleCount = staleCount + 1
                        staleAmount = staleAmount + wsOC.Cells(i, 4).Value
                    ElseIf daysOut > 90 Then
                        .Cells(auditRow, 6).Value = "Aging - Review"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 235, 156)
                    Else
                        .Cells(auditRow, 6).Value = "Current"
                        .Cells(auditRow, 6).Interior.Color = RGB(198, 239, 206)
                    End If
                End If

                auditRow = auditRow + 1
            End If
        Next i

        ' Summary
        auditRow = auditRow + 2
        .Cells(auditRow, 1).Value = "SUMMARY"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Total Outstanding Checks:"
        .Cells(auditRow, 2).Formula = "=SUM(D4:D" & (auditRow - 3) & ")"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Stale Dated (>180 days):"
        .Cells(auditRow, 2).Value = staleAmount
        .Cells(auditRow, 3).Value = staleCount & " items"
        If staleAmount > 0 Then
            .Cells(auditRow, 2).Interior.Color = RGB(255, 199, 206)
        End If

        ' Format
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 30
        .Columns("D:E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 20
        .Range("D:D").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Range("B:B").NumberFormat = "mm/dd/yyyy"

    End With

    Application.ScreenUpdating = True

    MsgBox "Outstanding Check Analysis Complete!" & vbCrLf & vbCrLf & _
           "Stale Items: " & staleCount & vbCrLf & _
           "Stale Amount: " & Format(staleAmount, "$#,##0"), vbInformation

End Sub
```

---

---

## Output Examples

### Generated `Cash_Audit` Worksheet

When you run `AuditCash_BankReconciliation()`, the macro generates a complete audit workpaper:

**Header Section:**
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ CASH & CASH EQUIVALENTS - AUDIT WORKPAPER                                   │
│ Period: 12/31/2024                                                          │
│ Prepared: JSMITH on 1/15/2025 2:34:22 PM                                    │
│ Materiality: $50,000    Trivial: $2,500                                     │
└─────────────────────────────────────────────────────────────────────────────┘
```

**TEST 1: GL TO BANK RECONCILIATION**
```
┌──────────┬───────────────────┬────────────┬────────────┬───────────┬────────────┬───────────┬────────────┬───────────┬───────────┐
│ Account  │ Bank Name         │ GL Balance │ Bank Rec   │ Diff      │ Bank Stmt  │ DIT       │ O/S Checks │ Recon Chk │ Status    │
├──────────┼───────────────────┼────────────┼────────────┼───────────┼────────────┼───────────┼────────────┼───────────┼───────────┤
│ 1010     │ First National    │ $118,500   │ $118,500   │ $0.00     │ $125,000   │ $8,500    │ $15,000    │ $0.00     │ ✓ PASS    │
│ 1020     │ Wells Fargo       │ $45,230    │ $45,230    │ $0.00     │ $52,730    │ $2,500    │ $10,000    │ $0.00     │ ✓ PASS    │
│ 1030     │ Chase Payroll     │ $22,150    │ $22,153    │ ($3.00)   │ $24,650    │ $0        │ $2,500     │ $0.00     │ ⚠ TRIVIAL │
│ 1040     │ Petty Cash        │ $500       │ $500       │ $0.00     │ N/A        │ N/A       │ N/A        │ $0.00     │ ✓ PASS    │
└──────────┴───────────────────┴────────────┴────────────┴───────────┴────────────┴───────────┴────────────┴───────────┴───────────┘
```
*Green = Pass | Yellow = Trivial | Red = Exception*

**TEST 2: CASH CUTOFF TESTING**
```
┌────────────┬─────────────┬─────────┬────────────────────────────┬────────────┬───────────┬───────────┬──────────────────┐
│ Date       │ JE Number   │ Account │ Description                │ Debit      │ Credit    │ Days Y/E  │ Cutoff Flag      │
├────────────┼─────────────┼─────────┼────────────────────────────┼────────────┼───────────┼───────────┼──────────────────┤
│ 12/27/2024 │ JE-2024-892 │ 1010    │ Customer Payment - ABC Co  │ $15,000    │           │ -4        │ ✓ Before Y/E     │
│ 12/30/2024 │ JE-2024-901 │ 1010    │ Wire Transfer - XYZ Inc    │ $45,000    │           │ -1        │ ✓ Before Y/E     │
│ 12/31/2024 │ JE-2024-912 │ 1020    │ Check #4521 - Vendor Pay   │            │ $8,500    │ 0         │ ✓ Before Y/E     │
│ 01/02/2025 │ JE-2025-003 │ 1010    │ Customer Deposit           │ $22,000    │           │ +2        │ ✗ AFTER Y/E      │
│ 01/03/2025 │ JE-2025-008 │ 1020    │ Wire - January Receipt     │ $35,000    │           │ +3        │ ✗ AFTER Y/E      │
└────────────┴─────────────┴─────────┴────────────────────────────┴────────────┴───────────┴───────────┴──────────────────┘
```
*Red items require investigation to ensure proper period*

**TEST 3: LARGE & UNUSUAL TRANSACTIONS**
```
┌────────────┬─────────────┬─────────┬─────────────────────────────┬────────────┬──────────────┬────────────────┐
│ Date       │ JE Number   │ Account │ Description                 │ Amount     │ % Material   │ Review Req'd   │
├────────────┼─────────────┼─────────┼─────────────────────────────┼────────────┼──────────────┼────────────────┤
│ 03/15/2024 │ JE-2024-156 │ 1010    │ Equipment Purchase - Truck  │ ($75,000)  │ 150%         │ ✗ MATERIAL     │
│ 06/22/2024 │ JE-2024-312 │ 1010    │ Insurance Proceeds          │ $52,000    │ 104%         │ ✗ MATERIAL     │
│ 09/30/2024 │ JE-2024-498 │ 1020    │ Dividend Payment            │ ($45,000)  │ 90%          │ ⚠ Review       │
│ 11/15/2024 │ JE-2024-623 │ 1010    │ Building Down Payment       │ ($125,000) │ 250%         │ ✗ MATERIAL     │
│ 12/28/2024 │ JE-2024-895 │ 1010    │ Year-End Customer Wire      │ $38,500    │ 77%          │ ⚠ Review       │
└────────────┴─────────────┴─────────┴─────────────────────────────┴────────────┴──────────────┴────────────────┘
```

**AUDIT SUMMARY**
```
┌─────────────────────────────────────────────────────────────────┐
│ AUDIT SUMMARY                                                   │
├─────────────────────────────────────────────────────────────────┤
│ Total GL Cash Balance:                          $186,380.00     │
│                                                                 │
│ Procedures Performed:                                           │
│   ✓ Bank reconciliation tie-out                                 │
│   ✓ Cutoff testing (5 days)                                     │
│   ✓ Large/unusual transaction review                            │
│   ☐ Bank confirmation (manual)                                  │
│   ☐ Subsequent disbursements (manual)                           │
│                                                                 │
│ CONCLUSION:                                                     │
│ [Document conclusion based on procedures performed]             │
└─────────────────────────────────────────────────────────────────┘
```

### Generated `Cash_Confirmations` Worksheet

```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ BANK CONFIRMATION TRACKER                                                                           │
├───────────────────┬─────────────┬──────────────┬────────────┬────────────┬────────────┬────────────┤
│ Bank Name         │ Account #   │ Confirm Date │ Sent Date  │ Resp Date  │ Confirmed  │ Book Bal   │
├───────────────────┼─────────────┼──────────────┼────────────┼────────────┼────────────┼────────────┤
│ First National    │ 1010        │ 12/31/2024   │ 01/05/2025 │ 01/12/2025 │ $125,000   │ $125,000   │
│ Wells Fargo       │ 1020        │ 12/31/2024   │ 01/05/2025 │            │            │ $52,730    │
│ Chase Payroll     │ 1030        │ 12/31/2024   │ 01/05/2025 │ 01/15/2025 │ $24,650    │ $24,650    │
├───────────────────┴─────────────┴──────────────┴────────────┴────────────┴────────────┴────────────┤
│ Difference │ Status              │                                                                 │
├────────────┼─────────────────────┤                                                                 │
│ $0.00      │ ✓ Received-Agrees   │                                                                 │
│            │ ⚠ Pending           │                                                                 │
│ $0.00      │ ✓ Received-Agrees   │                                                                 │
└────────────┴─────────────────────┴─────────────────────────────────────────────────────────────────┘
```

### Generated `OC_Analysis` Worksheet

```
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│ OUTSTANDING CHECK ANALYSIS                                                              │
├──────────┬────────────┬─────────────────────────────┬────────────┬───────────┬──────────┤
│ Check #  │ Date       │ Payee                       │ Amount     │ Days Out  │ Status   │
├──────────┼────────────┼─────────────────────────────┼────────────┼───────────┼──────────┤
│ 4512     │ 12/28/2024 │ Office Depot                │ $245.00    │ 18        │ ✓ Current│
│ 4508     │ 12/15/2024 │ ABC Supplies                │ $1,250.00  │ 31        │ ✓ Current│
│ 4485     │ 11/02/2024 │ Johnson Electric            │ $3,500.00  │ 74        │ ✓ Current│
│ 4421     │ 09/15/2024 │ Smith Consulting            │ $2,800.00  │ 122       │ ⚠ Aging  │
│ 4389     │ 05/22/2024 │ Old Vendor Inc              │ $750.00    │ 238       │ ✗ STALE  │
│ 4356     │ 03/10/2024 │ Former Employee Reimburse   │ $425.00    │ 311       │ ✗ STALE  │
├──────────┴────────────┴─────────────────────────────┴────────────┴───────────┴──────────┤
│ SUMMARY                                                                                 │
│ Total Outstanding Checks:                                           $8,970.00          │
│ Stale Dated (>180 days):                                           $1,175.00 (2 items) │
└─────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## Assertions Tested

| Assertion | Test | Pass Criteria |
|-----------|------|---------------|
| **Existence** | Bank confirmation | Confirmed balance matches |
| **Completeness** | All accounts reconciled | No missing accounts |
| **Valuation** | Reconciliation math | GL = Reconciled balance |
| **Cutoff** | Period-end transactions | Proper period recorded |
| **Rights** | Bank confirmation | Account in company name |

---

## Common Exceptions

| Exception | Cause | Resolution |
|-----------|-------|------------|
| Reconciliation difference | Timing, errors | Investigate, propose adjustment |
| Stale-dated checks | Vendor didn't cash | Write off to income |
| Cutoff error | Wrong period | Propose reclassification |
| Unreconciled account | Missing bank rec | Request from client |

---

## Sign-Off

```
Prepared By: _______________ Date: _______________

Reviewed By: _______________ Date: _______________

Conclusion:
[ ] Cash is fairly stated in all material respects
[ ] Proposed adjustments required (see schedule)
[ ] Additional procedures needed
```

---

[⬅️ Back to FS Auditing](../README.md) | [➡️ Next: Accounts Receivable](./accounts-receivable.md)
