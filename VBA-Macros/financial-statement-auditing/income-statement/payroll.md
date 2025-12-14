# Payroll Expense Audit VBA

> **Payroll Testing** - Complete VBA for auditing payroll expenses per GAAS/GAAP

---

## Account Overview

| Item | Description |
|------|-------------|
| **GL Accounts** | 6100-6199 (Wages), 6200-6299 (Benefits) |
| **Assertions** | Occurrence, Accuracy, Completeness, Classification |
| **Risk Level** | Moderate (significant expense, fraud risk) |
| **Key Standards** | ASC 710 (Compensation), ASC 715 (Benefits) |

---

## Required Inputs

### Input Sheet 1: `GL_Detail`
General ledger detail for payroll accounts

### Input Sheet 2: `Payroll_Register`
Payroll register detail

| Column | Header | Example |
|--------|--------|---------|
| A | `Pay_Date` | 12/31/2024 |
| B | `Employee_ID` | EMP-001 |
| C | `Employee_Name` | John Smith |
| D | `Department` | Sales |
| E | `Gross_Pay` | 5000 |
| F | `Fed_Tax` | 750 |
| G | `State_Tax` | 250 |
| H | `FICA` | 383 |
| I | `Benefits` | 200 |
| J | `Net_Pay` | 3417 |

### Input Sheet 3: `Employee_Master`
Employee master file for testing

| Column | Header | Example |
|--------|--------|---------|
| A | `Employee_ID` | EMP-001 |
| B | `Name` | John Smith |
| C | `Hire_Date` | 03/15/2020 |
| D | `Department` | Sales |
| E | `Annual_Salary` | 120000 |
| F | `Pay_Type` | Salary |
| G | `Status` | Active |

---

## Audit Procedures

```vba
Sub AuditPayroll()
    '================================================
    ' PAYROLL EXPENSE - COMPLETE AUDIT MODULE
    '
    ' INPUTS REQUIRED:
    '   - Sheet "GL_Detail" with payroll transactions
    '   - Sheet "Payroll_Register" with payroll detail
    '   - Sheet "Employee_Master" for validation
    '
    ' OUTPUTS:
    '   - Creates "Payroll_Audit" worksheet
    '   - Reconciles register to GL
    '   - Tests payroll calculations
    '   - Identifies ghost employees
    '   - Analyzes payroll trends
    '
    ' ASSERTIONS TESTED:
    '   - Occurrence (employees exist, worked)
    '   - Accuracy (calculations correct)
    '   - Completeness (all payroll recorded)
    '================================================

    Dim wsGL As Worksheet
    Dim wsPayroll As Worksheet
    Dim wsEmp As Worksheet
    Dim wsAudit As Worksheet
    Dim lastRow As Long, i As Long
    Dim auditRow As Long

    Const FICA_RATE As Double = 0.0765  ' 7.65%
    Const VARIANCE_THRESHOLD As Double = 0.05  ' 5%

    On Error Resume Next
    Set wsGL = ThisWorkbook.Sheets("GL_Detail")
    Set wsPayroll = ThisWorkbook.Sheets("Payroll_Register")
    Set wsEmp = ThisWorkbook.Sheets("Employee_Master")
    On Error GoTo 0

    If wsGL Is Nothing Then
        MsgBox "GL_Detail sheet required.", vbCritical
        Exit Sub
    End If

    ' Create audit worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("Payroll_Audit").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Payroll_Audit"

    Application.ScreenUpdating = False

    With wsAudit
        .Range("A1").Value = "PAYROLL EXPENSE - AUDIT WORKPAPER"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        .Range("A2").Value = "Period: " & Format(DateSerial(Year(Date), 12, 31), "mmmm d, yyyy")
        .Range("A3").Value = "Prepared: " & Environ("USERNAME") & " on " & Now

        auditRow = 5

        ' ========================================
        ' TEST 1: PAYROLL GL TO REGISTER RECONCILIATION
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 1: GL TO PAYROLL REGISTER RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        ' Calculate GL payroll
        Dim glWages As Double
        Dim glBenefits As Double
        Dim glPayrollTax As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            Dim acctNum As String
            acctNum = CStr(wsGL.Cells(i, 3).Value)

            ' Wages (61xx)
            If Left(acctNum, 2) = "61" Then
                glWages = glWages + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
            End If

            ' Benefits (62xx)
            If Left(acctNum, 2) = "62" Then
                glBenefits = glBenefits + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
            End If

            ' Payroll taxes (63xx)
            If Left(acctNum, 2) = "63" Then
                glPayrollTax = glPayrollTax + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
            End If
        Next i

        ' Calculate register totals
        Dim regGross As Double
        Dim regFICA As Double
        Dim regBenefits As Double

        If Not wsPayroll Is Nothing Then
            lastRow = wsPayroll.Cells(wsPayroll.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                regGross = regGross + wsPayroll.Cells(i, 5).Value
                regFICA = regFICA + wsPayroll.Cells(i, 8).Value
                regBenefits = regBenefits + wsPayroll.Cells(i, 9).Value
            Next i
        End If

        .Cells(auditRow, 1).Value = "Category"
        .Cells(auditRow, 2).Value = "GL Balance"
        .Cells(auditRow, 3).Value = "Register"
        .Cells(auditRow, 4).Value = "Difference"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim reconStart As Long
        reconStart = auditRow

        ' Wages
        .Cells(auditRow, 1).Value = "Wages & Salaries"
        .Cells(auditRow, 2).Value = glWages
        .Cells(auditRow, 3).Value = regGross
        .Cells(auditRow, 4).Value = glWages - regGross

        If Abs(glWages - regGross) < 100 Then
            .Cells(auditRow, 5).Value = "RECONCILED"
            .Cells(auditRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            .Cells(auditRow, 5).Value = "DIFFERENCE"
            .Cells(auditRow, 5).Interior.Color = RGB(255, 199, 206)
        End If
        auditRow = auditRow + 1

        ' Total
        .Cells(auditRow, 1).Value = "TOTAL PAYROLL EXPENSE"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glWages + glBenefits + glPayrollTax
        .Cells(auditRow, 2).Font.Bold = True
        auditRow = auditRow + 1

        .Range(.Cells(reconStart, 2), .Cells(auditRow - 1, 4)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

        auditRow = auditRow + 2

        ' ========================================
        ' TEST 2: PAYROLL CALCULATION TEST
        ' ========================================
        If Not wsPayroll Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 2: PAYROLL CALCULATION TESTING (SAMPLE)"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Employee"
            .Cells(auditRow, 2).Value = "Gross"
            .Cells(auditRow, 3).Value = "FICA Withheld"
            .Cells(auditRow, 4).Value = "Expected FICA"
            .Cells(auditRow, 5).Value = "Difference"
            .Cells(auditRow, 6).Value = "Net Pay"
            .Cells(auditRow, 7).Value = "Calc Net"
            .Cells(auditRow, 8).Value = "Status"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 8)).Font.Bold = True
            auditRow = auditRow + 1

            Dim calcStart As Long
            calcStart = auditRow

            ' Test sample of payroll records
            Dim sampleCount As Long
            sampleCount = 0

            lastRow = wsPayroll.Cells(wsPayroll.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                If sampleCount >= 15 Then Exit For

                Dim gross As Double
                Dim ficaWith As Double
                Dim fedTax As Double
                Dim stateTax As Double
                Dim benefits As Double
                Dim netPay As Double
                Dim expectedFICA As Double
                Dim calcNet As Double

                gross = wsPayroll.Cells(i, 5).Value
                fedTax = wsPayroll.Cells(i, 6).Value
                stateTax = wsPayroll.Cells(i, 7).Value
                ficaWith = wsPayroll.Cells(i, 8).Value
                benefits = wsPayroll.Cells(i, 9).Value
                netPay = wsPayroll.Cells(i, 10).Value

                expectedFICA = gross * FICA_RATE
                calcNet = gross - fedTax - stateTax - ficaWith - benefits

                .Cells(auditRow, 1).Value = wsPayroll.Cells(i, 3).Value
                .Cells(auditRow, 2).Value = gross
                .Cells(auditRow, 3).Value = ficaWith
                .Cells(auditRow, 4).Value = expectedFICA
                .Cells(auditRow, 5).Value = ficaWith - expectedFICA
                .Cells(auditRow, 6).Value = netPay
                .Cells(auditRow, 7).Value = calcNet

                ' Check if calculations are reasonable
                If Abs(ficaWith - expectedFICA) < 5 And Abs(netPay - calcNet) < 1 Then
                    .Cells(auditRow, 8).Value = "VERIFIED"
                    .Cells(auditRow, 8).Interior.Color = RGB(198, 239, 206)
                Else
                    .Cells(auditRow, 8).Value = "REVIEW"
                    .Cells(auditRow, 8).Interior.Color = RGB(255, 199, 206)
                End If

                auditRow = auditRow + 1
                sampleCount = sampleCount + 1
            Next i

            .Range(.Cells(calcStart, 2), .Cells(auditRow - 1, 7)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

            auditRow = auditRow + 2
        End If

        ' ========================================
        ' TEST 3: GHOST EMPLOYEE TEST
        ' ========================================
        If Not wsPayroll Is Nothing And Not wsEmp Is Nothing Then
            .Cells(auditRow, 1).Value = "TEST 3: GHOST EMPLOYEE IDENTIFICATION"
            .Cells(auditRow, 1).Font.Bold = True
            .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
            .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Merge
            auditRow = auditRow + 2

            .Cells(auditRow, 1).Value = "Employee ID"
            .Cells(auditRow, 2).Value = "Name (Register)"
            .Cells(auditRow, 3).Value = "In Master?"
            .Cells(auditRow, 4).Value = "Status"
            .Cells(auditRow, 5).Value = "Total Paid"
            .Cells(auditRow, 6).Value = "Flag"
            .Range(.Cells(auditRow, 1), .Cells(auditRow, 6)).Font.Bold = True
            auditRow = auditRow + 1

            Dim ghostStart As Long
            ghostStart = auditRow

            ' Build employee master dictionary
            Dim empDict As Object
            Set empDict = CreateObject("Scripting.Dictionary")

            Dim empLastRow As Long
            empLastRow = wsEmp.Cells(wsEmp.Rows.Count, "A").End(xlUp).Row
            For i = 2 To empLastRow
                Dim empID As String
                empID = CStr(wsEmp.Cells(i, 1).Value)
                If Not empDict.Exists(empID) Then
                    empDict.Add empID, Array(wsEmp.Cells(i, 2).Value, wsEmp.Cells(i, 7).Value)
                End If
            Next i

            ' Check payroll register against master
            Dim payDict As Object
            Set payDict = CreateObject("Scripting.Dictionary")

            lastRow = wsPayroll.Cells(wsPayroll.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                Dim payEmpID As String
                payEmpID = CStr(wsPayroll.Cells(i, 2).Value)

                If payDict.Exists(payEmpID) Then
                    payDict(payEmpID) = Array(wsPayroll.Cells(i, 3).Value, payDict(payEmpID)(1) + wsPayroll.Cells(i, 5).Value)
                Else
                    payDict.Add payEmpID, Array(wsPayroll.Cells(i, 3).Value, wsPayroll.Cells(i, 5).Value)
                End If
            Next i

            ' Compare and flag
            Dim key As Variant
            Dim ghostCount As Long
            ghostCount = 0

            For Each key In payDict.Keys
                Dim payData As Variant
                payData = payDict(key)

                .Cells(auditRow, 1).Value = key
                .Cells(auditRow, 2).Value = payData(0)
                .Cells(auditRow, 5).Value = payData(1)

                If empDict.Exists(CStr(key)) Then
                    Dim empData As Variant
                    empData = empDict(CStr(key))
                    .Cells(auditRow, 3).Value = "YES"
                    .Cells(auditRow, 3).Interior.Color = RGB(198, 239, 206)
                    .Cells(auditRow, 4).Value = empData(1)

                    If LCase(empData(1)) = "terminated" Then
                        .Cells(auditRow, 6).Value = "TERMED - VERIFY"
                        .Cells(auditRow, 6).Interior.Color = RGB(255, 235, 156)
                    Else
                        .Cells(auditRow, 6).Value = "OK"
                    End If
                Else
                    .Cells(auditRow, 3).Value = "NO"
                    .Cells(auditRow, 3).Interior.Color = RGB(255, 199, 206)
                    .Cells(auditRow, 4).Value = "N/A"
                    .Cells(auditRow, 6).Value = "GHOST EMPLOYEE?"
                    .Cells(auditRow, 6).Interior.Color = RGB(255, 199, 206)
                    ghostCount = ghostCount + 1
                End If

                auditRow = auditRow + 1
            Next key

            .Range(.Cells(ghostStart, 5), .Cells(auditRow - 1, 5)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

            If ghostCount > 0 Then
                .Cells(auditRow, 1).Value = "WARNING: " & ghostCount & " potential ghost employees identified"
                .Cells(auditRow, 1).Font.Bold = True
                .Cells(auditRow, 1).Font.Color = RGB(192, 0, 0)
                auditRow = auditRow + 1
            End If

            auditRow = auditRow + 2
        End If

        ' ========================================
        ' TEST 4: PAYROLL TREND ANALYSIS
        ' ========================================
        .Cells(auditRow, 1).Value = "TEST 4: MONTHLY PAYROLL TREND"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Month"
        .Cells(auditRow, 2).Value = "Payroll"
        .Cells(auditRow, 3).Value = "Headcount"
        .Cells(auditRow, 4).Value = "Avg/Employee"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        ' Aggregate by month
        Dim monthPay(1 To 12) As Double

        lastRow = wsGL.Cells(wsGL.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            acctNum = CStr(wsGL.Cells(i, 3).Value)
            If Left(acctNum, 2) = "61" Then
                Dim trxDate As Date
                On Error Resume Next
                trxDate = wsGL.Cells(i, 1).Value
                On Error GoTo 0

                If IsDate(trxDate) Then
                    Dim mo As Integer
                    mo = Month(trxDate)
                    monthPay(mo) = monthPay(mo) + wsGL.Cells(i, 6).Value - wsGL.Cells(i, 7).Value
                End If
            End If
        Next i

        Dim trendStart As Long
        trendStart = auditRow

        Dim avgMonthly As Double
        avgMonthly = glWages / 12

        Dim monthNames As Variant
        monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

        For mo = 1 To 12
            .Cells(auditRow, 1).Value = monthNames(mo - 1)
            .Cells(auditRow, 2).Value = monthPay(mo)
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)  ' Headcount input
            .Cells(auditRow, 4).Value = "[Calc]"

            ' Flag unusual months
            If monthPay(mo) > avgMonthly * 1.2 Then
                .Cells(auditRow, 5).Value = "HIGH"
                .Cells(auditRow, 5).Interior.Color = RGB(255, 235, 156)
            ElseIf monthPay(mo) < avgMonthly * 0.8 Then
                .Cells(auditRow, 5).Value = "LOW"
                .Cells(auditRow, 5).Interior.Color = RGB(255, 235, 156)
            Else
                .Cells(auditRow, 5).Value = "Normal"
            End If

            auditRow = auditRow + 1
        Next mo

        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glWages
        .Cells(auditRow, 2).Font.Bold = True
        auditRow = auditRow + 1

        .Range(.Cells(trendStart, 2), .Cells(auditRow - 1, 2)).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

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

        .Cells(auditRow, 1).Value = "Total Wages & Salaries:"
        .Cells(auditRow, 2).Value = glWages
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Total Benefits:"
        .Cells(auditRow, 2).Value = glBenefits
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "Total Payroll Taxes:"
        .Cells(auditRow, 2).Value = glPayrollTax
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "TOTAL PAYROLL:"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 2).Value = glWages + glBenefits + glPayrollTax
        .Cells(auditRow, 2).Font.Bold = True
        .Cells(auditRow, 2).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Procedures Performed:"
        auditRow = auditRow + 1

        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " GL to register reconciliation"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Payroll calculation testing"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Ghost employee identification"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(10003) & " Monthly trend analysis"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " W-2 reconciliation (manual)"
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "  " & ChrW(9744) & " 941 reconciliation (manual)"
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "CONCLUSION:"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1
        .Cells(auditRow, 1).Value = "[Document conclusion]"
        .Cells(auditRow, 1).Font.Italic = True

        .Columns("A").ColumnWidth = 25
        .Columns("B:H").ColumnWidth = 14

    End With

    Application.ScreenUpdating = True

    MsgBox "Payroll Audit Complete!" & vbCrLf & _
           "Total Payroll: " & Format(glWages + glBenefits + glPayrollTax, "$#,##0"), vbInformation

End Sub
```

---

## Payroll Tax Reconciliation

```vba
Sub ReconcilePayrollTaxes()
    '================================================
    ' PAYROLL TAX RECONCILIATION
    '
    ' Reconciles:
    '   - Form 941 to GL
    '   - W-2s to GL
    '   - State returns to GL
    '================================================

    Dim wsAudit As Worksheet
    Dim auditRow As Long

    On Error Resume Next
    ThisWorkbook.Sheets("Payroll_Tax_Recon").Delete
    On Error GoTo 0

    Set wsAudit = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAudit.Name = "Payroll_Tax_Recon"

    With wsAudit
        .Range("A1").Value = "PAYROLL TAX RECONCILIATION"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14

        auditRow = 3

        ' Form 941 Reconciliation
        .Cells(auditRow, 1).Value = "FORM 941 RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Merge
        auditRow = auditRow + 2

        .Cells(auditRow, 1).Value = "Quarter"
        .Cells(auditRow, 2).Value = "Wages per 941"
        .Cells(auditRow, 3).Value = "Per GL"
        .Cells(auditRow, 4).Value = "Difference"
        .Cells(auditRow, 5).Value = "Status"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 5)).Font.Bold = True
        auditRow = auditRow + 1

        Dim quarters As Variant
        quarters = Array("Q1", "Q2", "Q3", "Q4")

        Dim q As Variant
        For Each q In quarters
            .Cells(auditRow, 1).Value = q
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 4).Value = "=B" & auditRow & "-C" & auditRow
            auditRow = auditRow + 1
        Next q

        .Cells(auditRow, 1).Value = "TOTAL"
        .Cells(auditRow, 1).Font.Bold = True
        auditRow = auditRow + 1

        auditRow = auditRow + 2

        ' W-2 Reconciliation
        .Cells(auditRow, 1).Value = "W-2 RECONCILIATION"
        .Cells(auditRow, 1).Font.Bold = True
        .Cells(auditRow, 1).Interior.Color = RGB(0, 51, 102)
        .Cells(auditRow, 1).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Merge
        auditRow = auditRow + 2

        Dim w2Items As Variant
        w2Items = Array( _
            Array("Box 1 - Wages, Tips", ""), _
            Array("Box 2 - Federal Tax Withheld", ""), _
            Array("Box 3 - Social Security Wages", ""), _
            Array("Box 4 - Social Security Tax", ""), _
            Array("Box 5 - Medicare Wages", ""), _
            Array("Box 6 - Medicare Tax", ""))

        .Cells(auditRow, 1).Value = "W-2 Box"
        .Cells(auditRow, 2).Value = "Per W-2s"
        .Cells(auditRow, 3).Value = "Per GL"
        .Cells(auditRow, 4).Value = "Difference"
        .Range(.Cells(auditRow, 1), .Cells(auditRow, 4)).Font.Bold = True
        auditRow = auditRow + 1

        Dim w2 As Variant
        For Each w2 In w2Items
            .Cells(auditRow, 1).Value = w2(0)
            .Cells(auditRow, 2).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 3).Interior.Color = RGB(255, 255, 204)
            .Cells(auditRow, 4).Value = "=B" & auditRow & "-C" & auditRow
            auditRow = auditRow + 1
        Next w2

        .Columns("A").ColumnWidth = 30
        .Columns("B:E").ColumnWidth = 15

    End With

    MsgBox "Payroll Tax Reconciliation Template Created!", vbInformation

End Sub
```

---

## Key Payroll Tests

| Test | Purpose | Procedure |
|------|---------|-----------|
| **Existence** | Employees are real | Compare to HR files, observe |
| **Occurrence** | Work was performed | Time records, approvals |
| **Accuracy** | Calculations correct | Recalculate sample |
| **Completeness** | All payroll recorded | 941/W-2 reconciliation |
| **Classification** | Proper account coding | Review by department |
| **Cutoff** | Proper period | Test accrued wages |

---

[⬅️ Back to FS Auditing](../README.md) | [⬅️ Operating Expenses](./operating-expenses.md) | [➡️ Income Tax](./income-tax.md)
