# Audit Procedures VBA Macros

> **There's a VBA for That!** - Sample selection, confirmations, testing, and audit documentation

---

## Quick Reference

| Macro | What It Does |
|-------|--------------|
| [RandomSampleSelection](#random-sample-selection) | Select random items for testing |
| [MonetaryUnitSampling](#monetary-unit-sampling) | MUS/PPS sampling |
| [StratifiedSample](#stratified-sample) | Sample by strata/population groups |
| [GenerateConfirmations](#generate-confirmations) | Create confirmation letters |
| [TrackConfirmations](#track-confirmation-responses) | Monitor confirmation status |
| [TestFootings](#test-footings) | Verify mathematical accuracy |
| [CompareListings](#compare-two-listings) | Match two populations |
| [AgeReceivables](#age-receivables) | AR aging analysis |
| [SubsequentReceipts](#subsequent-receipts-testing) | Test subsequent collections |
| [CutoffTesting](#cutoff-testing) | Revenue/expense cutoff |

---

## Random Sample Selection

Select random items from a population for testing.

```vba
Sub RandomSampleSelection()
    '================================================
    ' Random Sample Selection
    ' Selects random items from population
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sampleSize As Long
    Dim popSize As Long
    Dim i As Long
    Dim randomRow As Long
    Dim selectedRows As Collection
    Dim markerCol As Integer

    Set ws = ActiveSheet
    Set selectedRows = New Collection

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    popSize = lastRow - 1  ' Excluding header

    sampleSize = Application.InputBox("Enter sample size:", "Random Sample", 25, Type:=1)
    If sampleSize = 0 Then Exit Sub

    If sampleSize > popSize Then
        MsgBox "Sample size cannot exceed population (" & popSize & ").", vbExclamation
        Exit Sub
    End If

    markerCol = Application.InputBox("Enter column number for sample marker:", "Random Sample", ws.UsedRange.Columns.Count + 1, Type:=1)

    ' Header for marker column
    ws.Cells(1, markerCol).Value = "Sample"
    ws.Cells(1, markerCol).Font.Bold = True

    Randomize

    ' Select random rows
    Do While selectedRows.Count < sampleSize
        randomRow = Int((popSize * Rnd) + 2)  ' +2 because row 1 is header

        ' Check if already selected
        On Error Resume Next
        selectedRows.Add randomRow, CStr(randomRow)
        If Err.Number = 0 Then
            ' Mark as selected
            ws.Cells(randomRow, markerCol).Value = "X"
            ws.Cells(randomRow, markerCol).Interior.Color = RGB(255, 255, 0)
        End If
        On Error GoTo 0
    Loop

    MsgBox "Selected " & sampleSize & " random items." & vbCrLf & _
           "Marked in column " & markerCol, vbInformation

End Sub

Sub RandomSampleWithSeed()
    '================================================
    ' Random Sample with Seed for Reproducibility
    '================================================

    Dim seed As Long

    seed = Application.InputBox("Enter random seed (for reproducibility):", "Random Seed", 12345, Type:=1)
    If seed = 0 Then Exit Sub

    Rnd (-1)
    Randomize seed

    Call RandomSampleSelection

End Sub
```

---

## Monetary Unit Sampling

Probability Proportional to Size (PPS) sampling.

```vba
Sub MonetaryUnitSampling()
    '================================================
    ' Monetary Unit Sampling (MUS/PPS)
    ' Selects items based on cumulative dollar amounts
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalPop As Double
    Dim sampleSize As Long
    Dim samplingInterval As Double
    Dim randomStart As Double
    Dim cumulativeTotal As Double
    Dim nextSamplePoint As Double
    Dim amountCol As Integer
    Dim markerCol As Integer
    Dim i As Long
    Dim itemsSelected As Long

    Set ws = ActiveSheet

    amountCol = Application.InputBox("Enter column number containing AMOUNTS:", "MUS Sampling", 3, Type:=1)
    If amountCol = 0 Then Exit Sub

    sampleSize = Application.InputBox("Enter desired sample size:", "MUS Sampling", 25, Type:=1)
    If sampleSize = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, amountCol).End(xlUp).Row
    markerCol = ws.UsedRange.Columns.Count + 1

    ' Calculate total population value
    totalPop = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol)))

    ' Calculate sampling interval
    samplingInterval = totalPop / sampleSize

    ' Random start between 0 and sampling interval
    Randomize
    randomStart = Rnd * samplingInterval

    ' Headers
    ws.Cells(1, markerCol).Value = "MUS Sample"
    ws.Cells(1, markerCol).Font.Bold = True
    ws.Cells(1, markerCol + 1).Value = "Cumulative $"
    ws.Cells(1, markerCol + 1).Font.Bold = True

    ' Select items
    cumulativeTotal = 0
    nextSamplePoint = randomStart
    itemsSelected = 0

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        cumulativeTotal = cumulativeTotal + Abs(ws.Cells(i, amountCol).Value)

        ' Record cumulative total
        ws.Cells(i, markerCol + 1).Value = cumulativeTotal

        ' Check if this item contains the sample point
        If cumulativeTotal >= nextSamplePoint And itemsSelected < sampleSize Then
            ws.Cells(i, markerCol).Value = "X"
            ws.Cells(i, markerCol).Interior.Color = RGB(255, 255, 0)
            itemsSelected = itemsSelected + 1
            nextSamplePoint = nextSamplePoint + samplingInterval
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "MUS SAMPLING COMPLETE" & vbCrLf & vbCrLf & _
           "Population Total: " & Format(totalPop, "$#,##0") & vbCrLf & _
           "Sampling Interval: " & Format(samplingInterval, "$#,##0") & vbCrLf & _
           "Random Start: " & Format(randomStart, "$#,##0") & vbCrLf & _
           "Items Selected: " & itemsSelected, vbInformation

End Sub
```

---

## Stratified Sample

Sample by population strata.

```vba
Sub StratifiedSample()
    '================================================
    ' Stratified Random Sampling
    ' Samples proportionally from different strata
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim amountCol As Integer
    Dim markerCol As Integer
    Dim i As Long
    Dim stratum As String
    Dim totalSample As Long

    ' Define strata thresholds
    Dim highThreshold As Double
    Dim medThreshold As Double
    Dim highSample As Long, medSample As Long, lowSample As Long
    Dim highCount As Long, medCount As Long, lowCount As Long

    Set ws = ActiveSheet

    amountCol = Application.InputBox("Enter AMOUNT column:", "Stratified Sample", 3, Type:=1)
    If amountCol = 0 Then Exit Sub

    highThreshold = Application.InputBox("Enter HIGH stratum threshold (test 100%):", "Stratified Sample", 10000, Type:=1)
    medThreshold = Application.InputBox("Enter MEDIUM stratum threshold:", "Stratified Sample", 1000, Type:=1)

    medSample = Application.InputBox("Sample size for MEDIUM stratum:", "Stratified Sample", 15, Type:=1)
    lowSample = Application.InputBox("Sample size for LOW stratum:", "Stratified Sample", 10, Type:=1)

    lastRow = ws.Cells(ws.Rows.Count, amountCol).End(xlUp).Row
    markerCol = ws.UsedRange.Columns.Count + 1

    ' Headers
    ws.Cells(1, markerCol).Value = "Stratum"
    ws.Cells(1, markerCol + 1).Value = "Sample"
    ws.Cells(1, markerCol).Font.Bold = True
    ws.Cells(1, markerCol + 1).Font.Bold = True

    Application.ScreenUpdating = False

    ' First pass: Assign strata and select all high items
    For i = 2 To lastRow
        If Abs(ws.Cells(i, amountCol).Value) >= highThreshold Then
            ws.Cells(i, markerCol).Value = "HIGH"
            ws.Cells(i, markerCol + 1).Value = "X"
            ws.Cells(i, markerCol + 1).Interior.Color = RGB(255, 199, 206)  ' Red
            highCount = highCount + 1
        ElseIf Abs(ws.Cells(i, amountCol).Value) >= medThreshold Then
            ws.Cells(i, markerCol).Value = "MEDIUM"
            medCount = medCount + 1
        Else
            ws.Cells(i, markerCol).Value = "LOW"
            lowCount = lowCount + 1
        End If
    Next i

    ' Random sample from medium stratum
    Dim medSelected As Long
    Randomize
    Do While medSelected < medSample And medSelected < medCount
        i = Int((lastRow - 1) * Rnd) + 2
        If ws.Cells(i, markerCol).Value = "MEDIUM" And ws.Cells(i, markerCol + 1).Value = "" Then
            ws.Cells(i, markerCol + 1).Value = "X"
            ws.Cells(i, markerCol + 1).Interior.Color = RGB(255, 235, 156)  ' Yellow
            medSelected = medSelected + 1
        End If
    Loop

    ' Random sample from low stratum
    Dim lowSelected As Long
    Do While lowSelected < lowSample And lowSelected < lowCount
        i = Int((lastRow - 1) * Rnd) + 2
        If ws.Cells(i, markerCol).Value = "LOW" And ws.Cells(i, markerCol + 1).Value = "" Then
            ws.Cells(i, markerCol + 1).Value = "X"
            ws.Cells(i, markerCol + 1).Interior.Color = RGB(198, 239, 206)  ' Green
            lowSelected = lowSelected + 1
        End If
    Loop

    Application.ScreenUpdating = True

    totalSample = highCount + medSelected + lowSelected

    MsgBox "STRATIFIED SAMPLE COMPLETE" & vbCrLf & vbCrLf & _
           "HIGH (100%): " & highCount & " items" & vbCrLf & _
           "MEDIUM: " & medSelected & " of " & medCount & vbCrLf & _
           "LOW: " & lowSelected & " of " & lowCount & vbCrLf & _
           "TOTAL SAMPLE: " & totalSample, vbInformation

End Sub
```

---

## Generate Confirmations

Create confirmation letters from listing.

```vba
Sub GenerateConfirmations()
    '================================================
    ' Generate Confirmation Letters
    ' Creates individual confirmation letters from AR/AP listing
    '================================================

    Dim wsData As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsConf As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim confNum As Long
    Dim companyName As String
    Dim address As String
    Dim balance As Double

    Set wsData = ActiveSheet

    ' Assumes columns: A=Company, B=Address, C=City/State/Zip, D=Balance
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    confNum = 0

    For i = 2 To lastRow
        ' Check if marked for confirmation
        If wsData.Cells(i, "E").Value = "X" Or wsData.Cells(i, "E").Value = "Confirm" Then

            confNum = confNum + 1
            companyName = wsData.Cells(i, "A").Value
            address = wsData.Cells(i, "B").Value & vbCrLf & wsData.Cells(i, "C").Value
            balance = wsData.Cells(i, "D").Value

            ' Create confirmation sheet
            Set wsConf = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsConf.Name = "Conf-" & confNum

            With wsConf
                .Range("A1").Value = "[YOUR FIRM NAME]"
                .Range("A2").Value = "[Address]"
                .Range("A3").Value = "[City, State ZIP]"

                .Range("A5").Value = Format(Date, "mmmm d, yyyy")

                .Range("A7").Value = companyName
                .Range("A8").Value = address

                .Range("A11").Value = "RE: Confirmation of Account Balance"
                .Range("A11").Font.Bold = True

                .Range("A13").Value = "Dear Sir or Madam:"

                .Range("A15").Value = "In connection with an audit of the financial statements of [CLIENT NAME], please confirm"
                .Range("A16").Value = "directly to our auditors the amount owed to (from) you as of [DATE]."

                .Range("A18").Value = "Our records indicate the following balance:"

                .Range("A20").Value = "Balance: " & Format(balance, "$#,##0.00")
                .Range("A20").Font.Bold = True
                .Range("A20").Font.Size = 12

                .Range("A23").Value = "Please indicate below whether this agrees with your records. If not, please provide details"
                .Range("A24").Value = "of any differences on the reverse side of this letter."

                .Range("A27").Value = "___ The balance shown is CORRECT"
                .Range("A28").Value = "___ The balance shown is INCORRECT (see reverse for details)"

                .Range("A31").Value = "Signature: _______________________________"
                .Range("A32").Value = "Title: _______________________________"
                .Range("A33").Value = "Date: _______________________________"

                .Range("A36").Value = "Please return this confirmation directly to:"
                .Range("A37").Value = "[AUDITOR ADDRESS]"

                ' Store confirmation number reference
                wsData.Cells(i, "F").Value = "Conf-" & confNum

            End With

        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Generated " & confNum & " confirmation letters.", vbInformation

End Sub
```

---

## Track Confirmation Responses

Monitor confirmation status and responses.

```vba
Sub TrackConfirmationResponses()
    '================================================
    ' Track Confirmation Responses
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sentCount As Long, receivedCount As Long, agreedCount As Long
    Dim exceptionCount As Long, noResponseCount As Long

    Set ws = ActiveSheet

    ' Assumes columns: A=Company, B=Balance, C=Sent Date, D=Response Date, E=Status, F=Difference
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Add Status dropdown if not exists
    On Error Resume Next
    ws.Range("E2:E" & lastRow).Validation.Delete
    ws.Range("E2:E" & lastRow).Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Formula1:="Sent,Received-Agrees,Received-Exception,No Response,Alternative Procedure"
    On Error GoTo 0

    ' Count statuses
    For i = 2 To lastRow
        Select Case ws.Cells(i, "E").Value
            Case "Sent"
                sentCount = sentCount + 1
            Case "Received-Agrees"
                receivedCount = receivedCount + 1
                agreedCount = agreedCount + 1
            Case "Received-Exception"
                receivedCount = receivedCount + 1
                exceptionCount = exceptionCount + 1
            Case "No Response"
                noResponseCount = noResponseCount + 1
        End Select
    Next i

    ' Conditional formatting
    ws.Range("E2:E" & lastRow).FormatConditions.Delete

    ws.Range("E2:E" & lastRow).FormatConditions.Add Type:=xlTextString, String:="Agrees", TextOperator:=xlContains
    ws.Range("E2:E" & lastRow).FormatConditions(1).Interior.Color = RGB(198, 239, 206)

    ws.Range("E2:E" & lastRow).FormatConditions.Add Type:=xlTextString, String:="Exception", TextOperator:=xlContains
    ws.Range("E2:E" & lastRow).FormatConditions(2).Interior.Color = RGB(255, 199, 206)

    ws.Range("E2:E" & lastRow).FormatConditions.Add Type:=xlTextString, String:="No Response", TextOperator:=xlContains
    ws.Range("E2:E" & lastRow).FormatConditions(3).Interior.Color = RGB(255, 235, 156)

    MsgBox "CONFIRMATION STATUS SUMMARY" & vbCrLf & vbCrLf & _
           "Total Confirmations: " & (lastRow - 1) & vbCrLf & _
           "Sent (awaiting): " & sentCount & vbCrLf & _
           "Received - Agrees: " & agreedCount & vbCrLf & _
           "Received - Exception: " & exceptionCount & vbCrLf & _
           "No Response: " & noResponseCount & vbCrLf & vbCrLf & _
           "Response Rate: " & Format(receivedCount / (lastRow - 1), "0%"), vbInformation

End Sub
```

---

## Test Footings

Verify mathematical accuracy of listings.

```vba
Sub TestFootings()
    '================================================
    ' Test Footings - Verify Column Totals
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim reportedTotal As Double
    Dim calculatedTotal As Double
    Dim difference As Double
    Dim results As String

    Set ws = ActiveSheet
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.UsedRange.Columns.Count

    results = "FOOTING TEST RESULTS" & vbCrLf & String(30, "-") & vbCrLf

    ' Test each numeric column
    For col = 1 To lastCol
        If IsNumeric(ws.Cells(lastRow, col).Value) Then
            reportedTotal = ws.Cells(lastRow, col).Value
            calculatedTotal = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, col), ws.Cells(lastRow - 1, col)))
            difference = reportedTotal - calculatedTotal

            If Abs(difference) > 0.01 Then
                results = results & "Column " & col & ": DIFFERENCE of " & Format(difference, "$#,##0.00") & vbCrLf
                ws.Cells(lastRow, col).Interior.Color = RGB(255, 199, 206)
            Else
                results = results & "Column " & col & ": Footed OK" & vbCrLf
                ws.Cells(lastRow, col).Interior.Color = RGB(198, 239, 206)
            End If
        End If
    Next col

    MsgBox results, vbInformation, "Footing Test"

End Sub

Sub CrossFootTest()
    '================================================
    ' Cross-Foot Test - Verify Row Totals
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim row As Long
    Dim reportedTotal As Double
    Dim calculatedTotal As Double
    Dim errorCount As Long

    Set ws = ActiveSheet
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.UsedRange.Columns.Count

    ' Assumes last column contains row totals
    For row = 2 To lastRow
        reportedTotal = ws.Cells(row, lastCol).Value
        calculatedTotal = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(row, 2), ws.Cells(row, lastCol - 1)))

        If Abs(reportedTotal - calculatedTotal) > 0.01 Then
            ws.Cells(row, lastCol).Interior.Color = RGB(255, 199, 206)
            errorCount = errorCount + 1
        Else
            ws.Cells(row, lastCol).Interior.Color = RGB(198, 239, 206)
        End If
    Next row

    If errorCount > 0 Then
        MsgBox errorCount & " cross-footing errors found (highlighted in red).", vbExclamation
    Else
        MsgBox "All rows cross-foot correctly!", vbInformation
    End If

End Sub
```

---

## Compare Two Listings

Match items between two populations (e.g., client vs auditor).

```vba
Sub CompareTwoListings()
    '================================================
    ' Compare Two Listings
    ' Matches items between two sheets/ranges
    '================================================

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsResults As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim keyCol1 As Integer, keyCol2 As Integer
    Dim amtCol1 As Integer, amtCol2 As Integer
    Dim i As Long, j As Long
    Dim resultRow As Long
    Dim matchFound As Boolean
    Dim matchCount As Long, unmatchCount As Long

    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws2 = ThisWorkbook.Sheets(2)

    ' Create results sheet
    Set wsResults = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResults.Name = "Comparison_Results"

    keyCol1 = 1: keyCol2 = 1
    amtCol1 = 2: amtCol2 = 2

    lastRow1 = ws1.Cells(ws1.Rows.Count, keyCol1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, keyCol2).End(xlUp).Row

    With wsResults
        .Range("A1").Value = "Key"
        .Range("B1").Value = "List 1 Amount"
        .Range("C1").Value = "List 2 Amount"
        .Range("D1").Value = "Difference"
        .Range("E1").Value = "Status"
        .Range("A1:E1").Font.Bold = True
    End With

    resultRow = 2

    Application.ScreenUpdating = False

    ' Compare List 1 to List 2
    For i = 2 To lastRow1
        matchFound = False

        For j = 2 To lastRow2
            If ws1.Cells(i, keyCol1).Value = ws2.Cells(j, keyCol2).Value Then
                matchFound = True

                wsResults.Cells(resultRow, 1).Value = ws1.Cells(i, keyCol1).Value
                wsResults.Cells(resultRow, 2).Value = ws1.Cells(i, amtCol1).Value
                wsResults.Cells(resultRow, 3).Value = ws2.Cells(j, amtCol2).Value
                wsResults.Cells(resultRow, 4).Formula = "=B" & resultRow & "-C" & resultRow

                If Abs(wsResults.Cells(resultRow, 4).Value) < 0.01 Then
                    wsResults.Cells(resultRow, 5).Value = "Match"
                    matchCount = matchCount + 1
                Else
                    wsResults.Cells(resultRow, 5).Value = "Difference"
                    wsResults.Cells(resultRow, 5).Interior.Color = RGB(255, 235, 156)
                End If

                resultRow = resultRow + 1
                Exit For
            End If
        Next j

        If Not matchFound Then
            wsResults.Cells(resultRow, 1).Value = ws1.Cells(i, keyCol1).Value
            wsResults.Cells(resultRow, 2).Value = ws1.Cells(i, amtCol1).Value
            wsResults.Cells(resultRow, 5).Value = "In List 1 Only"
            wsResults.Cells(resultRow, 5).Interior.Color = RGB(255, 199, 206)
            unmatchCount = unmatchCount + 1
            resultRow = resultRow + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "COMPARISON COMPLETE" & vbCrLf & vbCrLf & _
           "Matches: " & matchCount & vbCrLf & _
           "Unmatched: " & unmatchCount, vbInformation

End Sub
```

---

## Age Receivables

Create accounts receivable aging analysis.

```vba
Sub AgeReceivables()
    '================================================
    ' Age Receivables Analysis
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim invDate As Date
    Dim asOfDate As Date
    Dim daysPast As Long
    Dim dateCol As Integer, amtCol As Integer, ageCol As Integer

    Set ws = ActiveSheet

    dateCol = Application.InputBox("Invoice DATE column:", "AR Aging", 2, Type:=1)
    amtCol = Application.InputBox("AMOUNT column:", "AR Aging", 3, Type:=1)
    ageCol = ws.UsedRange.Columns.Count + 1

    asOfDate = Application.InputBox("As-of date:", "AR Aging", Date, Type:=1)
    If asOfDate = 0 Then asOfDate = Date

    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row

    ' Headers
    ws.Cells(1, ageCol).Value = "Days"
    ws.Cells(1, ageCol + 1).Value = "Aging Bucket"
    ws.Cells(1, ageCol).Font.Bold = True
    ws.Cells(1, ageCol + 1).Font.Bold = True

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        If IsDate(ws.Cells(i, dateCol).Value) Then
            invDate = ws.Cells(i, dateCol).Value
            daysPast = asOfDate - invDate

            ws.Cells(i, ageCol).Value = daysPast

            Select Case daysPast
                Case Is <= 30
                    ws.Cells(i, ageCol + 1).Value = "Current"
                    ws.Cells(i, ageCol + 1).Interior.Color = RGB(198, 239, 206)
                Case 31 To 60
                    ws.Cells(i, ageCol + 1).Value = "31-60"
                    ws.Cells(i, ageCol + 1).Interior.Color = RGB(255, 235, 156)
                Case 61 To 90
                    ws.Cells(i, ageCol + 1).Value = "61-90"
                    ws.Cells(i, ageCol + 1).Interior.Color = RGB(255, 199, 206)
                Case Is > 90
                    ws.Cells(i, ageCol + 1).Value = "Over 90"
                    ws.Cells(i, ageCol + 1).Interior.Color = RGB(255, 0, 0)
                    ws.Cells(i, ageCol + 1).Font.Color = RGB(255, 255, 255)
            End Select
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Aging complete!", vbInformation

End Sub
```

---

## Subsequent Receipts Testing

Test subsequent cash receipts for AR.

```vba
Sub SubsequentReceiptsTesting()
    '================================================
    ' Subsequent Receipts Testing
    ' Matches AR items to subsequent cash receipts
    '================================================

    Dim wsAR As Worksheet
    Dim wsReceipts As Worksheet
    Dim lastRowAR As Long, lastRowRec As Long
    Dim i As Long, j As Long
    Dim customer As String
    Dim amount As Double
    Dim collected As Double
    Dim statusCol As Integer

    Set wsAR = ThisWorkbook.Sheets("AR")
    Set wsReceipts = ThisWorkbook.Sheets("Receipts")

    lastRowAR = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row
    lastRowRec = wsReceipts.Cells(wsReceipts.Rows.Count, "A").End(xlUp).Row

    statusCol = wsAR.UsedRange.Columns.Count + 1

    wsAR.Cells(1, statusCol).Value = "Subsequent Collection"
    wsAR.Cells(1, statusCol).Font.Bold = True

    Application.ScreenUpdating = False

    ' For each AR item, search for subsequent receipt
    For i = 2 To lastRowAR
        customer = wsAR.Cells(i, 1).Value
        amount = wsAR.Cells(i, 3).Value
        collected = 0

        For j = 2 To lastRowRec
            If wsReceipts.Cells(j, 1).Value = customer Then
                collected = collected + wsReceipts.Cells(j, 2).Value
            End If
        Next j

        If collected >= amount Then
            wsAR.Cells(i, statusCol).Value = "Collected 100%"
            wsAR.Cells(i, statusCol).Interior.Color = RGB(198, 239, 206)
        ElseIf collected > 0 Then
            wsAR.Cells(i, statusCol).Value = "Partial: " & Format(collected / amount, "0%")
            wsAR.Cells(i, statusCol).Interior.Color = RGB(255, 235, 156)
        Else
            wsAR.Cells(i, statusCol).Value = "Not Collected"
            wsAR.Cells(i, statusCol).Interior.Color = RGB(255, 199, 206)
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Subsequent receipts testing complete!", vbInformation

End Sub
```

---

## Cutoff Testing

Test revenue and expense cutoff.

```vba
Sub CutoffTesting()
    '================================================
    ' Cutoff Testing
    ' Identifies transactions near period end
    '================================================

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim transDate As Date
    Dim periodEnd As Date
    Dim daysBefore As Long
    Dim daysAfter As Long
    Dim dateCol As Integer
    Dim flagCol As Integer
    Dim cutoffWindow As Long

    Set ws = ActiveSheet

    dateCol = Application.InputBox("Transaction DATE column:", "Cutoff Testing", 1, Type:=1)
    periodEnd = Application.InputBox("Period end date:", "Cutoff Testing", DateSerial(Year(Date), 12, 31), Type:=1)
    cutoffWindow = Application.InputBox("Days before/after to flag:", "Cutoff Testing", 5, Type:=1)

    lastRow = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row
    flagCol = ws.UsedRange.Columns.Count + 1

    ws.Cells(1, flagCol).Value = "Cutoff Flag"
    ws.Cells(1, flagCol).Font.Bold = True

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        If IsDate(ws.Cells(i, dateCol).Value) Then
            transDate = ws.Cells(i, dateCol).Value
            daysBefore = periodEnd - transDate
            daysAfter = transDate - periodEnd

            If daysBefore >= 0 And daysBefore <= cutoffWindow Then
                ws.Cells(i, flagCol).Value = "Before (" & daysBefore & " days)"
                ws.Cells(i, flagCol).Interior.Color = RGB(255, 235, 156)
            ElseIf daysAfter >= 0 And daysAfter <= cutoffWindow Then
                ws.Cells(i, flagCol).Value = "After (" & daysAfter & " days)"
                ws.Cells(i, flagCol).Interior.Color = RGB(255, 199, 206)
            End If
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Cutoff testing complete!" & vbCrLf & "Items within " & cutoffWindow & " days of " & Format(periodEnd, "mm/dd/yyyy") & " are flagged.", vbInformation

End Sub
```

---

## Audit Sampling Quick Reference

| Sample Type | When to Use |
|-------------|-------------|
| **Random** | General testing, no dollar weighting needed |
| **MUS/PPS** | Substantive testing where larger items more important |
| **Stratified** | Different risk levels in population |
| **Haphazard** | Quick judgmental selection |
| **Block/Systematic** | Testing specific periods |

---

[⬅️ Back to VBA Macros](../README.md) | [🏠 Back to Home](../../README.md)
