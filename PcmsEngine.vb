Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Custom Excel functions (formulas) for project controls / EVM calculations.
''' These are plain UDFs - usable as =PCMS_SPI(...) directly in any cell.
''' </summary>
Public Module PcmsFunctions

    <ExcelFunction(Description:="Schedule Performance Index = EV / PV")>
    Public Function PCMS_SPI(<ExcelArgument(Description:="Earned Value")> ev As Double,
                              <ExcelArgument(Description:="Planned Value")> pv As Double) As Object
        If pv = 0 Then Return ExcelError.ExcelErrorDiv0
        Return ev / pv
    End Function

    <ExcelFunction(Description:="Cost Performance Index = EV / AC")>
    Public Function PCMS_CPI(<ExcelArgument(Description:="Earned Value")> ev As Double,
                              <ExcelArgument(Description:="Actual Cost")> ac As Double) As Object
        If ac = 0 Then Return ExcelError.ExcelErrorDiv0
        Return ev / ac
    End Function

    <ExcelFunction(Description:="Schedule Variance = EV - PV")>
    Public Function PCMS_SV(ev As Double, pv As Double) As Double
        Return ev - pv
    End Function

    <ExcelFunction(Description:="Cost Variance = EV - AC")>
    Public Function PCMS_CV(ev As Double, ac As Double) As Double
        Return ev - ac
    End Function

    <ExcelFunction(Description:="Estimate at Completion = BAC / CPI")>
    Public Function PCMS_EAC(bac As Double, cpi As Double) As Object
        If cpi = 0 Then Return ExcelError.ExcelErrorDiv0
        Return bac / cpi
    End Function

    <ExcelFunction(Description:="Total Float in working days between two dates (Early Finish, Late Finish)")>
    Public Function PCMS_TOTAL_FLOAT(earlyFinish As DateTime, lateFinish As DateTime) As Double
        Return (lateFinish - earlyFinish).TotalDays
    End Function

    <ExcelFunction(Description:="Percent complete = Actual Duration / Planned Duration, capped at 100%")>
    Public Function PCMS_PCT_COMPLETE(actualDuration As Double, plannedDuration As Double) As Object
        If plannedDuration = 0 Then Return ExcelError.ExcelErrorDiv0
        Return Math.Min(1.0, actualDuration / plannedDuration)
    End Function

    <ExcelFunction(Description:="Variance at Completion = BAC - EAC")>
    Public Function PCMS_VAC(bac As Double, eac As Double) As Double
        Return bac - eac
    End Function

    <ExcelFunction(Description:="Estimate to Complete = EAC - AC")>
    Public Function PCMS_ETC(eac As Double, ac As Double) As Double
        Return eac - ac
    End Function

    <ExcelFunction(Description:="To-Complete Performance Index = (BAC - EV) / (BAC - AC)")>
    Public Function PCMS_TCPI(bac As Double, ev As Double, ac As Double) As Object
        Dim denom As Double = bac - ac
        If denom = 0 Then Return ExcelError.ExcelErrorDiv0
        Return (bac - ev) / denom
    End Function

    <ExcelFunction(Description:="EAC using the composite CPI*SPI method = AC + (BAC-EV)/(CPI*SPI)")>
    Public Function PCMS_EAC_COMPOSITE(bac As Double, ev As Double, ac As Double, cpi As Double, spi As Double) As Object
        Dim denom As Double = cpi * spi
        If denom = 0 Then Return ExcelError.ExcelErrorDiv0
        Return ac + (bac - ev) / denom
    End Function

    <ExcelFunction(Description:="Forecast finish date = data date + remaining duration / SPI")>
    Public Function PCMS_FORECAST_FINISH(dataDate As DateTime, remainingDuration As Double, spi As Double) As DateTime
        Dim factor As Double = If(spi > 0, 1.0 / spi, 1.0)
        Return dataDate.AddDays(remainingDuration * factor)
    End Function

    <ExcelFunction(Description:="Free float = Early Start of successor - Early Finish of predecessor")>
    Public Function PCMS_FREE_FLOAT(successorEarlyStart As DateTime, predecessorEarlyFinish As DateTime) As Double
        Return (successorEarlyStart - predecessorEarlyFinish).TotalDays
    End Function

    <ExcelFunction(Description:="Flags TRUE if Total Float is at or below the critical threshold (default 0)")>
    Public Function PCMS_IS_CRITICAL(totalFloat As Double, Optional thresholdDays As Double = 0) As Boolean
        Return totalFloat <= thresholdDays
    End Function

End Module

''' <summary>
''' Sheet-level operations triggered from the ribbon. Reads/writes the active
''' workbook's cells directly - no external database, per requirements.
''' </summary>
Public Module PcmsEngine

    ' Convention: sheet has a header row 1 with named columns.
    ' This helper finds a column index by header text (case-insensitive).
    Private Function FindColumn(ws As Worksheet, headerName As String) As Integer
        Dim usedRange As Range = ws.UsedRange
        Dim lastCol As Integer = usedRange.Columns.Count
        For c As Integer = 1 To lastCol
            Dim cellVal As Object = CType(ws.Cells(1, c), Range).Value2
            If cellVal IsNot Nothing AndAlso cellVal.ToString().Trim().ToLower() = headerName.ToLower() Then
                Return c
            End If
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Recalculates PCMS-related formula cells on the active sheet.
    ''' Placeholder for whatever "refresh" logic you want (e.g. re-pulling
    ''' latest EV/PV/AC into a summary block, recomputing rollups, etc.)
    ''' </summary>
    Public Sub RefreshSheet(ws As Worksheet)
        If ws Is Nothing Then Return
        ws.Calculate()
        ' Extend here: e.g. recompute rollup totals, refresh a dashboard block, etc.
    End Sub

    ''' <summary>
    ''' One validation finding: row, check name, severity, message.
    ''' </summary>
    Public Class PcmsIssue
        Public Property Row As Integer
        Public Property Check As String
        Public Property Severity As String ' "Error" / "Warning"
        Public Property Message As String
    End Class

    ''' <summary>
    ''' DCMA-style multi-check validation across a schedule/EVM sheet.
    ''' Runs whichever checks apply based on which columns actually exist -
    ''' missing columns are silently skipped rather than causing errors.
    ''' </summary>
    Public Function ValidateSheetDetailed(ws As Worksheet) As List(Of PcmsIssue)
        Dim results As New List(Of PcmsIssue)
        If ws Is Nothing Then Return results

        Dim usedRange As Range = ws.UsedRange
        Dim lastRow As Integer = usedRange.Rows.Count

        Dim colStart As Integer = FindColumn(ws, "Start Date")
        Dim colFinish As Integer = FindColumn(ws, "Finish Date")
        Dim colEV As Integer = FindColumn(ws, "EV")
        Dim colPV As Integer = FindColumn(ws, "PV")
        Dim colAC As Integer = FindColumn(ws, "AC")
        Dim colBAC As Integer = FindColumn(ws, "BAC")
        Dim colTotalFloat As Integer = FindColumn(ws, "Total Float")
        Dim colPredecessor As Integer = FindColumn(ws, "Predecessor")
        Dim colSuccessor As Integer = FindColumn(ws, "Successor")
        Dim colDuration As Integer = FindColumn(ws, "Duration")
        Dim colPctComplete As Integer = FindColumn(ws, "% Complete")

        For r As Integer = 2 To lastRow
            Dim GetVal As Func(Of Integer, Object) = Function(c) If(c > 0, CType(ws.Cells(r, c), Range).Value2, Nothing)

            ' Check 1: Missing/invalid dates
            If colStart > 0 AndAlso colFinish > 0 Then
                Dim sVal = GetVal(colStart)
                Dim fVal = GetVal(colFinish)
                If sVal Is Nothing OrElse fVal Is Nothing Then
                    results.Add(New PcmsIssue With {.Row = r, .Check = "Dates", .Severity = "Error", .Message = "Missing Start/Finish date."})
                ElseIf TypeOf sVal Is Double AndAlso TypeOf fVal Is Double Then
                    If CDbl(fVal) < CDbl(sVal) Then
                        results.Add(New PcmsIssue With {.Row = r, .Check = "Dates", .Severity = "Error", .Message = "Finish date before Start date."})
                    End If
                End If
            End If

            ' Check 2: Negative or extreme cost variance
            If colEV > 0 AndAlso colAC > 0 Then
                Dim evVal = GetVal(colEV)
                Dim acVal = GetVal(colAC)
                If TypeOf evVal Is Double AndAlso TypeOf acVal Is Double Then
                    Dim cv As Double = CDbl(evVal) - CDbl(acVal)
                    If CDbl(acVal) > 0 AndAlso Math.Abs(cv) / CDbl(acVal) > 0.5 Then
                        results.Add(New PcmsIssue With {.Row = r, .Check = "Cost Variance", .Severity = "Warning", .Message = $"CV = {cv:N0} (>50% of AC) - verify entry."})
                    End If
                End If
            End If

            ' Check 3: EV exceeds BAC (over 100% earned)
            If colEV > 0 AndAlso colBAC > 0 Then
                Dim evVal = GetVal(colEV)
                Dim bacVal = GetVal(colBAC)
                If TypeOf evVal Is Double AndAlso TypeOf bacVal Is Double Then
                    If CDbl(evVal) > CDbl(bacVal) Then
                        results.Add(New PcmsIssue With {.Row = r, .Check = "EV vs BAC", .Severity = "Error", .Message = "Earned Value exceeds Budget at Completion."})
                    End If
                End If
            End If

            ' Check 4: Negative or missing total float (DCMA-style)
            If colTotalFloat > 0 Then
                Dim tfVal = GetVal(colTotalFloat)
                If TypeOf tfVal Is Double AndAlso CDbl(tfVal) < 0 Then
                    results.Add(New PcmsIssue With {.Row = r, .Check = "Total Float", .Severity = "Error", .Message = $"Negative total float ({CDbl(tfVal):N1} days)."})
                End If
            End If

            ' Check 5: Missing logic (no predecessor and no successor - "dangling" activity)
            If colPredecessor > 0 AndAlso colSuccessor > 0 Then
                Dim predVal = GetVal(colPredecessor)
                Dim succVal = GetVal(colSuccessor)
                Dim predEmpty As Boolean = (predVal Is Nothing OrElse predVal.ToString().Trim() = "")
                Dim succEmpty As Boolean = (succVal Is Nothing OrElse succVal.ToString().Trim() = "")
                If predEmpty AndAlso succEmpty Then
                    results.Add(New PcmsIssue With {.Row = r, .Check = "Logic", .Severity = "Warning", .Message = "No predecessor or successor (open-ended activity)."})
                End If
            End If

            ' Check 6: Duration vs % Complete sanity (100% complete but duration remaining implied elsewhere is handled upstream)
            If colDuration > 0 Then
                Dim durVal = GetVal(colDuration)
                If TypeOf durVal Is Double AndAlso CDbl(durVal) <= 0 Then
                    results.Add(New PcmsIssue With {.Row = r, .Check = "Duration", .Severity = "Warning", .Message = "Zero or negative duration."})
                End If
            End If

            If colPctComplete > 0 Then
                Dim pctVal = GetVal(colPctComplete)
                If TypeOf pctVal Is Double AndAlso (CDbl(pctVal) < 0 OrElse CDbl(pctVal) > 1.0001) Then
                    results.Add(New PcmsIssue With {.Row = r, .Check = "% Complete", .Severity = "Error", .Message = $"% Complete out of range ({CDbl(pctVal):P0})."})
                End If
            End If
        Next

        Return results
    End Function

    ''' <summary>
    ''' Back-compat plain-text wrapper around ValidateSheetDetailed.
    ''' </summary>
    Public Function ValidateSheet(ws As Worksheet) As String
        Dim issues = ValidateSheetDetailed(ws)
        If issues.Count = 0 Then Return ""
        Dim sb As New System.Text.StringBuilder()
        For Each i In issues
            sb.AppendLine($"Row {i.Row} [{i.Severity}] {i.Check}: {i.Message}")
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Writes validation results to a new "PCMS_Validation" sheet (replacing
    ''' any prior one), so findings can be filtered/sorted/shared like normal data.
    ''' </summary>
    Public Sub ExportValidationToSheet(app As Microsoft.Office.Interop.Excel.Application, ws As Worksheet)
        Dim issues = ValidateSheetDetailed(ws)
        Dim wb As Workbook = app.ActiveWorkbook

        ' Remove existing output sheet if present
        On Error Resume Next
        Dim oldSheet As Worksheet = CType(wb.Sheets("PCMS_Validation"), Worksheet)
        If oldSheet IsNot Nothing Then
            app.DisplayAlerts = False
            oldSheet.Delete()
            app.DisplayAlerts = True
        End If
        On Error GoTo 0

        Dim outSheet As Worksheet = CType(wb.Sheets.Add(After:=ws), Worksheet)
        outSheet.Name = "PCMS_Validation"

        outSheet.Cells(1, 1) = "Row"
        outSheet.Cells(1, 2) = "Check"
        outSheet.Cells(1, 3) = "Severity"
        outSheet.Cells(1, 4) = "Message"
        CType(outSheet.Range("A1:D1"), Range).Font.Bold = True

        Dim r As Integer = 2
        For Each issue In issues
            outSheet.Cells(r, 1) = issue.Row
            outSheet.Cells(r, 2) = issue.Check
            outSheet.Cells(r, 3) = issue.Severity
            outSheet.Cells(r, 4) = issue.Message
            r += 1
        Next

        CType(outSheet.Columns("A:D"), Range).AutoFit()
    End Sub

End Module
