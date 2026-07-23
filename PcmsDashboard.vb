Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Visual analysis features: EVM S-curve dashboard and critical-path
''' highlighting, generated directly from the active sheet's data.
''' </summary>
Public Module PcmsDashboard

    Private Function FindColumn(ws As Worksheet, headerName As String) As Integer
        Dim usedRange As Range = ws.UsedRange
        Dim lastCol As Integer = usedRange.Columns.Count
        For c As Integer = 1 To lastCol
            Dim v As Object = CType(ws.Cells(1, c), Range).Value2
            If v IsNot Nothing AndAlso v.ToString().Trim().ToLower() = headerName.ToLower() Then Return c
        Next
        Return -1
    End Function

    ''' <summary>
    ''' Builds a "PCMS_Dashboard" sheet with cumulative PV/EV/AC totals and an
    ''' S-curve chart, plus headline SPI/CPI figures, from the active sheet's data.
    ''' </summary>
    Public Sub BuildDashboard(app As Microsoft.Office.Interop.Excel.Application, ws As Worksheet)
        Dim colPV = FindColumn(ws, "PV")
        Dim colEV = FindColumn(ws, "EV")
        Dim colAC = FindColumn(ws, "AC")
        Dim colFinish = FindColumn(ws, "Finish Date")

        If colPV < 0 OrElse colEV < 0 OrElse colAC < 0 Then
            Throw New Exception("Dashboard needs PV, EV, and AC columns on the active sheet.")
        End If

        Dim lastRow As Integer = ws.UsedRange.Rows.Count
        Dim wb As Workbook = app.ActiveWorkbook

        On Error Resume Next
        Dim oldSheet As Worksheet = CType(wb.Sheets("PCMS_Dashboard"), Worksheet)
        If oldSheet IsNot Nothing Then
            app.DisplayAlerts = False
            oldSheet.Delete()
            app.DisplayAlerts = True
        End If
        On Error GoTo 0

        Dim dash As Worksheet = CType(wb.Sheets.Add(), Worksheet)
        dash.Name = "PCMS_Dashboard"

        dash.Cells(1, 1) = "Sarwar PCMS - EVM Dashboard"
        CType(dash.Cells(1, 1), Range).Font.Bold = True
        CType(dash.Cells(1, 1), Range).Font.Size = 14

        dash.Cells(3, 1) = "Row"
        dash.Cells(3, 2) = "Finish Date"
        dash.Cells(3, 3) = "Cumulative PV"
        dash.Cells(3, 4) = "Cumulative EV"
        dash.Cells(3, 5) = "Cumulative AC"
        CType(dash.Range("A3:E3"), Range).Font.Bold = True

        Dim cumPV As Double = 0, cumEV As Double = 0, cumAC As Double = 0
        Dim outRow As Integer = 4

        For r As Integer = 2 To lastRow
            Dim pv = SafeDbl(CType(ws.Cells(r, colPV), Range).Value2)
            Dim ev = SafeDbl(CType(ws.Cells(r, colEV), Range).Value2)
            Dim ac = SafeDbl(CType(ws.Cells(r, colAC), Range).Value2)
            cumPV += pv
            cumEV += ev
            cumAC += ac

            dash.Cells(outRow, 1) = r
            If colFinish > 0 Then
                dash.Cells(outRow, 2) = CType(ws.Cells(r, colFinish), Range).Value2
            End If
            dash.Cells(outRow, 3) = cumPV
            dash.Cells(outRow, 4) = cumEV
            dash.Cells(outRow, 5) = cumAC
            outRow += 1
        Next

        ' Headline KPIs
        dash.Cells(1, 7) = "SPI"
        dash.Cells(2, 7) = If(cumPV = 0, "n/a", Math.Round(cumEV / cumPV, 2))
        dash.Cells(1, 8) = "CPI"
        dash.Cells(2, 8) = If(cumAC = 0, "n/a", Math.Round(cumEV / cumAC, 2))
        CType(dash.Range("G1:H1"), Range).Font.Bold = True

        ' S-curve chart
        Dim dataRange As Range = dash.Range(dash.Cells(3, 2), dash.Cells(outRow - 1, 5))
        Dim chartObj As ChartObject = CType(dash.ChartObjects(), ChartObjects).Add(400, 20, 500, 300)
        With chartObj.Chart
            .SetSourceData(dataRange)
            .ChartType = XlChartType.xlLine
            .HasTitle = True
            .ChartTitle.Text = "EVM S-Curve (Cumulative PV / EV / AC)"
        End With

        CType(dash.Columns("A:H"), Range).AutoFit()
        dash.Activate()
    End Sub

    Private Function SafeDbl(v As Object) As Double
        If v Is Nothing Then Return 0
        If TypeOf v Is Double Then Return CDbl(v)
        Return 0
    End Function

    ''' <summary>
    ''' Color-highlights rows on the active sheet based on Total Float:
    ''' red = critical (float &lt;= 0), yellow = near-critical (0 &lt; float &lt;= 5 days).
    ''' Non-destructive: only touches fill color, no data changes.
    ''' </summary>
    Public Sub HighlightCriticalPath(ws As Worksheet)
        Dim colFloat = FindColumn(ws, "Total Float")
        If colFloat < 0 Then
            Throw New Exception("Highlighting needs a 'Total Float' column on the active sheet.")
        End If

        Dim lastRow As Integer = ws.UsedRange.Rows.Count
        Dim lastCol As Integer = ws.UsedRange.Columns.Count

        For r As Integer = 2 To lastRow
            Dim cellVal = CType(ws.Cells(r, colFloat), Range).Value2
            Dim rowRange As Range = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))

            If TypeOf cellVal Is Double Then
                Dim tf As Double = CDbl(cellVal)
                If tf <= 0 Then
                    rowRange.Interior.Color = RGB(255, 199, 206) ' red
                ElseIf tf <= 5 Then
                    rowRange.Interior.Color = RGB(255, 235, 156) ' yellow
                Else
                    rowRange.Interior.ColorIndex = 0 ' none
                End If
            End If
        Next
    End Sub

End Module
