Imports System.Runtime.InteropServices
Imports System.Linq
Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports Microsoft.Office.Interop.Excel

<ComVisible(True)>
Public Class RibbonController
    Inherits ExcelRibbon

    Private ribbonUI As IRibbonUI
    Private paneVisible As Boolean = False

    Public Sub OnRibbonLoad(ribbon As IRibbonUI)
        ribbonUI = ribbon
    End Sub

    ' ---- Task pane toggle ----
    Public Sub OnTogglePane(control As IRibbonControl, pressed As Boolean)
        paneVisible = pressed
        TaskPaneManager.SetVisible(pressed)
        ribbonUI?.InvalidateControl("btnTogglePane")
    End Sub

    Public Function GetPaneVisible(control As IRibbonControl) As Boolean
        Return paneVisible
    End Function

    ' ---- Data actions: operate directly on the active workbook/sheet ----
    Public Sub OnRefreshData(control As IRibbonControl)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            PcmsEngine.RefreshSheet(ws)
        Catch ex As Exception
            ExcelDna.Logging.LogDisplay.WriteLine("PCMS Refresh error: " & ex.Message)
        End Try
    End Sub

    Public Sub OnValidate(control As IRibbonControl)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            Dim issues = PcmsEngine.ValidateSheetDetailed(ws)
            If issues.Count = 0 Then
                MsgBox("No issues found.", MsgBoxStyle.Information, "Sarwar PCMS - Validation")
            Else
                Dim errCount = issues.Where(Function(i) i.Severity = "Error").Count()
                Dim warnCount = issues.Where(Function(i) i.Severity = "Warning").Count()
                MsgBox($"{issues.Count} finding(s): {errCount} error(s), {warnCount} warning(s)." & vbCrLf &
                       "Use ""Export Findings"" for the full row-by-row list.",
                       MsgBoxStyle.Exclamation, "Sarwar PCMS - Validation")
            End If
        Catch ex As Exception
            ExcelDna.Logging.LogDisplay.WriteLine("PCMS Validate error: " & ex.Message)
        End Try
    End Sub

    Public Sub OnExportValidation(control As IRibbonControl)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            PcmsEngine.ExportValidationToSheet(app, ws)
        Catch ex As Exception
            ExcelDna.Logging.LogDisplay.WriteLine("PCMS Export error: " & ex.Message)
        End Try
    End Sub

    ' ---- SQLite history (mirrors Elwaha's local-DB architecture) ----
    Public Sub OnSaveSnapshot(control As IRibbonControl)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            Dim wbName As String = app.ActiveWorkbook.Name
            Dim id As Long = PcmsDatabase.SaveSnapshot(ws, wbName)
            MsgBox($"Snapshot #{id} saved to the local Sarwar PCMS database.", MsgBoxStyle.Information, "Sarwar PCMS")
        Catch ex As Exception
            MsgBox("Save failed: " & ex.Message, MsgBoxStyle.Critical, "Sarwar PCMS")
        End Try
    End Sub

    Public Sub OnLoadLastSnapshot(control As IRibbonControl)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim snapshots = PcmsDatabase.ListSnapshots()
            If snapshots.Count = 0 Then
                MsgBox("No snapshots saved yet.", MsgBoxStyle.Information, "Sarwar PCMS")
                Return
            End If
            PcmsDatabase.LoadSnapshotToSheet(app, snapshots(0).Id)
        Catch ex As Exception
            MsgBox("Load failed: " & ex.Message, MsgBoxStyle.Critical, "Sarwar PCMS")
        End Try
    End Sub

    ' ---- Visual analysis ----
    Public Sub OnBuildDashboard(control As IRibbonControl)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            PcmsDashboard.BuildDashboard(app, ws)
        Catch ex As Exception
            MsgBox("Dashboard failed: " & ex.Message, MsgBoxStyle.Critical, "Sarwar PCMS")
        End Try
    End Sub

    Public Sub OnHighlightCritical(control As IRibbonControl)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            PcmsDashboard.HighlightCriticalPath(ws)
        Catch ex As Exception
            MsgBox("Highlight failed: " & ex.Message, MsgBoxStyle.Critical, "Sarwar PCMS")
        End Try
    End Sub

    Public Sub OnAbout(control As IRibbonControl)
        MsgBox("Sarwar PCMS" & vbCrLf & "Project Controls Management System" & vbCrLf & "Built on Excel-DNA", MsgBoxStyle.Information, "About")
    End Sub

End Class
