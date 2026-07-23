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

    Private Const RibbonXml As String = "
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""tabPCMS"" label=""Sarwar PCMS"">
        <group id=""grpPane"" label=""Panel"">
          <toggleButton id=""btnTogglePane""
                        label=""Task Pane""
                        size=""large""
                        imageMso=""TaskPane""
                        onAction=""OnTogglePane""
                        getPressed=""GetPaneVisible""
                        screentip=""Open/Close Sarwar PCMS""
                        supertip=""Show or hide the Sarwar PCMS task pane for data entry and review."" />
        </group>
        <group id=""grpActions"" label=""Data"">
          <button id=""btnRefresh"" label=""Refresh Sheet Data"" size=""large"" imageMso=""RefreshAll""
                  onAction=""OnRefreshData""
                  supertip=""Recalculate PCMS custom fields on the active sheet."" />
          <button id=""btnValidate"" label=""Validate Schedule"" size=""large"" imageMso=""ReviewCheckContactNames""
                  onAction=""OnValidate""
                  supertip=""Run DCMA-style schedule/EVM checks on the active sheet (dates, float, logic, variances, % complete)."" />
          <button id=""btnExport"" label=""Export Findings"" size=""large"" imageMso=""ExportExcel""
                  onAction=""OnExportValidation""
                  supertip=""Write all validation findings to a new PCMS_Validation sheet."" />
        </group>
        <group id=""grpHistory"" label=""History (SQLite)"">
          <button id=""btnSaveSnapshot"" label=""Save Snapshot"" size=""large"" imageMso=""FileSave""
                  onAction=""OnSaveSnapshot""
                  supertip=""Save the active sheet's schedule/EVM data and validation findings to the local Sarwar PCMS database."" />
          <button id=""btnLoadSnapshot"" label=""Load Last Snapshot"" size=""large"" imageMso=""FileOpen""
                  onAction=""OnLoadLastSnapshot""
                  supertip=""Load the most recently saved snapshot into a new sheet for comparison."" />
        </group>
        <group id=""grpAnalysis"" label=""Analysis"">
          <button id=""btnDashboard"" label=""Build EVM Dashboard"" size=""large"" imageMso=""ChartInsertColumn""
                  onAction=""OnBuildDashboard""
                  supertip=""Generate an S-curve chart (cumulative PV/EV/AC) and headline SPI/CPI on a new sheet."" />
          <button id=""btnHighlight"" label=""Highlight Critical Path"" size=""large"" imageMso=""FillColorPicker""
                  onAction=""OnHighlightCritical""
                  supertip=""Color rows red (critical, float&lt;=0) or yellow (near-critical, float&lt;=5 days) based on Total Float."" />
        </group>
        <group id=""grpAbout"" label=""About"">
          <button id=""btnAbout"" label=""About"" size=""large"" imageMso=""Info""
                  onAction=""OnAbout"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>"

    Public Overrides Function GetCustomUI(RibbonID As String) As String
        Return RibbonXml
    End Function

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
