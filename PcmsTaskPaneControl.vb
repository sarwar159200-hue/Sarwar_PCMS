Imports System.Linq
Imports System.Windows.Forms
Imports System.Drawing
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

Public Class PcmsTaskPaneControl
    Inherits UserControl

    Private lblSheet As System.Windows.Forms.Label
    Private btnRefresh As System.Windows.Forms.Button
    Private btnValidate As System.Windows.Forms.Button
    Private btnExport As System.Windows.Forms.Button
    Private grid As DataGridView
    Private lblSummary As System.Windows.Forms.Label

    Public Sub New()
        Me.Dock = DockStyle.Fill

        lblSheet = New System.Windows.Forms.Label() With {
            .Text = "Active sheet: (none)",
            .Dock = DockStyle.Top,
            .Height = 24,
            .Font = New System.Drawing.Font("Segoe UI", 9, FontStyle.Bold)
        }

        btnRefresh = New System.Windows.Forms.Button() With {.Text = "Refresh Sheet Data", .Dock = DockStyle.Top, .Height = 30}
        AddHandler btnRefresh.Click, AddressOf OnRefreshClick

        btnValidate = New System.Windows.Forms.Button() With {.Text = "Validate Schedule", .Dock = DockStyle.Top, .Height = 30}
        AddHandler btnValidate.Click, AddressOf OnValidateClick

        btnExport = New System.Windows.Forms.Button() With {.Text = "Export Findings to Sheet", .Dock = DockStyle.Top, .Height = 30}
        AddHandler btnExport.Click, AddressOf OnExportClick

        lblSummary = New System.Windows.Forms.Label() With {.Text = "", .Dock = DockStyle.Top, .Height = 20}

        grid = New DataGridView() With {
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }
        grid.Columns.Add("Row", "Row")
        grid.Columns.Add("Check", "Check")
        grid.Columns.Add("Severity", "Severity")
        grid.Columns.Add("Message", "Message")

        Me.Controls.Add(grid)
        Me.Controls.Add(lblSummary)
        Me.Controls.Add(btnExport)
        Me.Controls.Add(btnValidate)
        Me.Controls.Add(btnRefresh)
        Me.Controls.Add(lblSheet)

        AddHandler Me.VisibleChanged, AddressOf OnVisibleChanged
    End Sub

    Private Sub OnVisibleChanged(sender As Object, e As EventArgs)
        If Me.Visible Then RefreshSheetLabel()
    End Sub

    Private Sub RefreshSheetLabel()
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            lblSheet.Text = "Active sheet: " & If(ws IsNot Nothing, ws.Name, "(none)")
        Catch
            lblSheet.Text = "Active sheet: (none)"
        End Try
    End Sub

    Private Sub OnRefreshClick(sender As Object, e As EventArgs)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            PcmsEngine.RefreshSheet(ws)
            lblSummary.Text = "Sheet recalculated: " & ws.Name
            RefreshSheetLabel()
        Catch ex As Exception
            lblSummary.Text = "Error: " & ex.Message
        End Try
    End Sub

    Private Sub OnValidateClick(sender As Object, e As EventArgs)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            Dim issues = PcmsEngine.ValidateSheetDetailed(ws)

            grid.Rows.Clear()
            For Each i In issues
                Dim idx = grid.Rows.Add(i.Row, i.Check, i.Severity, i.Message)
                If i.Severity = "Error" Then
                    grid.Rows(idx).DefaultCellStyle.BackColor = Color.MistyRose
                Else
                    grid.Rows(idx).DefaultCellStyle.BackColor = Color.LightYellow
                End If
            Next

            If issues.Count = 0 Then
                lblSummary.Text = "No issues found."
            Else
                Dim errCount = issues.Where(Function(i) i.Severity = "Error").Count()
                Dim warnCount = issues.Where(Function(i) i.Severity = "Warning").Count()
                lblSummary.Text = $"{issues.Count} finding(s): {errCount} error(s), {warnCount} warning(s)."
            End If
        Catch ex As Exception
            lblSummary.Text = "Error: " & ex.Message
        End Try
    End Sub

    Private Sub OnExportClick(sender As Object, e As EventArgs)
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
            Dim ws As Worksheet = app.ActiveSheet
            PcmsEngine.ExportValidationToSheet(app, ws)
            lblSummary.Text = "Findings exported to PCMS_Validation sheet."
        Catch ex As Exception
            lblSummary.Text = "Error: " & ex.Message
        End Try
    End Sub

End Class
