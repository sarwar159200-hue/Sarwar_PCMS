Imports System.Data.SQLite
Imports System.IO
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' Local SQLite persistence for Sarwar PCMS. Mirrors the architecture of the
''' Elwaha tool (local per-user SQLite DB alongside the add-in) so schedule
''' data, EVM snapshots, and validation history can be saved/loaded/tracked
''' over time - not just live in the open workbook.
''' </summary>
Public Module PcmsDatabase

    Private ReadOnly DbFolder As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Sarwar_PCMS")

    Private ReadOnly DbPath As String = Path.Combine(DbFolder, "pcms_data.sqlite")

    Private ReadOnly ConnString As String = $"Data Source={DbPath};Version=3;"

    ''' <summary>
    ''' Creates the DB file and schema if they don't already exist. Safe to
    ''' call on every add-in startup.
    ''' </summary>
    Public Sub EnsureDatabase()
        If Not Directory.Exists(DbFolder) Then Directory.CreateDirectory(DbFolder)
        If Not File.Exists(DbPath) Then SQLiteConnection.CreateFile(DbPath)

        Using conn As New SQLiteConnection(ConnString)
            conn.Open()
            Using cmd As New SQLiteCommand(conn)
                cmd.CommandText = "
                    CREATE TABLE IF NOT EXISTS Activities (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        SnapshotId INTEGER NOT NULL,
                        RowIndex INTEGER,
                        ActivityName TEXT,
                        StartDate TEXT,
                        FinishDate TEXT,
                        Duration REAL,
                        BAC REAL,
                        EV REAL,
                        PV REAL,
                        AC REAL,
                        TotalFloat REAL,
                        PctComplete REAL
                    );

                    CREATE TABLE IF NOT EXISTS Snapshots (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        SheetName TEXT,
                        WorkbookName TEXT,
                        SavedAtUtc TEXT
                    );

                    CREATE TABLE IF NOT EXISTS ValidationHistory (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        SnapshotId INTEGER NOT NULL,
                        RowIndex INTEGER,
                        CheckName TEXT,
                        Severity TEXT,
                        Message TEXT
                    );
                "
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

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
    ''' Saves the active sheet's schedule/EVM rows plus current validation
    ''' findings as one timestamped snapshot in the local SQLite DB.
    ''' </summary>
    Public Function SaveSnapshot(ws As Worksheet, wbName As String) As Long
        EnsureDatabase()

        Dim colName = FindColumn(ws, "Activity Name")
        Dim colStart = FindColumn(ws, "Start Date")
        Dim colFinish = FindColumn(ws, "Finish Date")
        Dim colDuration = FindColumn(ws, "Duration")
        Dim colBAC = FindColumn(ws, "BAC")
        Dim colEV = FindColumn(ws, "EV")
        Dim colPV = FindColumn(ws, "PV")
        Dim colAC = FindColumn(ws, "AC")
        Dim colFloat = FindColumn(ws, "Total Float")
        Dim colPct = FindColumn(ws, "% Complete")

        Dim lastRow As Integer = ws.UsedRange.Rows.Count
        Dim snapshotId As Long

        Using conn As New SQLiteConnection(ConnString)
            conn.Open()
            Using tx = conn.BeginTransaction()
                Using cmd As New SQLiteCommand(
                    "INSERT INTO Snapshots (SheetName, WorkbookName, SavedAtUtc) VALUES (@sheet, @wb, @ts);
                     SELECT last_insert_rowid();", conn, tx)
                    cmd.Parameters.AddWithValue("@sheet", ws.Name)
                    cmd.Parameters.AddWithValue("@wb", wbName)
                    cmd.Parameters.AddWithValue("@ts", DateTime.UtcNow.ToString("o"))
                    snapshotId = CLng(cmd.ExecuteScalar())
                End Using

                For r As Integer = 2 To lastRow
                    Dim GetVal As Func(Of Integer, Object) = Function(c) If(c > 0, CType(ws.Cells(r, c), Range).Value2, Nothing)

                    Using cmd As New SQLiteCommand(
                        "INSERT INTO Activities
                         (SnapshotId, RowIndex, ActivityName, StartDate, FinishDate, Duration, BAC, EV, PV, AC, TotalFloat, PctComplete)
                         VALUES (@sid, @row, @name, @start, @finish, @dur, @bac, @ev, @pv, @ac, @float, @pct)", conn, tx)
                        cmd.Parameters.AddWithValue("@sid", snapshotId)
                        cmd.Parameters.AddWithValue("@row", r)
                        cmd.Parameters.AddWithValue("@name", If(GetVal(colName)?.ToString(), ""))
                        cmd.Parameters.AddWithValue("@start", If(GetVal(colStart)?.ToString(), ""))
                        cmd.Parameters.AddWithValue("@finish", If(GetVal(colFinish)?.ToString(), ""))
                        cmd.Parameters.AddWithValue("@dur", SafeDouble(GetVal(colDuration)))
                        cmd.Parameters.AddWithValue("@bac", SafeDouble(GetVal(colBAC)))
                        cmd.Parameters.AddWithValue("@ev", SafeDouble(GetVal(colEV)))
                        cmd.Parameters.AddWithValue("@pv", SafeDouble(GetVal(colPV)))
                        cmd.Parameters.AddWithValue("@ac", SafeDouble(GetVal(colAC)))
                        cmd.Parameters.AddWithValue("@float", SafeDouble(GetVal(colFloat)))
                        cmd.Parameters.AddWithValue("@pct", SafeDouble(GetVal(colPct)))
                        cmd.ExecuteNonQuery()
                    End Using
                Next

                Dim issues = PcmsEngine.ValidateSheetDetailed(ws)
                For Each issue In issues
                    Using cmd As New SQLiteCommand(
                        "INSERT INTO ValidationHistory (SnapshotId, RowIndex, CheckName, Severity, Message)
                         VALUES (@sid, @row, @chk, @sev, @msg)", conn, tx)
                        cmd.Parameters.AddWithValue("@sid", snapshotId)
                        cmd.Parameters.AddWithValue("@row", issue.Row)
                        cmd.Parameters.AddWithValue("@chk", issue.Check)
                        cmd.Parameters.AddWithValue("@sev", issue.Severity)
                        cmd.Parameters.AddWithValue("@msg", issue.Message)
                        cmd.ExecuteNonQuery()
                    End Using
                Next

                tx.Commit()
            End Using
        End Using

        Return snapshotId
    End Function

    Private Function SafeDouble(v As Object) As Object
        If v Is Nothing Then Return DBNull.Value
        If TypeOf v Is Double Then Return v
        Return DBNull.Value
    End Function

    ''' <summary>
    ''' Returns a summary list of saved snapshots (Id, Sheet, Workbook, SavedAtUtc)
    ''' for display in the task pane history list.
    ''' </summary>
    Public Function ListSnapshots() As List(Of (Id As Long, Sheet As String, Workbook As String, SavedAt As String))
        EnsureDatabase()
        Dim results As New List(Of (Id As Long, Sheet As String, Workbook As String, SavedAt As String))
        Using conn As New SQLiteConnection(ConnString)
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT Id, SheetName, WorkbookName, SavedAtUtc FROM Snapshots ORDER BY Id DESC LIMIT 50", conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        results.Add((reader.GetInt64(0), reader.GetString(1), reader.GetString(2), reader.GetString(3)))
                    End While
                End Using
            End Using
        End Using
        Return results
    End Function

    ''' <summary>
    ''' Loads a previously saved snapshot's activity rows into a brand-new
    ''' sheet named "PCMS_Snapshot_&lt;id&gt;" for comparison against the live data.
    ''' </summary>
    Public Sub LoadSnapshotToSheet(app As Microsoft.Office.Interop.Excel.Application, snapshotId As Long)
        EnsureDatabase()
        Dim wb As Workbook = app.ActiveWorkbook

        Dim sheetName As String = $"PCMS_Snapshot_{snapshotId}"
        On Error Resume Next
        Dim existing As Worksheet = CType(wb.Sheets(sheetName), Worksheet)
        If existing IsNot Nothing Then
            app.DisplayAlerts = False
            existing.Delete()
            app.DisplayAlerts = True
        End If
        On Error GoTo 0

        Dim outSheet As Worksheet = CType(wb.Sheets.Add(), Worksheet)
        outSheet.Name = sheetName

        Dim headers = {"RowIndex", "ActivityName", "StartDate", "FinishDate", "Duration", "BAC", "EV", "PV", "AC", "TotalFloat", "PctComplete"}
        For c As Integer = 0 To headers.Length - 1
            outSheet.Cells(1, c + 1) = headers(c)
        Next
        CType(outSheet.Range(outSheet.Cells(1, 1), outSheet.Cells(1, headers.Length)), Range).Font.Bold = True

        Using conn As New SQLiteConnection(ConnString)
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT RowIndex, ActivityName, StartDate, FinishDate, Duration, BAC, EV, PV, AC, TotalFloat, PctComplete
                 FROM Activities WHERE SnapshotId = @sid ORDER BY RowIndex", conn)
                cmd.Parameters.AddWithValue("@sid", snapshotId)
                Using reader = cmd.ExecuteReader()
                    Dim r As Integer = 2
                    While reader.Read()
                        For c As Integer = 0 To headers.Length - 1
                            If Not reader.IsDBNull(c) Then
                                outSheet.Cells(r, c + 1) = reader.GetValue(c)
                            End If
                        Next
                        r += 1
                    End While
                End Using
            End Using
        End Using

        CType(outSheet.Columns("A:K"), Range).AutoFit()
    End Sub

End Module
