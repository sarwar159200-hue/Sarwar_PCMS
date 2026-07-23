Imports ExcelDna.Integration.CustomUI

''' <summary>
''' Manages the single Sarwar PCMS custom task pane instance.
''' </summary>
Public Module TaskPaneManager

    Private paneControl As CustomTaskPane

    Public Sub SetVisible(visible As Boolean)
        If visible Then
            If paneControl Is Nothing Then
                paneControl = CustomTaskPaneFactory.CreateCustomTaskPane(GetType(PcmsTaskPaneControl), "Sarwar PCMS")
                paneControl.Width = 320
            End If
            paneControl.Visible = True
        Else
            If paneControl IsNot Nothing Then
                paneControl.Visible = False
            End If
        End If
    End Sub

End Module
