Attribute VB_Name = "Module1"
Sub Main()
If App.PrevInstance = True Then
    MsgBox "Application is already running,", vbOKOnly + vbCritical, "Invalid Action"
    End
Else
    frmMain.Show
End If
End Sub
