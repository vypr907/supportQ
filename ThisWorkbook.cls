Option Explicit

Private Sub Workbook_Open()
'auto run macros to hide sheets and pop up GUI
    testCode = Application.InputBox("Run in test mode? (1=true, 0=false)" & vbCr _
    & "1 = True, 0 = False", "Startup", Type:=4)

    Dim start As startScreenFrm
    Set start = New startScreenFrm
    
    If testCode = True Then
        MsgBox "hi, I'm in test mode!"
        startScreenFrm.Show vbModeless
        Application.ScreenUpdating = True
        openSesame
    Else
        MsgBox "hi, I'm in regular mode!"
        start.Show vbModal
        byeFelicia
        Application.ScreenUpdating = False
    End If
End Sub


Private Sub Workbook_NewSheet(ByVal Sh As Object)
    'prevent users from creating new sheets
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        
        ActiveSheet.Delete
        
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub

