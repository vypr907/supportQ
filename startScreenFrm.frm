Private Sub startKioskBtn_Click()
    signInFrm.Show
End Sub

Private Sub qViewBtn_Click()
    queueView.Show
End Sub

Private Sub reportsBtn_Click()
    Application.DisplayAlerts = False
    If temp Is Nothing Then
        MsgBox "Starting reports engine...", , "Initializing"
        tempXL
        Set tmpSearch = temp.Sheets("Search")
        Set tmpLog = temp.Sheets("Log")
    Else
        MsgBox "Refreshing data...", , "Reports Engine"
    End If
    
    wb.Activate
    Application.ScreenUpdating = False
    temp.Windows(1).Visible = True
    reportsRun
    reportView.Show vbModeless
    Application.DisplayAlerts = True
End Sub

Private Sub setupBtn_Click()
    userMaintFrm.Show
End Sub

Private Sub exitBtn_Click()
    If Not temp Is Nothing Then
        temp.Close SaveChanges:=False
    End If
    gameOver
End Sub

Private Sub aboutBtn_Click()
    aboutFrm.Show
End Sub
