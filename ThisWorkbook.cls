Option Explicit

Private Sub Workbook_Open()
'auto run macros to hide sheets and pop up GUI
    start
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

