Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = vbFormControlMenu Then
      Cancel = True
      MsgBox "No.", , "NO QUACK"
   End If
End Sub

Private Sub userForm_Initialize()
    namer = "USER_GEN_REPORT_" & Format(Now(), "YYYY-MM-DD")
    With Me
        .reportNameBx.Value = namer
        .reportNameBx.SelStart = 0
        .reportNameBx.SelLength = Len(Me.reportNameBx)
    End With
End Sub

Private Sub quackBtn_Click()
    namer = Me.reportNameBx.Value
    
    If validName(namer) Then
        namer = Me.reportNameBx.Value
        Unload Me
    Else
        MsgBox "Bad quack my friend, try again!", , "Invalid Filename"
        Me.reportNameBx.SetFocus
        Me.reportNameBx.SelStart = 0
        Me.reportNameBx.SelLength = Len(Me.reportNameBx)
        Exit Sub
    End If
End Sub
