Option Explicit

Private Sub userForm_Initialize()

   'loops to fill comboboxes
   Dim item As Variant
   For Each item In dataSht.Range("reasonCode")
      With Me.reasonCboBx
         .AddItem item.Value
      End With
   Next item

   For Each item In dataSht.Range("branchOfSvc")
      With Me.branchCboBx
         .AddItem item.Value
      End With
   Next item

   For Each item In dataSht.Range("rank")
      With Me.rankCboBx
         .AddItem item.Value
      End With
   Next item

   'setting to first value "Select"
   Me.branchCboBx.ListIndex = 0
   Me.reasonCboBx.ListIndex = 0
   Me.rankCboBx.ListIndex = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   'making it so users cannot accidentally close the sign in form.
   Dim response

   If CloseMode = vbFormControlMenu Then
      Cancel = True
      Call comparison
      If password = "609eacoms" Then
         Cancel = False
      Else
         MsgBox "Invalid Password", , "Authentication Failure"
         Cancel = True
      End If
   End If
End Sub

Private Sub submitBtn_Click()
    'test for blank
    With Me
        If .surnameBx = "" Or _
            .fnameBx = "" Or _
            .rankCboBx.ListIndex < 1 Or _
            .branchCboBx.ListIndex < 1 Or _
            .phoneBx = "" Or _
            .reasonCboBx.ListIndex < 1 Then
            MsgBox "Please complete all fields!", , "Missing Info"
            Exit Sub
        End If
    End With
    queueAdd
    save
    clearForm
    surnameBx.SetFocus
    popUp "Thank you! Your Reference number is " & refID & ", and you are position " & _
        queuePos & " in the queue!", "Submission Received", 5
        
End Sub

