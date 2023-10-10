Option Explicit
Option Base 0

Dim response As Variant

Private Sub initialize()
   password = ""
End Sub

Private Sub okBtn_Click()
   password = Me.pwBox.Value
   Unload Me
End Sub

Private Sub cancelBtn_Click()
   Unload Me
End Sub
