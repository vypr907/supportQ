VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPwEntry 
   Caption         =   "Administrators Only"
   ClientHeight    =   1140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmPwEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPwEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PwSubmitBtn_Click()
    
    If adminUnlock(Me.pwTxtBx.Value, attempts) = True Then 'user entered correct pw
    'unlock the form
        unlockDat
    ElseIf attempts > 0 Then 'password no correct
        MsgBox "Nope. " & attempts & " attempts remaining.", vbOKOnly + vbExclamation, "#rekt"
        frmPwEntry.pwTxtBx.Value = ""
        frmPwEntry.pwTxtBx.SetFocus
        Exit Sub
    Else 'user has run out of attempts
        MsgBox "You have run out of attempts. Goodbye."
        frmStartMenuUser.lockBtn.Visible = True
        Unload frmPwEntry
        Exit Sub
    End If
    Unload frmPwEntry
End Sub
