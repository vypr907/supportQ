VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} nameFrm
   Caption         =   "WHATS YA NAME?"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3105
   OleObjectBlob   =   "nameFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "nameFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
