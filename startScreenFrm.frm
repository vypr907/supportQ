VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} startScreenFrm 
   Caption         =   "Start Up"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3495
   OleObjectBlob   =   "startScreenFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "startScreenFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub startKioskBtn_Click()
    'Dim f As Object
    'Set f = signInFrm
    'f.Show
    signInFrm.Show
End Sub

Private Sub qViewBtn_Click()

    'Call authorizer
    'Dim test As Boolean
    'test = authorizer()

    'If authorizer() = True Then'correct PIN entered
    '    queueView.Show
    'Else
    '    MsgBox "Access Denied.",vbOk,"No."
    '    Exit Sub
    'End If
    queueView.Show
End Sub

Private Sub reportsBtn_Click()
    'create new xl instance
    Call tempXL
    'copy data to temp xl instance
    'show userform
    reportView.Show vbModeless
    'close/delete temp xl instance
End Sub

Private Sub setupBtn_Click()
    userMaintFrm.Show
End Sub