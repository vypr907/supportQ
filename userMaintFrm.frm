VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userMaintFrm 
   Caption         =   "User Maintenance"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "userMaintFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userMaintFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addUserBtn_Click()
    Dim u As addUserFrm
    Set u = New addUserFrm
    u.Show
End Sub
