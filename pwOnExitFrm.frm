VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pwOnExitFrm 
   Caption         =   "Technicians Only"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3000
   OleObjectBlob   =   "pwOnExitFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pwOnExitFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0

Dim response As Variant

Private Sub initialize()
   password = ""
End Sub

Private Sub okBtn_Click()
   password = Me.pwBox.Value
   'response = password
   Unload Me
End Sub

Private Sub cancelBtn_Click()
   Unload Me
End Sub
'update
