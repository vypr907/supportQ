VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} authFrm 
   Caption         =   "Who are you?"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3105
   OleObjectBlob   =   "authFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "authFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub userForm_Initialize()
   setUsersRange
   'load data
   Dim item As Variant
   For Each item In usersRng'dataSht.Range("users")
      With Me.userCboBx
         .AddItem item.Value
      End With
   Next item
End Sub

Sub submitBtn_Click()
   'authorizer(Me.userCboBx.ListIndex, Me.pinEntryBx)
   Hide Me
End Sub