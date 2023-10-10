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

Sub userForm_Initialize()
    Dim userRow As Integer
    userRow = dataSht.Cells(Rows.Count, 5).End(xlUp).Offset(1, 0).row
    lastUserRow = userRow

    Me.userLB.ColumnCount = 4
    Me.userLB.ColumnWidths = "60,20,60,40"

    Me.userLB.RowSource = "listData!E2:I" & userRow

End Sub

Private Sub addUserBtn_Click()
    addUsrScreen.pinBx.Value = pinGen
    addUsrScreen.Show
End Sub


Sub userLB_AfterUpdate()
    '***** Verify that a row is selected first
    If userLB.ListIndex > -1 And userLB.Selected(userLB.ListIndex) Then
        '***** Use the data
    End If
End Sub
Private Sub rmUserBtn_Click()

    Dim selectedRow As Integer
    '***** Verify that a row is selected first
    If userLB.ListIndex = -1 Then
        Exit Sub
    End If
    If userLB.ListIndex > -1 And userLB.Selected(userLB.ListIndex) Then
        '***** Use the data
        selectedRow = userLB.ListIndex + 2
        removeUser (selectedRow)
    End If
End Sub
