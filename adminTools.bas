Attribute VB_Name = "adminTools"
'Code for adding/removing users


Public Sub addUser()
    MsgBox "addin' this shit!"
    Dim nextRow As Integer
    nextRow = lastUserRow
    
    Dim first As String
    Dim mi As String
    Dim last As String
    Dim pin As Integer
    Dim init As String

    With addUsrScreen
        first = .fnameBx.Value
        mi = .miBx.Value
        last = .lnameBx.Value
        pin = .pinBx.Value
        init = LEFT(.fnameBx,1)&LEFT(.miBx,1)&LEFT(.lnameBx,1)
    End With

    With dataSht
        .Cells(nextRow, 7).Value = first
        .Cells(nextRow, 8).Value = mi
        .Cells(nextRow, 9).Value = last
        .Cells(nextRow, 10).Value = pin
        .Cells(nextRow, 11).Value = init
    End With
End Sub

Public Function removeUser(row As Integer)
    MsgBox "BEGONE!"
    With dataSht
        .Cells(row, 7).Value = ""
        .Cells(row, 8).Value = ""
        .Cells(row, 9).Value = ""
        .Cells(row, 10).Value = ""
        .Cells(row, 11).Value = ""
    End With
End Function

Sub setUsersRange()
    'Set usersRng = dataSht.Range("G2:J" & lastUserRow)
    Set usersRng = dataSht.Cells("G2",lastUserRow)
End Sub