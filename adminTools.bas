'Code for adding/removing users


Public Sub addUser()
    MsgBox "addin' this shit!", , "boop"
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
        init = Left(.fnameBx, 1) & Left(.miBx, 1) & Left(.lnameBx, 1)
    End With

    With dataSht
        .Cells(nextRow, 7).Value = first
        .Cells(nextRow, 8).Value = mi
        .Cells(nextRow, 9).Value = last
        .Cells(nextRow, 10).Value = pin
        .Cells(nextRow, 11).Value = init
    End With
End Sub

Public Function pinGen()
    pinGen = Int(2 + Rnd * (9999 - 1111 + 1))
End Function

Public Function removeUser(row As Integer)
    MsgBox "BEGONE!"
    With dataSht
        .Cells(row, 5).Value = ""
        .Cells(row, 6).Value = ""
        .Cells(row, 7).Value = ""
        .Cells(row, 8).Value = ""
        .Cells(row, 9).Value = ""
    End With
End Function

Sub setUsersRange()
    Set usersRng = dataSht.Cells("G2", lastUserRow)
End Sub
