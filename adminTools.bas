Attribute VB_Name = "adminTools"
'Code for adding/removing users


Public Sub addUser()
    MsgBox "addin' this shit!"
    Dim nextRow As Integer
    nextRow = lastUserRow
    
    Dim first as Variant
    Dim mi as Variant
    Dim last as Variant
    dim pin as Integer

    With addUsrScreen
        first = .fnameBx.Value
        mi = .miBx.Value
        last = .lnameBx.Value
        pin = .pinBx.Value
    End With

    With dataSht
        .Cells(nextRow, 7).Value = first
        .Cells(nextRow, 8).Value = mi
        .Cells(nextRow, 9).Value = last
        .Cells(nextRow, 10).Value = pin
    End With
End Sub

Public Function removeUser(row as Integer)
    MsgBox "BEGONE!"
    With dataSht
        .Cells(row, 7).Value = ""
        .Cells(row, 8).Value = ""
        .Cells(row, 9).Value = ""
        .Cells(row, 10).Value = ""
    End With
End Function
