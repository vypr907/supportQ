VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addUserFrm 
   Caption         =   "Add User"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "addUserFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addUserFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub userForm_Initialize()
    Me.pinBx.Value = Int(2 + Rnd * (9999 - 1111 + 1))
End Sub

Private Sub cancelBtn_Click()

    'closing the add user form
    On Error Resume Next
    
    'verify form is empty
    'setting variables for message boxes
    Dim msg, style, title, response
    msg = "Are you sure you want to discard?"
    style = vbYesNo + vbExclamation
    title = "Caution!"

    'checking for valid input in the text boxes
    '---FIRST NAME-----------------------------
    If fnameBx.Value = "" Then
        lnameBx.SetFocus
    Else
        response = MsgBox(msg, style, title)
        If response = vbYes Then
            Me.lnameBx.Value = ""
            Unload Me
        Else
            Exit Sub
        End If
    End If
    '---LAST NAME------------------------------
    If lnameBx.Value = "" Then
        'cardNumBx.SetFocus
    Else
        response = MsgBox(msg, style, title)
        If response = vbYes Then
            Unload Me
        Else
            lnameBx.SetFocus
            Exit Sub
        End If
    End If

    Unload Me
End Sub

Private Sub saveBtn_Click()
    addUser
    'clear form
    With Me
        .fnameBx.Value = ""
        .miBx.Value = ""
        .lnameBx.Value = ""
        .pinBx.Value = Int(2 + Rnd * (9999 - 1111 + 1))
    End With
    MsgBox "done."
End Sub

