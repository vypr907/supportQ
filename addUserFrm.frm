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
    Me.pinBx.Value = pinGen
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
            With Me
                .fnameBx = ""
                .miBx = ""
                .lnameBx = ""
                .Hide
            End With
        Else
            Exit Sub
        End If
    End If
    '---LAST NAME------------------------------
    If lnameBx.Value = "" Then
        lnameBx.SetFocus
    Else
        response = MsgBox(msg, style, title)
        If response = vbYes Then
            With Me
                .fnameBx = ""
                .miBx = ""
                .lnameBx = ""
                .Hide
            End With
        Else
            lnameBx.SetFocus
            Exit Sub
        End If
    End If

    Me.Hide
End Sub

Private Sub saveBtn_Click()
    addUser
    'clear form
    With Me
        .fnameBx.Value = ""
        .miBx.Value = ""
        .lnameBx.Value = ""
        .pinBx.Value = pinGen
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
