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
    If firstNameBx.Value = "" Then
        lastNameBx.SetFocus
    Else
        response = MsgBox(msg, style, title)
        If response = vbYes Then
            Me.lastNameBx.Value = ""
            Me.cardNumBx.Value = ""
            Unload Me
        Else
            Exit Sub
        End If
    End If
    '---LAST NAME------------------------------
    If lastNameBx.Value = "" Then
        cardNumBx.SetFocus
    Else
        response = MsgBox(msg, style, title)
        If response = vbYes Then
            Unload Me
        Else
            'lastNameBx.SetFocus
            Exit Sub
        End If
    End If

    Unload Me
End Sub

Private Sub saveBtn_Click()
    addUser
End Sub
