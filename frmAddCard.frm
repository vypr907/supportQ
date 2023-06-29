VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddCard 
   Caption         =   "Add Customer"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5085
   OleObjectBlob   =   "frmAddCard.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnAddCard_Click()
'=========================================================

    'setting target for data entry
    Dim ws As Worksheet
    Set ws = Worksheets("dataStore")
    
    'establishing variables
    Dim lRow As Long 'variable for row
    Dim card2 As Long
    Dim first As String
    Dim last As String
    
    card2 = Me.cardNumBx.Value
    first = Me.firstNameBx.Value
    last = Me.lastNameBx.Value
    
    'verify user/card number doesn't already exist-----
    
    
    '--------------------------------------------------

    'setting variables for message boxes
    Dim msg, style, title, response
    msg = "Please ensure ALL fields are filled"
    style = vbOKOnly + vbExclamation
    title = "Oops!"

    'checking for valid input in the text boxes
    If firstNameBx.Value = "" Then
        response = MsgBox(msg, style, title)
        firstNameBx.SetFocus
        Exit Sub
    End If

    If lastNameBx.Value = "" Then
        response = MsgBox(msg, style, title)
        lastNameBx.SetFocus
        Exit Sub
    End If

    If cardNumBx.Value = "" Then
        response = MsgBox(msg, style, title)
        cardNumBx.SetFocus
        Exit Sub
    End If

    'find first empty row in database
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    'copy data to the database
    With ws
        .Cells(lRow, 1).Value = Me.cardNumBx.Value
        .Cells(lRow, 2).Value = Me.firstNameBx.Value
        .Cells(lRow, 3).Value = Me.lastNameBx.Value
    End With
    
    'increment total number of customers
    numCustomers = numCustomers + 1
    
    'enable Add Transactions
    frmstartMenu.addTransBtn.Enabled = True
    
    'clear out form for next entry
    Me.firstNameBx.Value = ""
    Me.lastNameBx.Value = ""
    Me.cardNumBx.Value = ""
    Me.firstNameBx.SetFocus
'==========================================================
End Sub

Private Sub btnClose_Click()
'closing the add customer form
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
    '---CARD NUMBER----------------------------
    If cardNumBx.Value = "" Then
        Unload Me
    Else
        response = MsgBox(msg, style, title)
        If response = vbYes Then
            Unload Me
        Else
            'cardNumBx.SetFocus
            Exit Sub
        End If
    End If
    
    
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    'start userform centered inside Excel screen-----------
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    '------------------------------------------------------
    
    If fromTrans = True Then
        Me.cardNumBx.Value = card
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'when user clicks the "X"
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "Plz use close button.", vbOKOnly + vbExclamation, "Hahahahahaha No."
    End If
End Sub
