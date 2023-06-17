VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} signInFrm 
   Caption         =   "Sign In"
   ClientHeight    =   10770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9225.001
   OleObjectBlob   =   "signInFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "signInFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
   Set wb = Workbooks("SupportQ_DEV")
   Set queueSht = wb.Sheets("Queue")
   Set logSht = wb.Sheets("Log")
   Set listSht = wb.Sheets("listData")

   'loops to fill comboboxes
   'For Each reasonItem In listSht.Range("reasonCode")
   '   With Me.reasonSelect
   '      .AddItem reasonItem.Value
   '   End With
   'Next reasonItem

   'For Each branchItem In listSht.Range("branchOfSvc")
   '   With Me.branchSelect
   '      .AddItem branchItem.Value
   '   End With
   'Next branchItem

   'setting to first value "Select"
   'Me.branchSelect.ListIndex = 0
   'Me.reasonSelect.ListIndex = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   'making it so users cannot accidentally close the sign in form.
   Dim response

   If CloseMode = vbFormControlMenu Then
      Cancel = True
      'response = MsgBox("Please enter")
      Call comparison
      'If InputBox("Enter password to close: ") = "609eacoms" Then
      If password = "609eacoms" Then
         Cancel = False
      Else
         MsgBox("Authentication Failure")
      End If
   End If
End Sub