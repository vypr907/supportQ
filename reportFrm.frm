VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} reportFrm 
   Caption         =   "REPORTS"
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16755
   OleObjectBlob   =   "reportFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "reportFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub userForm_Initialize()
   lastLogRow = logSht.Cells(Rows.Count,1).End(xlUp).Offset(1,0).row
   With Me
      .logLB.ColumnCount = 13
      .logLB.ColumnWidths = "15,70,60,50,35,35,40,60,120,150,25,65,65"
      .logLB.RowSource = "Log!A2:M" & lastLogRow
      .totRecordsBx = .logLB.ListCount - 1
   End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = vbFormControlMenu Then
      Cancel = True
      Me.Hide
   End If
End Sub
