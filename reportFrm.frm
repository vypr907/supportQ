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
   'load users
   Dim item As Variant
   For Each item In dataSht.Range("users")
      With Me.techCboBx2
         .AddItem item.Value
      End With
   Next item
   
   'load log entries
   Dim i,d,k
   i = 0
   d = 0
   k = 0
   lastLogRow = logSht.Cells(Rows.Count,1).End(xlUp).Offset(1,0).row
   With Me
      .logLB.ColumnCount = 13
      .logLB.ColumnWidths = "15,70,60,50,35,35,40,60,120,150,25,65,65"
      '.logLB.RowSource = "Log!A2:M" & lastLogRow
      For rw = 2 to lastLogRow
         .logLB.AddItem
         For i = 1 to 12
            .logLB.List(k,i-1) = logSht.Cells(rw,i)
         Next i
         k = k + 1
      Next rw
      .totRecordsBx = .logLB.ListCount - 1
   End With
End Sub

Sub searchBtn_Click()
   logSearch Me.techCboBx2.Value
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = vbFormControlMenu Then
      Cancel = True
      Me.Hide
   End If
End Sub
