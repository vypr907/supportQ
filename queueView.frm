VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} queueView 
   Caption         =   "QueueView"
   ClientHeight    =   8400.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945.001
   OleObjectBlob   =   "queueView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "queueView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lastRow As Integer

'MAIN QUEUE
Private Sub userform_Initialize()
   lastRow = qSht.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row

   Me.custQLB.ColumnCount = 9
   '                          #,time,surname,first,branch,shop,phone,reason,notes
   Me.custQLB.ColumnWidths = "15,0,50,40,35,30,60,120,80"
   Me.custQLB.RowSource = "Queue!A2:I" & lastRow

   qSizeBx = custQLB.ListCount - 1
   timeBx = Now

   'load users
   Dim item As Variant
   For Each item In dataSht.Range("users")
      With Me.techCboBx
         .AddItem item.Value
      End With
   Next item
   
End Sub

Private Sub refreshBtnOne_Click()
   lastRow = qSht.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row

   Me.custQLB.ColumnCount = 9
   '                          #,time,surname,first,branch,shop,phone,reason,notes
   Me.custQLB.ColumnWidths = "15,0,50,40,35,30,60,120,80"
   Me.custQLB.RowSource = "Queue!A2:I" & lastRow

   qSizeBx = custQLB.ListCount - 1
   timeBx = Now
End Sub

'TODO: Close button
Private Sub closeBtnOne_Click()
   Me.Hide
End Sub
'TODO: Take button


Private Sub MultiPage1_Click(ByVal Index As Long)
	If Index = 1 Then  'for example, if 2nd.page clicked (first page start from Index=0)
       'your code here
       MsgBox "Hi!"
	end if
end sub
'MY QUEUE
'TODO: my queue initialization
   'TODO: load my queue

'TODO: load data into boxes on select
'TODO: Save button
'TODO: RESOLVE button

