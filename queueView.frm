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

'***MAIN QUEUE
Private Sub userform_Initialize()
   MultiPage1.Value = 0
   lastRow = qSht.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row

   Me.custQLB.ColumnCount = 10
   '                          #,time,surname,first,branch,shop,phone,reason,notes
   Me.custQLB.ColumnWidths = "15,0,50,40,35,25,30,60,120,80"
   Me.custQLB.RowSource = "Queue!A2:J" & lastRow

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
   refresh(1)
End Sub
Private Sub refreshBtnTwo_Click()
   refresh(2)
End Sub

Private Sub closeBtnOne_Click()
   Me.Hide
End Sub
'TODO: Take button
Private Sub takeBtn_Click()

   '**THIS CODE WORKS, BUT WANT TO FORCE "FIFO"----------------------------
   Dim selectedRow As Integer
   Dim selectedUser As String
   '*Verify that a row is selected first
   If custQLB.ListIndex > -1 And custQLB.Selected(custQLB.ListIndex) Then
      '*Use the data
      selectedRow = custQLB.ListIndex + 2
      selectedUser = techCboBx.Value
      refID = custQLB.List(custQLB.ListIndex,0)
      takeEntry selectedRow,refID,selectedUser
   End If
   '**----------------------------------------------------------------------
End Sub

'***MY QUEUE
Private Sub MultiPage1_Click(ByVal Index As Long)
   Dim rw as Integer
   
	If Index = 1 Then  'for example, if 2nd.page clicked (first page start from Index=0)
      Me.myQLB.ColumnCount = 10
      '                     #,time,surname,first,branch,shop,phone,reason,notes
      Me.myQLB.ColumnWidths = "15,0,50,40,35,25,30,60,120,80"
      refresh(2)
	end if
end sub

Private Sub myQLB_Change()
   'do the shit.
   If Me.myQLB.ListIndex = -1 Then
      'empty the textboxes
      With Me
         .sNameBx = ""
         .fNameBx = ""
         .rankBx = ""
         .branchBx = ""
         .shopBx = ""
         .phoneBx = ""
         .reasonBx = ""
         .notesBx = ""
      End With
   Else 'Me.myQLB.ListIndex > -1 OR Me.myQLB.Selected(myQLB.ListIndex) Then
      With Me
         .sNameBx = .myQLB.List(myQLB.ListIndex,2) 
         .fNameBx = .myQLB.List(myQLB.ListIndex,3)
         .rankBx = .myQLB.List(myQLB.ListIndex,4)
         .branchBx = .myQLB.List(myQLB.ListIndex,5)
         .shopBx = .myQLB.List(myQLB.ListIndex,6)
         .phoneBx = .myQLB.List(myQLB.ListIndex,7)
         .reasonBx = .myQLB.List(myQLB.ListIndex,8)
         .notesBx = .myQLB.List(myQLB.ListIndex,9)
      End With
   End If
End Sub
'TODO: Save button
Private Sub saveBtn_Click()
   
   '***** Verify that a row is selected first
   If myQLB.ListIndex > -1 And myQLB.Selected(myQLB.ListIndex) Then
      'MsgBox userLB.List(userLB.ListIndex, 1) & ":" & userLB.List(userLB.ListIndex, 2)
      refID = myQLB.List(myQLB.ListIndex,0)
      saveNotes Me.notesBx,refID
      myQLB.ListIndex = 0
   End If
   'save 'also might as well save the workbook while we're at it
   refresh(2)
End Sub
'TODO: RESOLVE button