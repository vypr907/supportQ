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

Dim mcolEvents As Collection

Sub userForm_Initialize()
   'load users
   Dim item As Variant
   For Each item In dataSht.Range("users")
      With Me.techCboBx2
         .AddItem item.Value
      End With
   Next item
   'load reasons
   For Each item in dataSht.Range("reasonCode")
      With Me.rsnCboBx
         .AddItem item.Value
      End With
   Next item

   'load labels into collection for sort purposes later
   Dim cLblEvents As clUserFormEvents
   Dim ctl As MSForms.Control
   Set mcolEvents = New Collection
   For Each ctl In Me.Controls
      If TypeName(ctl) = "Label" Then
         Set cLblEvents = New clUserFormEvents
         Set cLblEvents.mLabelGroup = ctl
         mcolEvents.Add cLblEvents
      End If
   Next
   
   'load log entries
   Dim i,d,k
   i = 0
   d = 0
   k = 0
   lastLogRow = logSht.Cells(Rows.Count,1).End(xlUp).Offset(1,0).row
   With Me
      .logLB.ColumnCount = 13
      .logLB.ColumnWidths = "15,70,60,50,35,35,40,60,120,150,25,65,65"
      .logLB.RowSource = "Log!A2:M" & lastLogRow
      'CAN'T USE .AddItem with more than 10 columns
      'For rw = 2 to lastLogRow
      '   .logLB.AddItem
      '   For i = 1 to 12
      '      .logLB.List(k,i-1) = logSht.Cells(rw,i)
      '   Next i
      '   k = k + 1
      'Next rw
      .totRecordsBx = .logLB.ListCount - 1
   End With
End Sub

Sub searchBtn_Click()
   'lets validate some shit
   If tktAll.Value = True Then
      tktState = 0
   ElseIf tktOpen.Value = True Then
      tktState = 1
   ElseIf tktClosed.Value = True Then
      tktState = 2
   Else
      tktState = 0
   End If

   'put date validation shit here
   With startDateBx
      If .Value <> "" Then
         If IsDate(.Text) Then
            .Text = Format(DateValue(.Text), "mm/dd/yyyy")
            startDate = ">=" + .Text
         Else  
            MsgBox "Please enter a valid start date! (mm/dd/yyyy)"
            Exit Sub
         End If
      End If
   End With
   With endDateBx
      If .Value <> "" Then
         If IsDate(.Text) Then
            .Text = Format(DateValue(.Text), "mm/dd/yyyy")
            endDate = "<=" + .Text
         Else  
            MsgBox "Please enter a valid end date! (mm/dd/yyyy)"
            Exit Sub
         End If
      End If
   End With
   logSearch Me.techCboBx2.Value,Me.rsnCboBx.Value,startDate,endDate
End Sub

'Sub refLbl_Click()
'   MsgBox "Hi, I'm REF"
'End Sub

'Sub timeLbl_Click()
'   MsgBox "Hi, I'm TIME"
'End Sub

'Sub lnameLbl_Click()
'   MsgBox "Hi, I'm SURNAME"
'End Sub

'Sub branchLbl_Click()
'End Sub

'Sub rankLbl_Click()
'End Sub

'Sub shopLbl_Click()
'End Sub

'Sub rsnLbl_Click()
'End Sub

'Sub techLbl_Click()
'End Sub

'Sub takenLbl_Click()
'End Sub

'Sub resolvedLbl_Click()
'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = vbFormControlMenu Then
      Cancel = True
      Me.Hide
   End If
End Sub

Private Sub closeBtn_Click()
   Me.Hide
End Sub