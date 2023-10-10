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

Private Sub printBtn_Click()
    'MsgBox "Sorry, this function is still under construction.", , "ERR_404_func_not_found"
    printMe
End Sub

Sub userForm_Initialize()
    'load users
    Dim item As Variant
    For Each item In dataSht.Range("users")
        With Me.techCboBx2
            .AddItem item.Value
        End With
    Next item
    'load reasons
    For Each item In dataSht.Range("reasonCode")
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
    With temp
        Dim i, d, k
        i = 0
        d = 0
        k = 0
        lastLogRow = logSht.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row
        With Me
            .logLB.ColumnCount = 13
            .logLB.ColumnWidths = "15,70,60,50,35,35,40,60,120,150,25,65,65"
            .logLB.RowSource = "Log!A2:M" & lastLogRow
            .totRecordsBx = .logLB.ListCount - 1
        End With
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
            If IsDate(.text) Then
                .text = Format(DateValue(.text), "mm/dd/yyyy")
                startDate = ">=" + .text
            Else
                MsgBox "Please enter a valid start date! (mm/dd/yyyy)"
                Exit Sub
            End If
        End If
    End With
    With endDateBx
        If .Value <> "" Then
            If IsDate(.text) Then
                .text = Format(DateValue(.text), "mm/dd/yyyy")
                endDate = "<=" + .text
            Else
                MsgBox "Please enter a valid end date! (mm/dd/yyyy)"
                Exit Sub
            End If
        End If
    End With
    logSearch Me.techCboBx2.Value, Me.rsnCboBx.Value, startDate, endDate
End Sub

Public Sub printMe()
    Dim endRow As Integer
    'Dim picker As FileDialog
    'Dim myFolder As String
    
    With Me.logLB
        If .ListCount = 0 Then
            MsgBox "No data to print!", , "Empty Report"
            Exit Sub
        End If
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Dim prt As Workbook
        Set prt = Workbooks.Add
        
        'copy headers
        logSht.Range("A1:O1").Copy Destination:=prt.Sheets(1).Range("A1:O1")
        'copy contents of listbox
        Range("A2").Resize(.ListCount, .ColumnCount).Value = .List
    End With
    
    'Get name for report
    Call getName
    
    'get folder from user
    'CUSTOMER DECIDED THEY WANTED A HARD-CODED SAVE LOCATION--------------------------
    'Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    
    'With picker
    '    .title = "Select where to save your report"
    '    .AllowMultiSelect = False
    '    If .Show <> -1 Then Exit Sub 'checking to see if user clicked cancel
    '    myFolder = .SelectedItems(1) & "\"
    'End With
    '----------------------------------------------------------------------------------
    
    'format sheet for printing
    With prt.Sheets(1)
        endRow = .Cells(Rows.Count, 1).End(xlUp).Offset(2, 0).row
        For k = 2 To endRow
            If .Cells(k, 2).Value <> "" Then
                .Cells(k, 2).Value = CDate(Cells(k, 2).Value)
            End If
            If .Cells(k, 12).Value <> "" Then
                .Cells(k, 12).Value = CDate(Cells(k, 12).Value)
            End If
            If .Cells(k, 13).Value <> "" Then
                .Cells(k, 13).Value = CDate(Cells(k, 13).Value)
            End If
        Next k
        .Columns("O").Delete 'don't need last 'Date' column
        .Columns("N").Delete 'don't need 'Resolved' column
        .Columns("J").Delete 'don't need 'Notes' column
        .Columns("H").Delete 'don't need 'Phone' column
        .Cells.Font.Size = 9
        .Columns("A:O").AutoFit
        .PageSetup.Orientation = xlLandscape
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=ThisWorkbook.Path & "\REPORTS\" & namer _
            & "_" & Format(Now(), "YYYY-MM-DD")
        'CUST DECIDED THEY WANTED A HARDCODE SAVE LOCATION--------------------------------------
        '.ExportAsFixedFormat Type:=xlTypePDF, filename:=myFolder & "report.pdf"
        '---------------------------------------------------------------------------------------
    End With
    prt.Close False
    MsgBox "File has been created.", , "Success!"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If CloseMode = vbFormControlMenu Then
      Cancel = True
      Me.Hide
   End If
End Sub

Private Sub closeBtn_Click()
   Me.Hide
   temp.Windows(1).Visible = False
End Sub

