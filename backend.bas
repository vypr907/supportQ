Attribute VB_Name = "backend"
Option Explicit
Option Base 0

'Module Variables
Dim good2go As Boolean

'loading data sheets
Sub init()
    Set wb = Workbooks("SupportQ_DEV")
    Set qSht = wb.Sheets("Queue")
    Set logSht = wb.Sheets("Log")
    Set dataSht = wb.Sheets("listData")

    authorized = False
    lastUserRow = dataSht.Cells(Rows.Count, 7).End(xlUp).Offset(1, 0).row
    'activeworkbook.Names.Add Name:="users", RefersToR2C11:="=COUNTA(C" & ColNo & ")"
    'activeworkbook.Names.Add Name:="users", RefersToR1C1:="=OFFSET($K$1,1,0,COUNTA(listData!$K:$K),1)"
End Sub

'sub to verify password
Sub comparison()
    Dim f As pwOnExitFrm
    Set f = New pwOnExitFrm
    'Set password = f.password
    f.Show
End Sub

'sub to validate all user entries
Sub validate()
    'check the things, if good:
    good2go = True
    'else things are not good
    'good2Go = False
    MsgBox ("Hi!")
End Sub

Sub queueAdd()
On Error Resume Next
    'sub to do all the work putting the user's entries into the queue
    
    Dim lastRow As Integer
    Dim currentRow As Integer
    Dim surname As Variant
    Dim fname As Variant
    Dim branch As Variant
    Dim shop As Variant
    Dim phone As Variant
    Dim reason As Variant
    Dim notes As Variant

    Call validate
    Do While good2go = False
    Loop

    'load the values from the userform
    With signInFrm
        surname = .surnameBx
        fname = .fnameBx
        branch = .branchCboBx.Value
        shop = .shopBx
        phone = .phoneBx
        reason = .reasonCboBx.Value
        notes = .notesBx
    End With

    'find the last row
    lastRow = logSht.Cells(Rows.Count, 1).End(xlUp).row
    'get the value
    refID = logSht.Cells(lastRow, 1).Value
    'increment the value and row, and place the value
    refID = refID + 1
    currentRow = lastRow + 1
    logSht.Cells(currentRow, 1).Value = refID

    'POST TO LOG
    With logSht
        .Cells(currentRow, 1).Value = refID
        .Cells(currentRow, 2).Value = Format(Now, "mm/dd/yyyy HH:mm")
        .Cells(currentRow, 3).Value = surname
        .Cells(currentRow, 4).Value = fname
        .Cells(currentRow, 5).Value = branch
        .Cells(currentRow, 6).Value = shop
        .Cells(currentRow, 7).Value = phone
        .Cells(currentRow, 8).Value = reason
        .Cells(currentRow, 9).Value = notes
    End With
    'POST TO QUEUE
    With qSht
        .Cells(currentRow, 1).Value = refID
        .Cells(currentRow, 2).Value = Format(Now, "mm/dd/yyyy HH:mm")
        .Cells(currentRow, 3).Value = surname
        .Cells(currentRow, 4).Value = fname
        .Cells(currentRow, 5).Value = branch
        .Cells(currentRow, 6).Value = shop
        .Cells(currentRow, 7).Value = phone
        .Cells(currentRow, 8).Value = reason
        .Cells(currentRow, 9).Value = notes
    End With

End Sub

Sub clearForm
    'empty the sign-in form
    With signInFrm
        'TODO: empty the form
        .surnameBx.Value = ""
        .fnameBx.Value = ""
        .branchCboBx.ListIndex = -1
        .rankBx.Value = ""
        .shopBx.Value = ""
        .phoneBx.Value = ""
        .reasonCboBx.ListIndex = -1
        .notesBx.Value = ""
    End With
End Sub

Sub start()
    testCode = Application.InputBox("Run in test mode? (1=true, 0=false)" & vbCr _
    & "1 = True, 0 = False", "Startup", Type:=4)

    init

    'Dim start As startScreenFrm
    Set startScreen = New startScreenFrm
    Set signIn = New signInFrm
    Set queueScreen = New queueView
    Set addUsrScreen = New addUserFrm

    
    If testCode = True Then
        MsgBox "hi, I'm in test mode!"
        startScreenFrm.Show vbModeless
        Application.ScreenUpdating = True
        openSesame
    Else
        MsgBox "hi, I'm in regular mode!"
        startScreenFrm.Show vbModal
        byeFelicia
        Application.ScreenUpdating = False
    End If
End Sub

Sub save()
    Windows("SupportQ_DEV.xlsm").Activate 'make sure to only close this excel doc
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub
Sub gameOver()
    'sub to save and close Excel 
    MsgBox "Game over!!!"
    
    byeFelicia 'lock and hide sheets
    Windows("SupportQ_DEV.xlsm").Activate 'make sure to only close this excel doc
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    ActiveWorkbook.Close SaveChanges:=False
    
End Sub

Public Function takeEntry(row As Integer, ref As Integer, usr As String)
    Dim logRow As Integer
    Dim found As Range
    MsgBox "Yoink!"
    
    'TODO: Find row by matching ref
    Set found = logSht.Range("A:A").Find(What:=ref)
    logRow = found.Row

    'STEP ONE: mark logSht w/user and timestamp
    With logSht
        .Cells(logRow,10).Value = usr
        .Cells(logRow,11).Value = Now
    End With

    'STEP TWO: remove entry from queue
    With qSht
        .Cells(row, 1).EntireRow.Delete
    End With
    refresh

End Function

Sub refresh()
    Dim lastRow As Integer
    lastRow = qSht.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row
    With queueView
        .custQLB.ColumnCount = 9
        '                  #,time,surname,first,branch,shop,phone,reason,notes
        .custQLB.ColumnWidths = "15,0,50,40,35,30,60,120,80"
        .custQLB.RowSource = "Queue!A2:I" & lastRow
        .qSizeBx = .custQLB.ListCount - 1
        .timeBx = Now
    End With
End Sub