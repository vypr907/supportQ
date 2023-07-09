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
    
    Dim logRow As Integer
    Dim qRow As Integer
    Dim currentRow As Integer
    Dim currentQRow As Integer
    Dim surname As Variant
    Dim fname As Variant
    Dim branch As Variant
    Dim rank As Variant
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
        rank = .rankCboBx.Value
        shop = .shopBx
        phone = .phoneBx
        reason = .reasonCboBx.Value
        notes = .notesBx
    End With

    'find the last row of the queue
    qRow = qSht.Cells(Rows.Count, 1).End(xlUp).row
    'find the last row of the log
    logRow = logSht.Cells(Rows.Count, 1).End(xlUp).row
    'get the value
    refID = logSht.Cells(logRow, 1).Value
    'increment the value and row, and place the value
    refID = refID + 1
    currentQRow = qRow + 1
    currentRow = logRow + 1
    logSht.Cells(currentRow, 1).Value = refID

    'POST TO LOG
    With logSht
        .Cells(currentRow, 1).Value = refID
        .Cells(currentRow, 2).Value = Format(Now, "mm/dd/yyyy HH:mm")
        .Cells(currentRow, 3).Value = surname
        .Cells(currentRow, 4).Value = fname
        .Cells(currentRow, 5).Value = branch
        .Cells(currentRow, 6).Value = rank
        .Cells(currentRow, 7).Value = shop
        .Cells(currentRow, 8).Value = phone
        .Cells(currentRow, 9).Value = reason
        .Cells(currentRow, 10).Value = notes
    End With
    'POST TO QUEUE
    With qSht
        .Cells(currentQRow, 1).Value = refID
        .Cells(currentQRow, 2).Value = Format(Now, "mm/dd/yyyy HH:mm")
        .Cells(currentQRow, 3).Value = surname
        .Cells(currentQRow, 4).Value = fname
        .Cells(currentQRow, 5).Value = branch
        .Cells(currentQRow, 6).Value = rank
        .Cells(currentQRow, 7).Value = shop
        .Cells(currentQRow, 8).Value = phone
        .Cells(currentQRow, 9).Value = reason
        .Cells(currentQRow, 10).Value = notes
    End With

End Sub

Sub clearForm
    'empty the sign-in form
    With signInFrm
        'TODO: empty the form
        .surnameBx.Value = ""
        .fnameBx.Value = ""
        .branchCboBx.ListIndex = -1
        .rankCboBx.ListIndex = -1
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
        byeFelicia
        startScreenFrm.Show vbModal
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
    
    Set found = logSht.Range("A:A").Find(What:=ref)
    logRow = found.Row

    'STEP ONE: mark logSht w/user and timestamp
    updateLog 1,ref,usr

    'STEP TWO: remove entry from queue
    With qSht
        .Cells(row, 1).EntireRow.Delete
    End With
    refresh(1)
    save

End Function

Sub refresh(q As Integer)
    Dim rw as Integer
    Dim i,d,k as Integer
    
    If q = 1 Then 'refresh main queue
        lastQRow = qSht.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row
        With queueView
            .custQLB.ColumnCount = 10
            '                  #,time,surname,first,branch,shop,phone,reason,notes
            .custQLB.ColumnWidths = "15,0,50,40,35,25,30,60,120,80"
            .custQLB.RowSource = "Queue!A2:J" & lastQRow
            .qSizeBx = .custQLB.ListCount - 1
            .timeBx = Now
        End With
    Else 'If q = 2 Then 'refresh user queue
        'ensure that there is a "user" selected
        k = 0
        With queueView
            .myQLB.Clear
            If .techCboBx.ListIndex = -1 Then
                MsgBox "Sorry, a user must be selected",vbOk + vbExclamation,"Missing User"
                .MultiPage1.Value = 0
                .techCboBx.SetFocus
            End If
            lastLogRow = logSht.Cells(Rows.Count, 1).End(xlUp).Offset(1,0).row
            For rw = 2 to lastLogRow
                If logSht.Range("K" & CStr(rw))= .techCboBx.Value Then 'if user's initials are in tech column
                    If IsEmpty(logSht.Range("M" & CStr(rw))) Then 'only load unresolved "tickets"
                        .myQLB.AddItem 
                        For i = 1 to 10
                            .myQLB.List(k,i-1) = logSht.Cells(rw,i)
                        Next i
                        k = k + 1
                    End If
                End If
            Next rw
        End With
    End If
End Sub

Public Function dudeWheresMyRow(ref as Integer)
    Dim found As Range
    Set found = logSht.Range("A:A").Find(What:=ref)
    dudeWheresMyRow = found.Row
End Function

Public Function saveNotes(text as String, ref as Integer)

    Dim found As Range
    MsgBox "Nice save, bro!",,"Bro."

    With logSht
        .Cells(dudeWheresMyRow(ref), 10).Value = text
    End With
    save
End Function

Public Function updateLog(q as Integer, ref as integer, Optional usr as String)
    Dim here As Integer
    here = dudeWheresMyRow(ref)
    With logSht
        If q = 1 Then 'User has taken from the queue
            .Cells(here,11).Value = usr
            .Cells(here,12).Value = Now
        Else 'user has resolved an entry
            .Cells(here,13).Value = Now
        End If
    End With
    save
End Function