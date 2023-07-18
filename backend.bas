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
    'testCode = Application.InputBox("Run in test mode? (1=true, 0=false)" & vbCr _
    & "1 = True, 0 = False", "Startup", Type:=4)

    init

    'Dim start As startScreenFrm
    Set startScreen = New startScreenFrm
    Set signIn = New signInFrm
    Set queueScreen = New queueView
    Set addUsrScreen = New addUserFrm
    Set reportView = New reportFrm

    'temp bypassing
    testCode = True

    If testCode = True Then
        'MsgBox "hi, I'm in test mode!"
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

Public Function dudeWheresMyRow(ref as Integer)
'Function to return the row of a record via Reference Number
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
End Function