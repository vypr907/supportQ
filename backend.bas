Attribute VB_Name = "backend"
Option Explicit
Option Base 0


'loading data sheets
Sub init()
    'Set wb = Workbooks("SupportQ_DEV")
    Set wb = Excel.ActiveWorkbook
    Set qSht = wb.Sheets("Queue")
    Set logSht = wb.Sheets("Log")
    Set dataSht = wb.Sheets("listData")
    Set searchSht = wb.Sheets("Search")

    authorized = False
    lastUserRow = dataSht.Cells(Rows.Count, 7).End(xlUp).Offset(1, 0).row
End Sub

'sub to verify password
Sub comparison()
    Dim f As pwOnExitFrm
    Set f = New pwOnExitFrm
    f.Show
End Sub

Sub getName()
    Dim n As nameFrm
    Set n = New nameFrm
    n.Show
End Sub

Sub validate()
'sub to validate all user entries
    'check the things, if good:
    good2Go = True
    'else things are not good
    'good2Go = False
End Sub

Function validName(ByVal fileName As String) As Boolean
    Application.ScreenUpdating = False
    Dim tp As Workbook
    
    'check for nothing
    If fileName = "" Then
        validName = False
        GoTo ExitProc
    End If
    
    'create temp file
    On Error GoTo InvalidName:
    
    Set tp = Workbooks.Add
    tp.SaveAs Environ("temp") & "\" & fileName & ".xlsx", 51
    
    On Error Resume Next
    
    'close temp file
    tp.Close False
    Kill Environ("temp") & "\" & fileName & ".xlsx"
    
    'Name is valid, exit function
    validName = True
    GoTo ExitProc
    
'If file cannot be created
InvalidName:
    On Error Resume Next
    
    'close temp excel file
    tp.Close False
    
    'file name is NOT valid, exit function
    validName = False
    
ExitProc:
Application.ScreenUpdating = True
End Function

Sub clearForm()
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
    MsgBox "Running startup...", , "Welcome!"
    init
    
    Set startScreen = New startScreenFrm
    Set signIn = New signInFrm
    Set queueScreen = New queueView
    Set addUsrScreen = New addUserFrm
    Set reportView = New reportFrm

    startScreenFrm.Show vbModeless
    Application.ScreenUpdating = True
End Sub

Sub save()
    Windows("SupportQ_DEV.xlsm").Activate 'make sure to only close this excel doc
    Application.DisplayAlerts = False
    ThisWorkbook.save
    Application.DisplayAlerts = True
End Sub

Sub gameOver()
    'sub to save and close Excel
    MsgBox "Game over!!!", , "nuke.exe"

    Windows("SupportQ_DEV.xlsm").Activate 'make sure to only close this excel doc
    Application.DisplayAlerts = False
    ThisWorkbook.save
    Application.DisplayAlerts = True
    ActiveWorkbook.Close SaveChanges:=False
    Application.Quit
End Sub

Public Function dudeWheresMyRow(ref As Integer)
'Function to return the row of a record via Reference Number
    Dim found As Range
    Set found = logSht.Range("A:A").Find(What:=ref)
    dudeWheresMyRow = found.row
End Function

Public Function saveNotes(text As String, ref As Integer)

    Dim found As Range
    'MsgBox "Nice save, bro!", , "Bro."
    popUp "Nice save, bro!", "Bro.", 1
    With logSht
        .Cells(dudeWheresMyRow(ref), 10).Value = text
    End With
    save
End Function

Public Function updateLog(q As Integer, ref As Integer, Optional usr As String)
    Dim here As Integer
    here = dudeWheresMyRow(ref)
    With logSht
        If q = 1 Then 'User has taken from the queue
            .Cells(here, 11).Value = usr
            .Cells(here, 12).Value = Now
        Else 'user has resolved an entry
            .Cells(here, 13).Value = Now
            .Cells(here, 14).Value = True
        End If
    End With
End Function

Public Function ColNumToLetter(ColNumber As Integer)
    Dim ColLetter As String
    'Convert To Column Letter
    ColNumToLetter = Split(Cells(1, ColNumber).Address, "$")(1)
End Function

Public Function popUp(msg As String, title As String, time As Integer)
    Dim InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'set the message box to close after 10 seconds
    'time = 1
    Select Case InfoBox.popUp(msg, time, title, 0)
        Case 1, -1
            Exit Function
    End Select
End Function

Public Function popUpTest(msg As String, title as String, time As Integer, Optional ByVal buttons As VbMsgBoxStyle = vbOK) As VbMsgBoxResult
    Dim fso As Object 'FileSystemObject
    Dim wss As Object 'WshShell
    Dim TempFile As String

    On Error GoTo ExitPoint
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wss = CreateObject("WScript.Shell")

    With fso
        TempFile = .BuildPath(.GetSpecialFolder(2).Path, .GetTempName & ".vbs")
        With .CreateTextFile(TempFile)
            .WriteLine "Set wss = CreateObject(""WScript.Shell"")" & vbCrLf & _
            "i = wss.Popup(""" & msg & """," & time & ", """ & title & _
            """," & buttons & ")" & "" & vbCrLf & "WScript.Quit i"
            .Close
        End With
    End With

    popUpTest = wss.Run(TempFile, 1, True)
    fso.DeleteFile TempFile, True

    ExitPoint:
End Function

Public Sub testMsg()
    Dim val As String
    
    val = InputBox("Enter some text:", "This is a title")
    popUp (val)
End Sub

