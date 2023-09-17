Attribute VB_Name = "lockdown"
'module to store all code to protect the worksheet
Option Explicit

Global tglState As Boolean 'false means neither button is toggled
Private Const multiPass As String = "laszloffyRu1z"
Global flag As Boolean
Dim i As Integer, j As Integer
Dim bClosing As Boolean
Global testCode As Boolean
Global numRecords As Long

Sub openSesame()
    'MsgBox "Open Sesame!"
    Dim sht As Worksheet
    'If testCode = True Then   'openSesame should always unlock/unhide everything in test mode
    For Each sht In ActiveWorkbook.Worksheets
        sht.Unprotect password:=multiPass
        sht.Visible = xlSheetVisible
    Next sht
    'End If
End Sub

Sub byeFelicia()
    MsgBox "Bye Felicia!"
    Dim sht As Worksheet
    If testCode = False Then
        For Each sht In ActiveWorkbook.Worksheets
       'MsgBox ("Sheet: " & sht.Name)
            sht.Protect password:=multiPass
            If sht.Name = "GUI" Then
            Else
                sht.Visible = xlSheetVeryHidden
            End If
        Next sht
    End If
End Sub

'Public Sub gameOver()
'sub to save and close Excel
'    MsgBox "Game over!!!"
'    If testCode = False Then
'        byeFelicia
'        Windows("SupportQ_DEV.xlsm").Activate 'make sure to only close this excel doc
'        Application.DisplayAlerts = False
'        ThisWorkbook.Save
'        Application.DisplayAlerts = True
'        ActiveWorkbook.Close SaveChanges:=False
'        Application.Quit
'    End If 
'End Sub

Function authorizer() as Boolean
'function to authenticate you to view queue
    Dim f As New authFrm
    Set f.authFrm = Me
    Dim pin As Integer
    Dim vPIN As Integer
    Dim vUSR As Integer
    Dim vUserFName

    f.Show

    pin = f.pinEntryBx.Value
    vUSR = f.userCboBx.ListIndex
    vUSR = vUSR + 2

    With dataSht
        vUserFName = .Cells(vUSR, 6).Value
        vPIN = .Cells(vUSR, 9).Value
    End With

    MsgBox (vUserFName &" was found!")

    If vPIN = pin Then'matches
        authorized = True
        authorizer = True
    Else
        authorized = False
        authorizer = False
    End If
End Function
