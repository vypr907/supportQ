Attribute VB_Name = "Module1"
Option Explicit

Global transID As Integer 'global variable to increment transaction numbers
Global plusMinus As Boolean 'global variable to determine whether transaction amt
'is deposit or withdrawal. True is deposit, False is withdrawal
Global tglState As Boolean 'false means neither button is toggled
Global usrID As Integer 'variable to hold cashier ID
Private Const multiPass As String = "laszloffyRu1z" 'password to lock/unlock sheets
Global flag As Boolean
Global card As Long
Dim i As Integer, j As Integer
Dim bClosing As Boolean
Global testCode As Boolean
Global numRecords As Long
Global numCustomers As Long
Global fromTrans As Boolean
Global attempts As Integer

'creating instances of forms for admin and normal user
Global frmStartMenuAdmin As frmstartMenu
Global frmStartMenuUser As frmstartMenu




Sub GUI_btnAddCard_Click()
'NOT USED
    'opens Add Customer form
    'frmAddCard.Show vbModeless
End Sub

Sub GUI_btnAddTrans_Click()
'NOT USED
    'keep screen from flashing
    'Application.ScreenUpdating = False
    
    'unprotecting/unhiding sheets
    'openSesame
    
    'show form
    'frmTransEntry.Show vbModeless
End Sub

Sub openSesame()
    'MsgBox "Open Sesame!"
    Dim sht As Worksheet
    
    For Each sht In ActiveWorkbook.Worksheets
        sht.Unprotect password:=multiPass
        sht.Visible = xlSheetVisible
    Next sht
    'End If
End Sub

Sub byeFelicia()
    'MsgBox "Bye Felicia!"
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
    Application.ScreenUpdating = False
End Sub

Sub gameOver()
    'sub to save and close Excel
    
    'MsgBox "Game over!!!"
    'save numCustomers to sheet------------------------------
    Dim customers As Range
    Set customers = ThisWorkbook.Worksheets("dataStore").Range("I2")
    customers.Value = numCustomers
    '--------------------------------------------------------
    byeFelicia 'lock and hide sheets
    Windows("EagleCash - Test Blank.xlsm").Activate 'make sure to only close this excel doc
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    ActiveWorkbook.Close SaveChanges:=False
    
End Sub

Sub GetData()

'If the card number box contains a numeric value, searches for card number and loads the corresponding record.
'If code cannot find the card number, empties the other text boxes.
'If the card number box does not contain a numeric value, clears form

    On Error Resume Next
    
    'variables for the different worksheets
    Dim wsLog As Worksheet
    Dim wsData As Worksheet
    Dim wsCardSearch As Worksheet
    Dim wsBalSearch As Worksheet
    Set wsLog = Worksheets("transactionLog")
    Set wsData = Worksheets("dataStore")
    Set wsCardSearch = Worksheets("cardSearch")
    Set wsBalSearch = Worksheets("balSearch")
    
    'MsgBox "GetData has loaded"
    
    
    Dim rowNum As Long
    Dim searchRow As Long
    Dim addMe As Range
    Dim x As Integer
    Dim response
    
    'ensure a clean pull of data
    wsCardSearch.Range("A2:C1000000").ClearContents
    wsBalSearch.Range("A2:C1000000").ClearContents
    
    Set addMe = wsBalSearch.Cells(Rows.Count, 3).End(xlUp).Offset(1, 0)
    
    rowNum = 2
    searchRow = 2
    'verifying numeric input----------------------------------------------------
    If IsNumeric(frmTransEntry.cardNumBx.Value) Then 'verifying numeric
        flag = False
        i = 2
        card = frmTransEntry.cardNumBx.Value
        
        Do While wsData.Cells(i, 1).Value <> ""

            If wsData.Cells(i, 1).Value = card Then 'card number was found, load data
                flag = True
                'MsgBox "match found!"  assigning data to userform
                With frmTransEntry
                    .firstNmBx.Value = wsData.Cells(i, 2).Value
                    .lastNmBx.Value = wsData.Cells(i, 3).Value
                    .submitBtn.Enabled = True
                End With
                
                'searching/loading transactions----------------------------
                Do Until wsLog.Cells(rowNum, 1).Value = ""

                    If wsLog.Cells(rowNum, 5).Value = card Then
                        wsBalSearch.Cells(searchRow, 1).Value = wsLog.Cells(rowNum, 2).Value 'Transaction ID
                        wsBalSearch.Cells(searchRow, 2).Value = wsLog.Cells(rowNum, 3).Value 'Amount
                        wsBalSearch.Cells(searchRow, 3).Value = wsLog.Cells(rowNum, 4).Value 'Cashier ID
                        searchRow = searchRow + 1
                    End If
                    rowNum = rowNum + 1
            
                Loop
    
                If searchRow = 2 Then
                MsgBox "No transactions found for this account.", vbOKOnly, "Transaction History"
                'Exit Sub
                End If
                '----------------------------------------------------------
                
                'adding found transactions to listbox----------------------
                Dim tempIndex As Long
                
                With frmTransEntry.ListBox1
                    .ColumnCount = 3
                    'NEW
                    MsgBox "value = " + Range("A2").Resize(searchRow - 2, 3).Value
                    .List = wsBalSearch.Range("A2").Resize(searchRow - 2, 3).Value
                    
                    'formatting second column as currency
                    For tempIndex = 0 To .ListCount - 1
                        .List(tempIndex, 1) = (Format(Val(.List(tempIndex, 1)), "$#,##0.00"))
                    Next
                    .ColumnWidths = "40;70;10"
                End With
                '----------------------------------------------------------
                'add balance to user form----------------------------------
                frmTransEntry.balBx.Value = wsBalSearch.Range("E2").Value
                frmTransEntry.balBx = Format(frmTransEntry.balBx, "$#,##0.00")
                '----------------------------------------------------------
                
                
            End If

            i = i + 1
        Loop
        If flag = False Then 'card number is not found, reset the form
            response = MsgBox("Card Number not found, do you wish to add this as a new customer?", _
            vbYesNo, "Customer Not Found")
            If response = vbYes Then
                fromTrans = True
                frmAddCard.Show vbModal
                Exit Sub
            ElseIf response = vbNo Then
                fromTrans = False
                ClearForm
                Exit Sub
            End If
        End If

    Else
        ClearForm 'anything other than a number gets auto-cleared
    End If
    '---------------------------------------------------------------------------
        
End Sub

Sub ClearForm()

    On Error Resume Next
    
    'clear out form for next entry
    With frmTransEntry
        .cardNumBx.Value = ""
        .amtBx.Value = ""
        .firstNmBx.Value = ""
        .lastNmBx.Value = ""
        .balBx.Value = ""
        tglState = False
        .depositTgl.Enabled = True
        .depositTgl.Value = False
        .withdrawTgl.Enabled = True
        .withdrawTgl.Value = False
        .ListBox1.RowSource = ""
        .ListBox1.Clear
        .submitBtn.Enabled = False
        .cardNumBx.SetFocus
    End With

End Sub

Sub enterCashier()
Line1: 'reset point for input validation
    Application.ScreenUpdating = True 'turn on screen updating so that inputbox works correctly
    usrID = Application.InputBox("Please enter your cashier ID: ", "Welcome")
    If usrID = False Then 'X was pressed
        MsgBox "I'm sorry Dave, I can't let you do that...", vbOKOnly + vbExclamation, "Error....Error....Errrrrr...."
        GoTo Line1
        'Exit Sub
    End If
    If usrID - Int(usrID) <> 0 Then 'checking for valid INT input
        MsgBox "That wasn't a valid number!", vbOKOnly, "Oops!"
        GoTo Line1
    End If
    Application.ScreenUpdating = False  'turning back off for visual
End Sub

Function adminUnlock(password As String, x As Integer) As Boolean
'function that takes a password and number of attempts allowed, and returns t/f

    Dim touchdown As Boolean
    
    If password = multiPass Then
        touchdown = True
    Else
        x = x - 1
        touchdown = False
        adminUnlock = touchdown 'return value
        Exit Function
    End If

    adminUnlock = touchdown 'return value
        
End Function

Sub start()
    
    'auto run macros to hide sheets and pop up GUI
    byeFelicia 'make sure everything is hidden/locked to look pretty
    
    'by default, start as normal user
    Application.ScreenUpdating = False 'turn screenupdate off so it keeps looking pretty
    openSesame 'show/unlock sheets, so shit works
    
    Dim data As Worksheet
    Dim customers As Range
    Set data = ThisWorkbook.Worksheets("dataStore")
    
    'load numCustomers---------------------------------------------------------
    Set customers = data.Range("I2")
    numCustomers = customers.Value
    '--------------------------------------------------------------------------
    'load transID--------------------------------------------------------------
    transID = data.Range("I1").Value
    '--------------------------------------------------------------------------
    attempts = 3
    
    'everything loaded, show the form
    'set start menu forms for user and admin
    Set frmStartMenuUser = New frmstartMenu
    Set frmStartMenuAdmin = New frmstartMenu
    
    frmStartMenuUser.Show vbModal
    
End Sub

Sub unlockDat()
    'unlock the form
    'MsgBox "it's a match"
    testCode = True
    openSesame
    Unload frmPwEntry
    frmStartMenuUser.Hide
    frmStartMenuAdmin.lockBtn.Visible = False
    frmStartMenuAdmin.Show vbModeless
    
End Sub

Sub lockDat()
    'locks the form
    byeFelicia
    Unload frmStartMenuAdmin
    frmStartMenuUser.Show vbModal
End Sub

