VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransEntry 
   Caption         =   "Add Transaction"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "frmTransEntry.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTransEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, j As Integer, flag As Boolean
Dim bClosing As Boolean



Private Sub depositTgl_Click()
'user is wanting to deposit, disable withdrawal toggle

    On Error Resume Next
    
    If frmTransEntry.depositTgl.Value = True Then
        frmTransEntry.withdrawTgl.Enabled = False
        plusMinus = True 'deposit
        tglState = True
    Else
        frmTransEntry.withdrawTgl.Enabled = True
    End If
    
    
End Sub

Private Sub findBtn_Click()
    'manually activates GetData function, rather than auto
    GetData 'sub that pulls the data - Module1
    Me.cardNumBx.SetFocus
End Sub

Private Sub submitBtn_Click()
'==================================================================================================

    On Error Resume Next
    
    'setting target for transaction entry
    Dim ws As Worksheet
    Set ws = Worksheets("transactionLog")
    
    'establishing variables
    Dim iReply As Integer 'data validation
    Dim lRow As Long 'variable for row
    Dim amt  'variable to store amount entered
    Dim fieldsFilled As Boolean  'variable to make sure all required fields are filled
    fieldsFilled = False
    
    'setting variables for message boxes
    Dim msg, style, title, response
    msg = "Please enter an amount!"
    style = vbOKOnly + vbExclamation
    title = "Oops!"
    
    'prevents error message if they choose to cancel
    If bClosing = True Then Exit Sub
    
    'do/while loop to ensure all fields are correct for submission----------------------------
    Do   'do
    'checking for input in card number field
    'actual card number will be validated via GetData
        If cardNumBx.Value = "" Then 'card number box left blank
            response = MsgBox("This is a mandatory field. " & "Please click 'Retry', or Cancel to close the form.", _
            vbExclamation + vbRetryCancel, "Card Number Empty")
            If response = vbCancel Then 'they want to close the form
                bClosing = True
                Unload Me 'close the userform
                Exit Sub 'stop other validations from running
            Else 'they wish to enter a card number
                Cancel = True
                cardNumBx.SetFocus
                Exit Sub
            End If
        
        ElseIf amtBx.Value = "" Then 'amount field is blank
            response = MsgBox("This is a mandatory field. " & "Please click 'Retry', or Cancel to close the form.", _
            vbExclamation + vbRetryCancel, "Amount is Empty")
        
            If response = vbCancel Then 'they want to close the form
                bClosing = True
                Unload Me 'close the userform
                Exit Sub 'stop other validations from running
            Else 'they wish to enter an amount
                amtBx.Value = Format(amtBx, "$0.00")
                Cancel = True 'return focus back to amount box
                amtBx.SetFocus
                Exit Sub
            End If

        ElseIf Not IsNumeric(amtBx) Then 'not a number
            response = MsgBox("Entry must be numeric. " & "Please click 'Retry', or Cancel to close the form.", _
            vbExclamation + vbRetryCancel, "Amount is not Numeric")
            If response = vbCancel Then 'they want to close the form
                bClosing = True
                Unload Me 'close the userform
                Exit Sub
            Else 'they wish to enter an amount
                amtBx.Value = ""
                amtBx.Value = Format(amtBx, "$0.00")
                Cancel = True 'return focus back to amount box
                amtBx.SetFocus
                Exit Sub
            End If
            
        ElseIf amtBx.Value < 0 Or amtBx.Value > 2147483646 Then 'amount is too big or too small
            response = MsgBox("Entry is too large or negative. " & "Please click 'Retry', or Cancel to close the form.", _
            vbExclamation + vbRetryCancel, "Amount is out of range!")
            If response = vbCancel Then 'they want to close the form
                bClosing = True
                Unload Me 'close the userform
                Exit Sub
            Else 'they wish to enter an amount
                amtBx.Value = ""
                amtBx.Value = Format(amtBx, "0.00")
                Cancel = True
                amtBx.SetFocus 'return focus back to amount box
                Exit Sub
            End If
        
        Else 'all ok, simple format as currency
            amtBx = Format(amtBx, "$0.00")
            amt = CDec(amtBx.Value)
            fieldsFilled = True
        End If
        
    Loop Until fieldsFilled = True '-----------------------------------------------------------------
    
    'make sure Deposit or Withdrawal is selected
    If tglState = False Then
            response = MsgBox("Please select Deposit or Withdrawal", style, title)
            depositTgl.SetFocus
            Exit Sub
    End If
    
    'make amount pos or neg based on Deposit/Withdrawal selected
    If plusMinus = True Then
        amt = Abs(amt) 'ensure value is positive
    Else
        amt = amt * -1 'ensure value is negative
    End If
    
    'copying data to the transaction log--------------------------------
    'find first empty row in database
    lRow = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Row
    
    'copy data to the transaction log
    With ws
        .Cells(lRow, 1).Value = Format(Now, "m/dd/yyyy " + "HH:mm:ss") 'timestamp
        .Cells(lRow, 2).Value = transID + 1 'creates a unique transaction id from a global variable
        transID = transID + 1 'immediately updates global variable
        .Cells(lRow, 3).Value = amt 'amount of transaction
        .Cells(lRow, 4).Value = usrID ' ID for user
        .Cells(lRow, 5).Value = card 'user's card number
    End With
    '-------------------------------------------------------------------
    
    'enable View Transactions button
    frmstartMenu.transHistBtn.Enabled = True
    
    ClearForm
    
'____________________________________________________________________________________________________
End Sub


Private Sub UserForm_Initialize()
'========================================================================
    'start userform centered inside Excel screen-----------
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    '------------------------------------------------------
    
    'prevent screen flash
    ApplicationScreenUpdate = False
    
    On Error Resume Next
    
    'unlock/unhide sheets (should already be open)
    'openSesame
    
    'prompt for cashier ID---------------------------------------
    'usrID = InputBox("Please enter your cashier ID: ", "Welcome")
    enterCashier
    '------------------------------------------------------------
    'disabling 'Find' button, as not used
    'Me.findBtn.Enabled = False
    cardNumBx.SetFocus
    
    'creating and initializing variables when loading the form
    Dim transAmt As Long  'amount of transaction
    Dim currentBal As Long 'user's current account balance
    Dim cardNum As Long  'stores user-entered number for records lookup/edit
    
    'disable submit button until valid customer is found
    Me.submitBtn.Enabled = False
    
    
    'variables for the different worksheets
    Dim wsLog As Worksheet
    Dim wsData As Worksheet
    Dim wsCardSearch As Worksheet
    Dim wsBalSearch As Worksheet
    Set wsLog = Worksheets("transactionLog")
    Set wsData = Worksheets("dataStore")
    Set wsCardSearch = Worksheets("cardSearch")
    Set wsBalSearch = Worksheets("balSearch")
    
    'clearing previous searches
    wsCardSearch.Range("A2:C10000").ClearContents
    wsBalSearch.Range("A2:C10000").ClearContents
    wsBalSearch.Range("G2").Value = "Hi!"
    
    'load transaction number from backup cell
    transID = wsData.Range("I1").Value
'__________________________________________________________________________________________________
End Sub

Private Sub cardNumBx_Change()
    'On Error Resume Next
'DISABLING... USE 'FIND' BUTTON INSTEAD
    'auto populates various fields based on card number entry
    'MsgBox "change event has triggered"
    'GetData 'sub that pulls the data - Module1
    'Me.cardNumBx.SetFocus

End Sub

Private Sub closeBtn_Click()
'closing the add transaction form
    On Error Resume Next
    
    'backing up transaction ID number to reference next time form runs
    'declare object variable to hold reference to cell holding value
    Dim transactionID As Range
    
    'identify cell
    Set transactionID = ThisWorkbook.Worksheets("dataStore").Range("I1")
    
    'set cell value
    transactionID.Value = transID
    
    'hide sheets again (only exit should hide the sheets)
    'byeFelicia
    
    Unload Me
    
End Sub

Private Sub withdrawTgl_Click()
'user is wanting to withdrawal, disable deposit toggle

    On Error Resume Next
    
    If frmTransEntry.withdrawTgl.Value = True Then
        frmTransEntry.depositTgl.Enabled = False
        plusMinus = False 'withdrawal
        tglState = True
    Else
        frmTransEntry.depositTgl.Enabled = True
    End If
    

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "Hey there friend! Please utilize the 'Close' button!", vbOKOnly + vbExclamation, "You are wrong."
    End If
End Sub
