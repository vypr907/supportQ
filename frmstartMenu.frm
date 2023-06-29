VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStartMenu 
   Caption         =   "Kash Kard"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmstartMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmstartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addCustBtn_Click()
    'opens Add Customer form
    If testCode = True Then
        fromTrans = False
        frmAddCard.Show vbModeless
    Else
        fromTrans = False
        frmAddCard.Show vbModal
    End If
End Sub

Private Sub addTransBtn_Click()
    'opens Transaction Entry form
    If testCode = True Then
        frmTransEntry.Show vbModeless
    Else
        frmTransEntry.Show vbModal
    End If
End Sub

Private Sub exitBtn_Click()
'exits everything
    gameOver

End Sub

Private Sub lockBtn_Click()
'unlocks form for admin access
    attempts = 3
    frmPwEntry.Show vbModal
    
End Sub

Private Sub transHistBtn_Click()
'form to view all transaction data
'hopefully add sort for customer and cashier
    'Example code from stack overflow------------------------------
    'Dim lastrow as long
    'lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    'Range("A3:D" & lastrow).Sort key1:=Range("B3:B" & lastrow), _
    '   order1:=xlAscending, Header:=xlNo
    '--------------------------------------------------------------
    If testCode = True Then
        frmTransHist.Show vbModeless
    Else
        frmTransHist.Show vbModal
    End If
    'MsgBox "This Function is still under construction!", vbOKOnly + vbExclamation, "Please Pardon our Dust"
End Sub

Private Sub unlockBtn_Click()
'relocks the form
    lockDat
End Sub

Private Sub UserForm_Initialize()

    'start userform centered inside Excel screen-----------
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    '------------------------------------------------------
    
    'disable View Transactions until at least one transaction has been added
    If transID < 1 Then
        frmStartMenuUser.transHistBtn.Enabled = False
    End If
    'disable Add Transaction until at least one customer has been added
    If numCustomers < 1 Then
        frmStartMenuUser.addTransBtn.Enabled = False
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "That's gonna be a No from me, dawg! Please use the 'Exit' button.", vbOKOnly + vbExclamation, "Nope"
    End If
End Sub

