VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clUserFormEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public WithEvents mLabelGroup As MSForms.Label
Dim tgl As Boolean

Private Sub mLabelGroup_Click()
    'MsgBox mLabelGroup.Caption & " has been pressed"
    If mLabelGroup.Caption = "#" Then
      'listBoxSort reportFrm.logLB,0,1,tgl+1
      lbSort 1,1,tgl+1
      If tgl = True Then
        tgl = False
        MsgBox "Sort Ascending"
      Else
        tgl = True
        MsgBox "Sort Descending"
      End If
    End If
    If mLabelGroup.Caption = "Time" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Surname" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Branch" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Rank" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Shop" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Reason" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Tech" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Taken" Then
      MsgBox "I can sort by REF!"
    End If
    If mLabelGroup.Caption = "Resolved" Then
      MsgBox "I can sort by REF!"
    End If
End Sub