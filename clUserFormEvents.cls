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
      lbSort 1,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Time" Then
      lbSort 2,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Surname" Then
      lbSort 3,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Branch" Then
      lbSort 5,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Rank" Then
      lbSort 6,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Shop" Then
      lbSort 7,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Reason" Then
      lbSort 9,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Tech" Then
      lbSort 11,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Taken" Then
      lbSort 12,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
    If mLabelGroup.Caption = "Resolved" Then
      lbSort 13,1,tgl+1
      If tgl = True Then
        tgl = False
      Else
        tgl = True
      End If
    End If
End Sub