Attribute VB_Name = "backend"
Option Explicit
Option Base 0

public password as Variant

'loading data sheets
    Public wb As Workbook
    Public queueSht As Worksheet
    Public listSht As Worksheet
    Public logSht As Worksheet

'sub to verify password
Sub comparison()
    
    Dim f as pwOnExitFrm
    Set f = New pwOnExitFrm
    'Set password = f.password
    f.Show

End Sub
'update