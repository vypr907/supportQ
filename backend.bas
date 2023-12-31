Attribute VB_Name = "backend"
Option Explicit
Option Base 0

'Global Variables
public password as Variant
public refID As Integer

'Module Variables
Dim good2go As Boolean

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

'sub to validate all user entries
sub validate()
    'check the things, if good:
    good2Go = True
    'else things are not good
    'good2Go = False
    MsgBox("Hi!")
End Sub

sub queueAdd()
On Error Resume Next
    'sub to do all the work putting the user's entries into the queue
    
    Dim lastRow As Integer
    Dim currentRow As Integer
    Dim surname As Variant
    Dim fname As Variant
    Dim branch As Variant
    Dim shop As Variant
    Dim phone As Variant
    Dim reason As Variant
    Dim notes As Variant

    Call validate
    Do While good2Go = False
    loop

    'load the values from the userform
    With signInFrm
        surname = .surnameBx
        fname = .fnameBx
        branch = .branchCboBx.Value
        shop = .shopBx
        phone = .phoneBx
        reason = .reasonCboBx.Value
        notes = .notesBx
    End With

    'find the last row
    lastRow = logSht.Cells(Rows.Count, 1).End(xlUp).row
    'get the value
    refID = logSht.Cells(lastRow, 1).Value
    'increment the value and row, and place the value
    refID = refID + 1
    currentRow = lastRow + 1
    logSht.Cells(currentRow, 1).Value = refID

    'POST TO LOG
    With logSht
        .Cells(currentRow, 1).Value = refID
        .Cells(currentRow, 2).Value = Format(Now, "mm/dd/yyyy HH:mm")
        .Cells(currentRow, 3).Value = surname
        .Cells(currentRow, 4).Value = fname
        .Cells(currentRow, 5).Value = branch
        .Cells(currentRow, 6).Value = shop
        .Cells(currentRow, 7).Value = phone
        .Cells(currentRow, 8).Value = reason
        .Cells(currentRow, 9).Value = notes
    End With
    'POST TO QUEUE
    With queueSht
        .Cells(currentRow, 1).Value = refID
        .Cells(currentRow, 2).Value = Format(Now, "mm/dd/yyyy HH:mm")
        .Cells(currentRow, 3).Value = surname
        .Cells(currentRow, 4).Value = fname
        .Cells(currentRow, 5).Value = branch
        .Cells(currentRow, 6).Value = shop
        .Cells(currentRow, 7).Value = phone
        .Cells(currentRow, 8).Value = reason
        .Cells(currentRow, 9).Value = notes
    End With

End Sub 