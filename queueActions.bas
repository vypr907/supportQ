Attribute VB_Name = "queueActions"
'For all things pertaining to the operation of a functional queue

'Proposed:
' queue is one sheet. "my queue" is filtered based off technician initials in
' respective column. "admin queue" shows all current entries. "main queue" shows
' only un-taken entries. log is merely a backup of the queue.
'
'Original Idea:
' main queue is one sheet. "my queue" is separate sheet, populated by the 'Take'
' button. Main queue grows and shrinks as users submit and technicians take.
' Upon technician RESOLVE, entry is then copied to "resolved queue" or Log sheet
' and removed from "my queue". "Admin queue" is just Log.
'
'Actual Implementation:
' Queue is one sheet. Grows via user submit, and shrinks via tech 'Take' button.
' "My Queue" is filtered by tech and resolved status from Log sheet. Log is one 
' sheet that is updated with Tech actions such as 'Take', notes editing, and 
' 'Resolve' button.

'TODO: loadQueue function. accepts int variable to determine which q to load
' 1. Main 2. User 3. Log
Sub refresh(q As Integer)
'Sub to refresh listboxes from either qSht or logSht
    Dim rw as Integer
    Dim i,d,k as Integer
    
    If q = 1 Then 'refresh main queue
        lastQRow = qSht.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).row
        With queueView
            .custQLB.ColumnCount = 10
            '                  #,time,surname,first,branch,shop,phone,reason,notes
            .custQLB.ColumnWidths = "15,0,50,40,35,25,30,60,120,80"
            .custQLB.RowSource = "Queue!A2:J" & lastQRow
            .qSizeBx = .custQLB.ListCount - 1
            .timeBx = Now
        End With
    ElseIf q = 2 Then 'refresh user queue
        'ensure that there is a "user" selected
        k = 0
        With queueView
            .myQLB.Clear
            If .techCboBx.ListIndex = -1 Then
                MsgBox "Sorry, a user must be selected",vbOk + vbExclamation,"Missing User"
                .MultiPage1.Value = 0
                .techCboBx.SetFocus
            End If
            lastLogRow = logSht.Cells(Rows.Count, 1).End(xlUp).Offset(1,0).row
            For rw = 2 to lastLogRow
                If logSht.Range("K" & CStr(rw))= .techCboBx.Value Then 'if user's initials are in tech column
                    If IsEmpty(logSht.Range("M" & CStr(rw))) Then 'only load unresolved "tickets"
                        .myQLB.AddItem 
                        For i = 1 to 10
                            .myQLB.List(k,i-1) = logSht.Cells(rw,i)
                        Next i
                        k = k + 1
                    End If
                End If
            Next rw
        End With
    ElseIf q = 3 Then 'refreshing the reports form (from logSht)
        i = 0
        d = 0
        k = 0
        lastLogRow = logSht.Cells(Rows.Count,1).End(xlUp).Offset(1,0).row
        With reportView
            .logLB.ColumnCount = 12
            '                  #,time,surname,first,branch,shop,phone,reason,notes
            .logLB.ColumnWidths = "15,0,50,40,35,25,30,60,120,80,80,80"
            'cannot use rowsource due to later need to use .additem
            '.logLB.RowSource = "Log!A2:M" & lastLogRow
            For rw = 2 to lastLogRow
                .logLB.AddItem
                For i = 1 to 12
                    .logLB.List(k,i-1) = logSht.Cells(rw,i)
                Next i
                k = k + 1
            Next rw
            .totRecordsBx = .logLB.ListCount - 1
        End With
    End If
End Sub

'FUNCTION FOR TECH TO "TAKE" A QUEUE ENTRY
Public Function takeEntry(row As Integer, ref As Integer, usr As String)
    Dim logRow As Integer
    Dim found As Range
    MsgBox "Yoink!"
    
    Set found = logSht.Range("A:A").Find(What:=ref)
    logRow = found.Row

    'STEP ONE: mark logSht w/user and timestamp
    updateLog 1,ref,usr

    'STEP TWO: remove entry from queue
    With qSht
        .Cells(row, 1).EntireRow.Delete
    End With
    refresh(1)

End Function


Sub queueAdd()
On Error Resume Next
    'sub to do all the work putting the user's entries into the queue
    
    Dim lastRow As Integer
    Dim currentRow As Integer
    Dim surname As Variant
    Dim fname As Variant
    Dim branch As Variant
    Dim rank As Variant
    Dim shop As Variant
    Dim phone As Variant
    Dim reason As Variant
    Dim notes As Variant

    Call validate
    Do While good2go = False
    Loop

    'load the values from the userform
    With signInFrm
        surname = .surnameBx
        fname = .fnameBx
        branch = .branchCboBx.Value
        rank = .rankCboBx.Value
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
        .Cells(currentRow, 6).Value = rank
        .Cells(currentRow, 7).Value = shop
        .Cells(currentRow, 8).Value = phone
        .Cells(currentRow, 9).Value = reason
        .Cells(currentRow, 10).Value = notes
    End With
    'POST TO QUEUE
    With qSht
        .Cells(currentRow, 1).Value = refID
        .Cells(currentRow, 2).Value = Format(Now, "mm/dd/yyyy HH:mm")
        .Cells(currentRow, 3).Value = surname
        .Cells(currentRow, 4).Value = fname
        .Cells(currentRow, 5).Value = branch
        .Cells(currentRow, 6).Value = rank
        .Cells(currentRow, 7).Value = shop
        .Cells(currentRow, 8).Value = phone
        .Cells(currentRow, 9).Value = reason
        .Cells(currentRow, 10).Value = notes
    End With

End Sub