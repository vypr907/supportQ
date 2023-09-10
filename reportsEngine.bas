Attribute VB_Name = "reportsEngine"

'TODO: a function that takes all the potential search options and loads the form accordingly

Public Function logSearch(Optional tech as String, Optional rsn as String, _
Optional startRng as Variant, Optional endRng as Variant)
    'TODO: CODE GOES HERE
    Dim critRng As Range
    Dim dataRng as Range
    Dim resultRng as Range
    Dim srchRslts as Range

    temp.Activate

    With temp
        Dim status as String
        Dim isRngEmpty as Integer
        'place the values in the criteria range
        If tktState = 0 Then
            status = ""
        ElseIf tktState = 1 Then
            status = False
        Else
            status = True
        End If

        'With searchSht
        With tmpSearch
            .Cells(2,18).Value = startRng
            .Cells(2,19).Value = endRng
            .Cells(2,20).Value = tech
            .Cells(2,21).Value = status
            .Cells(2,22).Value = rsn

            Set critRng = .Range("myCriteria")
            Set resultRng = .Range("copyToRng")
            'Set srchRslts = .Range("searchResults")
        End With
        With tmpLog
            Set dataRng = .Range("logSearchRng")
        End With

        'run AdvancedFilter
        dataRng.AdvancedFilter xlFilterCopy, critRng, resultRng
        
        If Not IsError([searchResults]) Then
        'If Not IsError([srchRslts]) Then
            With reportView
                .logLB.RowSource = "searchResults"  'is this actually the correct sheet?
                '.logLB.RowSource = resultRng.Address 'ensuring this actually pulls from the temp workbook
                '.logLB.RowSource = srchRslts.Address
                .fndRecordsBx.Value = .logLB.ListCount
            End With
            'If Application.WorksheetFunction.CountA(rng) = 0 Then '=> This is redundant.
            '    MsgBox "Range is blank"
            'End If
        Else
            'MsgBox "No results found! Resetting..."
            reportView.logLB.RowSource = dataRng.Address 'ensuring this actually pulls from the temp workbook
            reportView.rsnCboBx.ListIndex = -1
        End If
    End With
End Function

Sub listBoxSort(oLB as MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
    Dim vaItems As Variant
    Dim i As Long, j As Long
    Dim c As Integer
    Dim vTemp As Variant

    'Put the items in a variant array
    vaItems = oLb.List

    'Sort the Array Alphabetically(1)
    If sType = 1 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                'Sort Ascending (1)
                If sDir = 1 Then
                    If vaItems(i, sCol) > vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If

                'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If vaItems(i, sCol) < vaItems(j, sCol) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If

            Next j
        Next i
        'Sort the Array Numerically(2)
        '(Substitute CInt with another conversion type (CLng, CDec, etc.) depending on type of numbers in the column)
    ElseIf sType = 2 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                'Sort Ascending (1)
                If sDir = 1 Then
                    If CInt(vaItems(i, sCol)) > CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If

                'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If

            Next j
        Next i
    End If

    'Set the list to the array
    oLb.List = vaItems
End Sub

Sub lbSort(sCol As Integer, sType As Integer, sDir As Integer)
    Dim sortDir as Variant
    Dim colLtr As String

    Dim sortRng As Range
    Dim resultsRng As Range

    With temp
        With tmpSearch
            Set sortRng = .Range("sortable")
            'Set resultsRng = .Range("searchResults")
            Set resultsRng = .Range("Search!$A$2:$O$"& lastLogRow)
        End With

        If sDir = 1 Then
            sortDir = xlAscending
        Else
            sortDir = xlDescending
        End If

        colLtr = ColNumToLetter(sCol)&"1"
        'With searchSht.Sort
        With tmpSearch.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range(colLtr), Order:=sortDir
            .SetRange sortRng
            .Header = xlYes
            .Apply
        End With
    End With
    
    With reportView
        '.logLB.RowSource = tmpSearch.Range("searchResults")
        .logLB.RowSource = "Search!" & resultsRng.Address
        .fndRecordsBx.Value = .logLB.ListCount
    End With
End Sub

Public Sub test2()
    Dim critRng as Range
    Set critRng = Range("myCriteria")
    Dim dataRng as Range
    Set dataRng = Range("logSearchRng")
    Dim resultRng as Range
    Set resultRng = Range("copyToRng")

    'run AdvancedFilter
    dataRng.AdvancedFilter xlFilterCopy, critRng, resultRng
End Sub

Public Sub sortTest()
init
    Dim sCol,sType,sDir As Integer
    Dim sortDir as Variant
    Dim colLtr As String
    With searchSht
        sCol = .Range("Q11").Value
        sType = .Range("R11").Value
        sDir = .Range("S11").Value
    End With
    If sDir = 1 Then
        sortDir = xlAscending
    Else
        sortDir = xlDescending
    End If
    
    Range("sortable").Sort Key1:=Range(sCol), _
                                Order1:=sortDir, _
                                Header:=xlYes
End Sub

Public Sub reportsRun_err()
    'Sub to refresh the temp workbook
    Application.DisplayAlerts = False
    tmpLog.Delete
    tmpSearch.Delete

    wb.Sheets(Array("Log", "Search")).Copy Before:=temp.Sheets(1)
    
    're-set tmp sheet variables
    Set tmpSearch = temp.Sheets("Search")
    Set tmpLog = temp.Sheets("Log")
    Application.DisplayAlerts = True
End Sub

Public Sub reportsRun()
    
    Dim lastTmpLogRow As Long

    lastLogRow = logSht.Cells(logSht.Rows.Count, 2).End(xlUp).Row

    'refresh the temp workbook
    logSht.Range("A2:O" & lastLogRow).Copy tmpLog.Range("A2:O" & lastLogRow)
    'wb.Sheets("Log").Range("logSearchRng").Copy temp.Sheets("Log").Range("logSearchRng")

    'load the search range in case user does a sort before a search
    tmpLog.Range("A2:O" & lastLogRow).Copy tmpSearch.Range("A2:O"& lastLogRow)
    'wb.Sheets("Log").Range("logSearchRng").Copy temp.Sheets("Search").Range("searchResults")
End Sub

Public Sub tempXL()
    Dim filename As String
    Dim folderPath As String
    Dim filePath As String

    folderPath = "C:\"

    filename = "temp_reportData.xlsx"
    filePath = folderPath & filename

    If Dir(filePath) <> "" Then
        'MsgBox "File exists!"
        'commenting out to test use test workbook instead
        Kill(filePath) 'easier to wipe and re-create, than to try to run comparisons
        Set temp = Workbooks.Add
        'Set temp = Workbooks.Open(filePath)
        temp.SaveAs folderPath & filename
    Else
        'MsgBox "File does not exist, creating..."
        Set temp = Workbooks.Add
        temp.SaveAs folderPath & filename
    End If

    'copy needed sheets to temp workbook
    wb.Sheets(Array("Log", "Search")).Copy Before:=temp.Sheets(1)
    'MsgBox "hello"
    
    'hows about we just hide the workbook instead of closing it
    'temp.Close SaveChanges:=True
    temp.Windows(1).Visible = False
End Sub

Public Sub unalive()
    'Dim wbName As String
    'Dim book as Workbook
    'Dim xSelect As String
    'For Each book in Application.Workbooks
    '    wbName = wbName & book.Name & vbCrLf
    'Next
    'xTitleId = "The Unaliver"
    'xSelect = Application.InputBox("Enter which workbook you want to close:" &vbCrLf & wbName, xTitleId, "", Type: = 2)
    'Application.Workbooks(xSelect).Close SaveChanges:= False
    Application.Workbooks("temp_reportData.xlsx").Close SaveChanges:= False

End Sub