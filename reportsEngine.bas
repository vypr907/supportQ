Attribute VB_Name = "reportsEngine"
'a function that takes all the potential search options and loads the form accordingly

Public Function logSearch(Optional tech As String, Optional rsn As String, _
Optional startRng As Variant, Optional endRng As Variant)
    Dim critRng As Range
    Dim dataRng As Range
    Dim resultRng As Range
    Dim srchRslts As Range
    
    temp.Activate
    With temp
        Dim status As String
        Dim isRngEmpty As Integer
        
        'place the values in the criteria range
        If tktState = 0 Then
            status = ""
        ElseIf tktState = 1 Then
            status = False
        Else
            status = True
        End If

        With tmpSearch
            .Cells(2, 18).Value = startRng
            .Cells(2, 19).Value = endRng
            .Cells(2, 20).Value = tech
            .Cells(2, 21).Value = status
            .Cells(2, 22).Value = rsn
            
            Set critRng = .Range("myCriteria")
            Set resultRng = .Range("copyToRng")
        End With
        With tmpLog
            Set dataRng = .Range("logSearchRng")
        End With
        
        'run AdvancedFilter
        dataRng.AdvancedFilter xlFilterCopy, critRng, resultRng
        
        If Not IsError([searchResults]) Then
            With reportView
                .logLB.RowSource = "searchResults"
                .fndRecordsBx.Value = .logLB.ListCount
            End With
        Else
            MsgBox "No results found! Resetting...", , "Nope."
            reportView.logLB.RowSource = dataRng.Address
            reportView.rsnCboBx.ListIndex = -1
        End If
    End With
End Function

Sub listBoxSort(oLB As MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
    Dim vaItems As Variant
    Dim i As Long, j As Long
    Dim c As Integer
    Dim vTemp As Variant

    'Put the items in a variant array
    vaItems = oLB.List

    'Sort the Array Alphabetically(1)
    If sType = 1 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                'Sort Ascending (1)
                If sDir = 1 Then
                    If vaItems(i, sCol) > vaItems(j, sCol) Then
                        For c = 0 To oLB.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If

                'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If vaItems(i, sCol) < vaItems(j, sCol) Then
                        For c = 0 To oLB.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
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
                        For c = 0 To oLB.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If

                'Sort Descending (2)
                ElseIf sDir = 2 Then
                    If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLB.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
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
    oLB.List = vaItems
End Sub
Sub lbSort(sCol As Integer, sType As Integer, sDir As Integer)
    Dim sortDir As Variant
    Dim colLtr As String
    
    Dim sortRng As Range
    Dim resultsRng As Range
    
    With temp
        With tmpSearch
            Set sortRng = .Range("Search!$A$1:$O$" & lastLogRow)
            Set resultsRng = .Range("Search!$A$2:$O$" & lastLogRow)
        End With
        
        If sDir = 1 Then
            sortDir = xlAscending
        Else
            sortDir = xlDescending
        End If
    
        colLtr = ColNumToLetter(sCol) & "1"
        With tmpSearch.Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range(colLtr), Order:=sortDir
            .SetRange sortRng
            .Header = xlYes
            .Apply
        End With
    End With
    
    With reportView
        .logLB.RowSource = "Search!" & resultsRng.Address
        .fndRecordsBx.Value = .logLB.ListCount
    End With
End Sub

Public Sub reportsRun()
    Dim lastTmpLogRow As Long
    
    lastLogRow = logSht.Cells(logSht.Rows.Count, 2).End(xlUp).row
    
    'refresh the temp workbook
    logSht.Range("A2:O" & lastLogRow).Copy tmpLog.Range("A2:O" & lastLogRow)
    
    'load the search range in case user does a sort before a search
    tmpLog.Range("A2:O" & lastLogRow).Copy tmpSearch.Range("A2:O" & lastLogRow)
End Sub


Public Sub tempXL()
    Dim fileName As String
    Dim folderPath As String
    Dim filePath As String
    
    folderPath = Application.ActiveWorkbook.Path & "\"
    
    fileName = "temp_reportData.xlsx"
    filePath = folderPath & fileName
    
    If Dir(filePath) <> "" Then 'File exists
        Set temp = Workbooks.Open(filePath)
        temp.SaveAs folderPath & fileName
    Else 'file does not exist, create
        Set temp = Workbooks.Add
        temp.SaveAs folderPath & fileName
    End If
    
    'copy needed sheets to temp workbook
    wb.Sheets(Array("Log", "Search")).Copy Before:=temp.Sheets(1)
    
    temp.Windows(1).Visible = False
End Sub

Public Sub unalive()
    If temp Is Nothing Then
        Exit Sub
    Else
        Set temp = Nothing
        Application.Workbooks("temp_reportData.xlsx").Close SaveChanges:=False
    End If
End Sub
