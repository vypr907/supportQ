Attribute VB_Name = "reportsEngine"

'TODO: a function that takes all the potential search options and loads the form accordingly

Public Function logSearch(Optional tech as String, Optional rsn as String, _
Optional startRng as Variant, Optional endRng as Variant)
    'TODO: CODE GOES HERE
    Dim critRng As Range
    Dim dataRng as Range
    Dim resultRng as Range

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
            'If Application.WorksheetFunction.CountA(rng) = 0 Then '=> This is redundant.
            '    MsgBox "Range is blank"
            'End If
        Else
            'MsgBox "No such range" '==> This is practically your black range as you are using dynamic named range.
            MsgBox "No results found! Resetting..."
            reportView.logLB.RowSource = "Log!A2:M" & lastLogRow
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

    With tmpSearch
        Set sortRng = .Range("sortable")
        Set resultsRng = .Range("searchResults")
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
        '.SetRange Range("sortable")
        .SetRange sortRng
        .Header = xlYes
        .Apply
    End With
    temp.Activate  'TODO: see if we can't get rid of the .Activate without breaking shit
    With reportView
        '.logLB.RowSource = tmpSearch.Range("searchResults")
        .logLB.RowSource = resultsRng
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