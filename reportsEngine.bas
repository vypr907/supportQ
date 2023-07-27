Attribute VB_Name = "reportsEngine"

'TODO: a function that takes all the potential search options and loads the form accordingly

Public Function logSearch(Optional tech as String, Optional rsn as String, _
Optional startRng as Variant, Optional endRng as Variant)
    'TODO: CODE GOES HERE
    Dim status as String
    Dim isRngEmpty as Integer
    'Dim rng As Range
    'place the values in the criteria range
    If tktState = 0 Then
        status = ""
    ElseIf tktState = 1 Then
        status = False
    Else
        status = True
    End If

    With searchSht
        .Cells(2,18).Value = startRng
        .Cells(2,19).Value = endRng
        .Cells(2,20).Value = tech
        .Cells(2,21).Value = status
        .Cells(2,22).Value = rsn
    End With

    Dim critRng as Range
    Set critRng = Range("myCriteria")
    Dim dataRng as Range
    Set dataRng = Range("logSearchRng")
    Dim resultRng as Range
    Set resultRng = Range("copyToRng")

    'run AdvancedFilter
    dataRng.AdvancedFilter xlFilterCopy, critRng, resultRng
    
    If Not IsError([searchResults]) Then
        'Set rng = [searchResults]
        reportView.logLB.RowSource = "searchResults"
        'If Application.WorksheetFunction.CountA(rng) = 0 Then '=> This is redundant.
        '    MsgBox "Range is blank"
        'End If
    Else
        'MsgBox "No such range" '==> This is practically your black range as you are using dynamic named range.
        MsgBox "No results found! Resetting..."
        reportView.logLB.RowSource = "Log!A2:M" & lastLogRow
        reportView.rsnCboBx.ListIndex = -1
    End If
    'Set rng = [searchResults]
    'If Application.WorksheetFunction.CountA(rng) = 0 Then
    '    MsgBox "Range is blank!"
    'End If
    'isRngEmpty = Application.WorksheetFunction.CountA(Range("searchResults"))
    'If isRngEmpty > 0 Then
    '    reportView.logLB.RowSource = "searchResults"
    'Else  
    '    MsgBox "No results found! Resetting..."
    '    reportView.logLB.RowSource = "Log!A2:M" & lastLogRow
    'End If

End Function

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