Attribute VB_Name = "reportsEngine"

'TODO: a function that takes all the potential search options and loads the form accordingly

Public Function logSearch(Optional tech as String, Optional rsn as String, Optional opn as Boolean, _
Optional startRng as Variant, Optional endRng as Variant)
    'TODO: CODE GOES HERE
    'ensure that there is a "user" selected
    k = 0
    With reportView
        .logLB.Clear
        lastLogRow = logSht.Cells(Rows.Count, 1).End(xlUp).Offset(1,0).row
        For rw = 2 to lastLogRow
            If logSht.Range("K" & CStr(rw))= tech Then 'if user's initials are in tech column
                .logLB.AddItem 
                For i = 1 to 13
                    .logLB.List(k,i-1) = logSht.Cells(rw,i)
                Next i
                k = k + 1
            End If
        Next rw
    End With

End Function