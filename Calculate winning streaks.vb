'Calculate winning streaks
Sub google()
    Dim r As Range
    Set r = Range("B2:E5")
    
    a = r.Count
    t = 0
    o = 0
    For i = 0 To a
        p = r(i)
        If r(i) = "W" Then
            t = t + 1
            If t > o Then
                o = t
            End If
        Else
            t = 0
        End If
    Next i
    MsgBox "Longest winning streak: " & o
End Sub