Private Function get_tournament_key(ByVal tour As String) As Variant
    '
    ' This function returns the tournament's key
    ' by searching the tour table Query
    '
    Dim wk As Worksheet
    Set wk = Worksheets("Tournaments")
    Dim t As Range
    Set t = wk.Range("tour_table")
    Dim q(3) As Variant
    
    On Error Resume Next
    r = t.Find(tour, searchorder:=xlRows, MatchCase:=True).Row
    'Use minus one to get the upper row
    q(0) = t(r - 1, 1)
    q(1) = t(r - 1, 3)
    q(2) = t(r - 1, 6)
    q(3) = t(r - 1, 5)
    tour_key = q(0)
    
    If IsEmpty(tour_key) Then
        get_tournament_key = ""
    Else
        get_tournament_key = q
    End If
End Function

Sub complete_tournament_sheet()
    '
    ' Use this sub to get the key of a given
    ' tournament by searching the tour table
    '
    Dim wk As Worksheet
    Set wk = Worksheets("MainTour")
    Dim tournament_range As Range
    Set tournament_range = wk.Range(wk.Range("C1").Offset(1, 0).Address, wk.Range("C1").End(xlDown).Address)
    range_count = tournament_range.Count
    Dim q_array As Variant
    Dim v, h As Long
    v = 1
    h = 0
    '
    ' The for-each loop is used in order to cycle through
    ' all the range
    '
    For Each tour In tournament_range
        q_array = get_tournament_key(tour)
        For Each w In q_array
            '
            ' A second for loop is used to cycle through the
            ' array that was sent back with the tournament
            ' carachteristics
            '
            wk.Range("B1").Offset(v, h).Value = w
            'We move to the next column
            h = h + 1
            'We skip the second column for the third
            If h = 1 Then
                h = h + 1
            End If
        Next
        'We move to the next row
        v = v + 1
        'We reset the column letter to come back
        'to the first column of the row
        h = 0
    Next
End Sub
