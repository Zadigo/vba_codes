Dim wk_couts As Worksheet
Dim wk_compte_exploitation As Worksheet

Sub calculate_fixed_costs()
    '
    ' Use this module to calculate the fixed costs.
    ' It starts from range (P9:x) and goes to the right
    ' with a FOR-function to calculate the fixed costs
    ' for each year.
    '
    Dim r As Range
    Set r = wk_couts.Range("F9", wk_couts.Range("F9").End(xlDown))
    
    Dim results As Long
    
    For i = 0 To 3
        results = add_range_numbers(r.Offset(0, i))
        wk_compte_exploitation.Range("H20").Offset(0, i).Value = results
    Next i
End Sub

Sub calculate_var_costs()
    '
    ' Use this module to calculate the variable costs.
    ' It starts from range (P9:x) and goes to the right
    ' with a FOR-function to calculate the variable costs
    ' for each year.
    '
    Dim r As Range
    Set r = wk_couts.Range("P9", wk_couts.Range("P9").End(xlDown))
    
    Dim results As Long
    
    For i = 0 To 3
        results = add_range_numbers(r.Offset(0, i))
        wk_compte_exploitation.Range("H15").Offset(0, i).Value = results
    Next i
End Sub

' Private Sub iterator_engine(ByVal this_range As Range, ByVal cell_address as String)
'     For i = 0 to 3
'         results = add_range_numbers(this_range.Offset(0, i))
'         wk_compte_exploitation.Range(cell_address).Offset(0, i).Value = results
'     Next
' End Sub

Private Function add_range_numbers(ByVal range_values As Range) As Long
    'Add numbers
    add_range_numbers = Application.WorksheetFunction.Sum(range_values)
End Function

Private Sub Class_Initialize()
    'Set
    Set wk_couts = Worksheets("Couts")
    Set wk_compte_exploitation = Worksheets("Compte d'exploitation")
End Sub
