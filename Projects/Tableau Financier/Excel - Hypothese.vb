Dim wk_tdb As Worksheet
Dim wk_exploitation As Worksheet

Public Sub implement_revenues()
    '
    ' Use this module to implement a an hypothesis created on
    ' revenues into any cell in the cash flow
    '
    Dim current_cell_col, current_cell_row As Long
    current_cell_col = ActiveCell.Column
    current_cell_row = ActiveCell.Row
    '
    'We have to make sure that the change the user is trying to
    'make is containe within the revenue range ($H$10:$K$13)
    '
    If current_cell_col >= 8 And current_cell_col <= 11 Then
        If current_cell_row >= 10 And current_cell_row <= 13 Then
            ActiveCell.Value = wk_tdb.Range("chiffre_affaire_ht_millier")
        End If
    End If
End Sub

Private Sub Class_Initialize()
    'Set
    Set wk_tdb = Worksheets("TDB")
    Set wk_exploitation = Worksheets("Compte d'exploitation")
End Sub
