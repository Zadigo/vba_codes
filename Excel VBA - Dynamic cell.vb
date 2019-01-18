Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim current_cell As Range
    Dim current_cell_address As String
    Dim previous_cell As Range
    
    'Store the previous cell address
    Set previous_cell = Range("A1")
    
    current_cell_address = ActiveCell.Address
    previous_cell.Value = current_cell_address
    
    Set current_cell = Range(current_cell_address)
    
    'current_cell.Cells.Interior.Color = RGB(43, 42, 78)
    'current_cell_value = current_cell.Value
    'current_cell.FormulaLocal = "=SUM(A1:A1)"
    'current_cell.Copy previous_cell
End Sub
