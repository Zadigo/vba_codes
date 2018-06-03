Dim wk As Worksheet
Dim static_cells(7) As String

Sub quick_increase_by_percentage()
    '
    ' You can use this module to increase the value of a cell
    ' by x% quickly gor testing hypothesis
    '
    Dim active_cell_value As Variant
    Dim active_cell_address As String
    active_cell_value = ActiveCell.Value
    
    If IsNull(active_cell_value) Or IsEmpty(active_cell_value) Then
        Exit Sub
    End If
    
    'We have to make sure that we are
    'dealing with a number
    If IsNumeric(active_cell_value) Then
        active_cell_address = ActiveCell.Address
        For i = 0 To 7
            If active_cell_address = static_cells(i) Then
                Exit Sub
            End If
        Next i
        'Make percentage dynamic
        active_cell_value = active_cell_value * 1.3
        ActiveCell.Value = active_cell_value
    End If
End Sub

Public Sub increase_whole_line_percentage()
    '
    ' Use this sub to increase a whole given line
    ' from activecell by a given percentage
    '
    Set wk = Worksheets("Couts")
    Dim active_cell_column As Long
    Dim active_range_address As String
    Dim value_range As Range
    active_cell_column = ActiveCell.Column
    
    Set value_range = Selection
    lgt = value_range.Count
    
    active_range_address = value_range.Address
    
    If active_range_address = "$F$6:$I$6" Or active_range_address = "$P$6:$S$6" Then
        Exit Sub
    End If
    
    For i = 1 To lgt
        If Not IsNumeric(value_range(1, i)) Or IsEmpty(value_range(1, i)) Then
            n = 1
        End If
    Next i
    
    If n = 1 Then
        Exit Sub
    End If
    
    For i = 1 To lgt
       value_range(1, i) = Round(value_range(1, i) * 1.2, 2)
    Next i
End Sub

Private Sub Class_Initialize()
    Set wk = Worksheets("Couts")
    
    'Static cells are used to prevent the user
    'from increasing values of unwanted cells
    static_cells(0) = "$F$6"
    static_cells(1) = "$G$6"
    static_cells(2) = "$H$6"
    static_cells(3) = "$I$6"
    static_cells(4) = "$P$6"
    static_cells(5) = "$Q$6"
    static_cells(6) = "$R$6"
    static_cells(7) = "$S$6"
End Sub
