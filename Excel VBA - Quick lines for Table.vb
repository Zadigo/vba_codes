Sub create_Table_Borders()
    'Set range for borders
	Dim j As Range
    Set j = Range("B2:D10")
    
	'Store areas
    Dim elements(5) As String
    elements(0) = xlEdgeBottom
    elements(1) = xlEdgeTop
    elements(2) = xlEdgeLeft
    elements(3) = xlEdgeRight
    elements(4) = xlInsideHorizontal
    elements(5) = xlInsideVertical
    
	'Implement
    For Each w In elements
        j.Cells.Borders(w).LineStyle = xlContinuous
        j.Cells.Borders(w).Weight = xlThin
		
		'
		' Optional
		'
		
		'Medium weight borders around table
		If w = xlEdgeTop Or w = xlEdgeBottom Or w = xlEdgeLeft Or w = xlEdgeRight Then
            j.Cells.Borders(w).Weight = xlMedium
        End If
    Next
End Sub

Sub create_Table_BordersAnywhere()
	'Get current range selected
    Dim current_Selection As Range
    Set current_Selection = Selection
    
	'Set in range
    Dim create_Table As Range
    Set create_Table = Range(current_Selection.Address)
    
    Dim elements(5) As String
    elements(0) = xlEdgeBottom
    elements(1) = xlEdgeTop
    elements(2) = xlEdgeLeft
    elements(3) = xlEdgeRight
    elements(4) = xlInsideHorizontal
    elements(5) = xlInsideVertical
    
    For Each w In elements
        j.Cells.Borders(w).LineStyle = xlContinuous
        j.Cells.Borders(w).Weight = xlThin
        
        If w = xlEdgeTop Or w = xlEdgeBottom Or w = xlEdgeLeft Or w = xlEdgeRight Then
            j.Cells.Borders(w).Weight = xlMedium
        End If
    Next
End Sub
