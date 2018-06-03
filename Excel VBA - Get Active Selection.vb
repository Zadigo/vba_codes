Option Explicit

Sub get_ActiveShapeSelected()
	'
	' This gets elements on the current shape selected by the user
	'

    'Set sheet
    Dim current_Sheet As Worksheet
    Set current_Sheet = Worksheets(ActiveSheet.Index)
    
    'Set variables for shape
    Dim activeShape As Shape
    Dim userSelection As Variant
        
    'Get user selection
    Set userSelection = ActiveWindow.Selection
    
    'Error handling
    On Error GoTo noShapeSelection
	'On Error Resume Next
	
	'Set shape
    Set activeShape = ActiveSheet.Shapes(userSelection.Name)
	
	' Assigning VBA code to shape
	'activeShape.OnAction = "'" & ActiveSheet.Index & "'!google" 'google is macro name
    
noShapeSelection:
    Exit Sub
End Sub
