Option Explicit

Sub assign_MacroToShape()
	'
	' This macro assigns a macro to a given shape
	'

	'Set variables
    Dim active_Book As Workbook
    Set active_Book = ActiveWorkbook
    
	Dim active_Sheet As Worksheet
    Set active_Sheet = Worksheets(ActiveSheet.Index)
    
	Dim assignMacro as String
    assignMacro = "'" & activeBook.Name & "'!fashion"
    
	'Assign macro
    u.Shapes.Item(1).OnAction = assignMacro
End Sub

Sub fashion()
	'TO DO
End Sub