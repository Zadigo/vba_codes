Sub Adding_Text_toShape()
	'Set element to put in shape'
    Dim element As String
    element = ""
	
	'Select shape'
    'ActiveSheet.Shapes.Range(Array("Name of the shape")).Select
	
	'Place element in shape'
	'--> Chr(13) creates new line'
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Fact" & Chr(13) & element
End Sub

Sub Moving_aShape()
	'Example of moving shape to left'
    Selection.ShapeRange.IncrementLeft -372
	
	'Example of moving shape to right'
    Selection.ShapeRange.IncrementTop -39.75
End Sub

Sub Adding_aChart()
	'Selecting chart to include'
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
	
	'Choose source range'
    ActiveChart.SetSourceData Source:=Range("Stages!$B$8:$G$32")
	
	'Move chart'
    ActiveSheet.Shapes("Chart 4").IncrementLeft -210
	
	'Resize chart'
    ActiveSheet.Shapes("Chart 4").ScaleWidth 1.1604166667, msoFalse, msoScaleFromTopLeft
End Sub

Sub InsertTextBox()
	'Insert text box'
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 156.75, 148.5, 223.5, 82.5).Select
	
	'Insert text in box'
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = ""
End Sub


Sub Widden_Column()
	Columns("F:F").ColumnWidth = 15.88
End Sub

Sub InsertPicture()
	'Inserting a picture'
    ActiveSheet.Pictures.Insert("C:\Users\Zadig\Pictures\9kfvQD.jpg").Select
	
	'Change picture name'
	Selection.ShapeRange.Name = "Give a name"
	
	'Delete picture'
    Selection.Delete
End Sub