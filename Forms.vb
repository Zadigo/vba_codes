sub manipulating_ObjectInForms ()
	'
	' Works on active viewed form
	
	'Explicit
	Forms!Matches.TourYear = "2012"
	
	'Implicit
	Forms!Matches!TourYear = "2012"
	
	'Access object of a subform
	Forms!Matches.TourYear.ctlSubForm.Forms!...
	
	'Runs a macro after update
	Forms(0).AfterUpdate = "[Macro]"
End Sub