Sub create_NewForm()
	'Set parameters
    Dim j As Form
    Set j = CreateForm
	'Get current form name
    y = Application.CurrentObjectName
	'Set variable to object
    Set j = Forms(y)
	'
	' Error handling
	'
    On Error Resume Next
	'Set recordsource to table...
	j.RecordSource = "Facebook"
	'Close form
    DoCmd.Close acForm, "Form2", acSaveYes
	'Rename form
    DoCmd.Rename "Google", acForm, y
End Sub

Sub manipulate_Form()
	'Open form
    DoCmd.OpenForm "Form1", acDesign, , , acFormEdit, acHidden
	'Get form name
    w = Forms![Form1].Command0.Name
	'Close form
    DoCmd.Close acForm, "Form1", acSaveNo
End Sub

' Apply a filter to forms
' DoCmd.ApplyFilter , "Filiere = 'Bac'"
' DoCmd.ApplyFilter , "Filiere = 'Bac' AND Niveau = 'Master'"
