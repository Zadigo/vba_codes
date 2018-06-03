Sub create_NewButton()
	'Set parameter
    Dim btn As Control
    
	'Open form in hidden mode
    DoCmd.OpenForm "Google", acDesign, , , acFormEdit, acHidden
    
    On Error Resume Next
    
	'Create button
    Set btn = CreateControl("Google", acCommandButton, acDetail)
    
	'Move
    k.Move 2500, 2500, 1500, 700
    
	'Get control name
    this_name = k.Name
    
	'Add caption
    Forms("Google").Controls(this_name).Caption = "Google"
    
	'Close form
    DoCmd.Close acForm, "Google", acSaveYes
End Sub

Sub create_NewButtons()
	'Set array
    Dim btn(0 To 1) As Control
	'Form & control
    Dim o As Form
	Dim f As Control
	Dim leftMove As Long
	leftMove = 2000
    
	'Open form hidden
    DoCmd.OpenForm "Google", acDesign, , , acFormEdit, acHidden
    
    Set o = Forms("Google")
	
    For i = 0 To 1
		'Create buttons
		Set btn(i) = CreateControl("Google", acCommandButton, acDetail)
		
		'Set object to control
        Set f = Forms("Google").Controls(i)
		'Move controls
        f.Move 1000, leftMove
        
		'When control is index 1, put x name and caption
        If i = 0 Then
            f.Name = "Email_Button"
            f.Caption = "Email"
        Else
            f.Name = "Validate_Button"
            f.Caption = "Validate"
        End If
		
		'Move below
        leftMove = leftMove + 1000
    Next i
    
	'Close form
    DoCmd.Close acForm, "Google", acSaveYes
End Sub