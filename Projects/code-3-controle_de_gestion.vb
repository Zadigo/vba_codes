Dim wk As Worksheet

Private Sub CommandButton1_Click()
    '
    ' When the user clicks the button, we need to know
    ' which value needs to be implemented from which
    ' control element
    '
    result = get_valeur_de_Optionbutton(0)
    If result = False Then
        '
        ' We need to test the value that the user entered
        ' to not get errors on the sheet
        '
        wk.Range("jours_activités").Value = (TextBox1.Value * 1)
    Else
        get_valeur_list = ComboBox1.Value
        Select Case get_valeur_list
            Case Is = "Normal":
                wk.Range("jours_activités").Value = 365
            Case Is = "Business":
                wk.Range("jours_activités").Value = 251
            Case Is = "Jours ouvrées, Sans Weekend, Avec vacances":
                wk.Range("jours_activités").Value = 260
        End Select
    End If
    
    '
    ' Create object to change phrases
    '
    Dim changer_les_phrases As PhrasesEngine
    Set changer_les_phrases = New PhrasesEngine
    changer_les_phrases.change_phrases
End Sub

Private Sub OptionButton1_Change()
    '
    ' We detect change on the option button number 1
    ' and if the user changed to 'personalized',
    ' we enable the corresponding elements for him/her
    '
    result = get_valeur_de_Optionbutton(0)
    If result = False Then
        TextBox1.Enabled = True
        TextBox1.Value = 365
        ComboBox1.Enabled = False
    Else
        TextBox1.Enabled = False
        TextBox1.Value = ""
        ComboBox1.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim weeks(2) As String
    Set wk = Worksheets("Config")
    'Enable system radio button as true
    OptionButton1.Value = True
    result = get_valeur_de_Optionbutton(0)
    If result Then
        TextBox1.Enabled = False
    End If
    
    'Initializing the combo box
    weeks(0) = "Normal"
    weeks(1) = "Business"
    weeks(2) = "Jours ouvrées, Sans Weekend, Avec vacances"
    
    For Each week In weeks
        ComboBox1.AddItem week
    Next
End Sub

Private Function get_valeur_de_Optionbutton(Optional ByVal option_button As Long = 0) As Boolean
    '
    ' This is a helper that retrieves the current value
    ' of the option buttons
    '
    If option_button = 0 Then
        get_valeur_de_Optionbutton = OptionButton1.Value
    ElseIf option_button = 1 Then
        get_valeur_de_Optionbutton = OptionButton1.Value
    Else
        Err.Raise 0
    End If
End Function

Private Sub UserForm_Terminate()
    ' DOES NOT WORK
    ' This is to set the textbox to the new value
    ReglagesForm.TextBox5.Value = wk.Range("jours_par_semaines").Value
End Sub
