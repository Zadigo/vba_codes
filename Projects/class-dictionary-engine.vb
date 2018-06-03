Dim pcg As Worksheet
Dim pcg_dictionary As Dictionary

Private Sub Class_Initialize()
    UserForm1.ListBox1.Clear
    Call set_Params
    Call create_dictionary
End Sub

Private Sub set_Params()
    '
    ' This sub sets the global parameters for the class.
    '
    Set pcg = Worksheets("PCG")
    Set pcg_dictionary = New Dictionary
End Sub

Private Sub create_dictionary()
    '
    ' This sub is used to initialize the values in the dictionary
    ' when the class is called by outside programs.
    ' It iterates on the PCG sheets in order to collect the values
    ' and integrate them to the dictionary.
    '
    top_row_adress = pcg.Range("A2").Row
    bottom_row_adress = pcg.Range("A2").End(xlDown).Row
    
    For i = top_row_adress To bottom_row_adress
        pcg_dictionary.Add pcg.Range("A" & i).Value, pcg.Range("A" & i).Offset(0, 1).Value
    Next i
End Sub

Public Sub fill_list_box_item(Optional specific_class As Long = 0)
    '
    ' Use this sub to fill the list box with the values for the
    ' user to choose from.
    '
    ' -- Use 0 to get all items from the dictionary
    ' -- Use a class number betweet 1 and 7 to get a specific class
    '
    If specific_class = 0 Then
        For Each dictionary_item In pcg_dictionary.Items
            UserForm1.ListBox1.AddItem dictionary_item
        Next
    End If
    
    If specific_class > 0 Then
        Call regex_engine(specific_class)
    End If
End Sub

Private Sub regex_engine(class_number As Long)
    '
    ' This sub is a helper that tests a dictionary key against
    ' a specific regex pattern and returns the ones that
    ' corresponds to it.
    '
    Dim regex As RegExp
    Set regex = New RegExp
    
    With regex
        .MultiLine = False
        .IgnoreCase = True
        .Global = True
        .pattern = "(^[" & class_number & "]\d+)"
    End With
    
    '
    ' I am using a regex pattern (^[0]\d+) to filter the
    ' different keys of the classes
    '

    For Each dictionary_key In pcg_dictionary.Keys
        If regex.test(dictionary_key) Then
            UserForm1.ListBox1.AddItem pcg_dictionary(dictionary_key)
        End If
    Next
End Sub

