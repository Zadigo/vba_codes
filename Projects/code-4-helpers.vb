'
' TO DO
' Créer un tableau de bords où les coûts se classe par section afin permettre le coût par section
' /// Tester si c'est possible de retrouver la clée à partir d'une valeur dans un dictionnaire
' Créer un panneau envoyer email à [...] pour permettre la proposition de suggestion et de signalement de problème
' Créer un programme python qui puisse permettre la création d'un LOG FILE
' Permettre l'export en PDF des tableaux
' Relier le PCG à un ODBC ?????? OU créer un fichier PCG online qui se mette à jour automatiquement pour toutes les copies et si des personnes veulent des classes spécifiques
'

Sub open_list_box()
    UserForm1.Show
End Sub

Public Sub ouvrir_reglage()
    ReglagesForm.Show
End Sub

Public Sub changer_les_phrases_helper()
    '
    ' This helper creates an object in order to change the phrases
    ' on the front page
    '
    Dim changer_les_phrases As PhrasesEngine
    Set changer_les_phrases = New PhrasesEngine
    changer_les_phrases.change_phrases
End Sub

Public Function set_worksheet_object_helper(Optional wk_number As Long = 0) As Worksheet
    '
    ' This helper sets itself to the worksheet object and
    ' returns it to the calling function
    '
    Select Case wk_number
        Case 0:
            Set set_worksheet_object_helper = Worksheets("Analyse")
        Case 1:
            Set set_worksheet_object_helper = Worksheets("Config")
        Case Else:
            Err.Raise 0
    End Select
End Function

Public Sub regex_engine(ByVal regex_pattern As String, _
                        ByVal string_to_test As String, _
                        Optional ByVal class_number As Long = 0, _
                        Optional ByVal this_dictionary As Dictionary, _
                        Optional ByVal from_dictionary As Boolean = False)
    '
    ' This sub is a helper that tests either a dictionary key against
    ' a specific regex pattern and returns the ones that
    ' corresponds to it /OR/ tests a unique specific string combined
    ' with to a given pattern.
    '
    ' This sub was developed in order to create a dynamix REGEX sub
    ' for all the functions or classes created withint this worksheet.
    '
    Dim regex As RegExp
    Set regex = New RegExp
    
    With regex
        .MultiLine = False
        .IgnoreCase = True
        .Global = True
        .pattern = regex_pattern
    End With
    
    '
    ' We have to make sure that a dictionary would have been received
    ' otherwise, consider this a regular REGEX string test sub.
    '
    If from_dictionary = True And Not this_dictionary Is Nothing Then
        'Specific for loop created for the UserForm in order to avoid
        'calling this function numerous times and optimize memory
        For Each dictionary_key In this_dictionary.Keys
            If regex.test(dictionary_key) Then
                UserForm1.ListBox1.AddItem this_dictionary(dictionary_key)
            End If
        Next
    ElseIf from_dictionary = True And this_dictionary Is Nothing Then
        MsgBox "There was no dictionary", vbCritical, "No dictionary"
        Exit Sub
    ElseIf from_dictionary = False Then
        If Not regex.test(string_to_test) Then
            ' TO DO
            ' This should call an error in a class ErrorDic created for
            ' specific custom errors
            MsgBox "Veuillez faire x...", vbInformation, "Title"
            Exit Sub
        End If
    End If
End Sub

Private Sub show_messages(error_number As Long, message As String, message_type As String)
    '
    ' Use this sub to cycle through the error dictionary and show the
    ' user any custom error message that is not built in Excel.
    '
    ' vbInformation, vbCritical, vbOkOnly, vbYesNo, vbYesNoCancel
    '
    MsgBox message, message_type, "Title"
End Sub

'
' Test
'
'

Sub test()
    regex_engine "[a-b]", "Kendall", from_dictionary:=True
End Sub
