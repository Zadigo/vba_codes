Dim wk_one As Worksheet
Dim wk_two As Worksheet
Dim phrase_one, phrase_five, phrase_seven As String
Dim type_entreprise As String

Private Sub Class_Initialize()
    'Set sheets at initialization
    Call get_Worksheet_Helper
    type_entreprise = wk_two.Range("type_entreprise").Value
    Call set_phrases_Helper
End Sub

Private Sub get_Worksheet_Helper()
    Set wk_one = Worksheets("Analyse")
    Set wk_two = Worksheets("Config")
End Sub

Private Sub set_phrases_Helper()
    Dim c_var, t_var, d_var As String
    '
    ' This helper is used to create the phrases with the variables
    ' that were set or calculated in the 'Config' worksheet
    '
    phrase_one = "L'entreprise fonctionne " & wk_two.Range("jours_activités").Value & _
                 " jours par semaines soit un nombre total de " & _
                 Round(wk_two.Range("semaines_activités").Value, 2) & " semaines."
    '
    ' TO DO
    '
    
    '
    ' I am using this technique in order to get the correct values to display
    ' depending on the fact if the enterprise is a restaurant or bar instead
    ' of a digital based type project
    '
    If type_entreprise = "numérique" Then
        c_var = "ca_numérique"
        t_var = "frequentation_mensuelle"
        d_var = "mois"
    Else
        c_var = "ca_restauration"
        t_var = "frequentation_journalière"
        d_var = "jours"
    End If
    phrase_five = "Pour une fréquentation de " & wk_two.Range(t_var).Value & " clients par " & d_var & ", " & _
                  "le chiffre d'affaire annuel est de " & Round(wk_two.Range(c_var).Value, 2) & "€ par an"
    '
    ' TO DO
    '
    phrase_seven = "Le prix unitaire utilisé pour l'estimation du C.A. est de " & wk_two.Range("N10").Value & "€ soit " & _
                   wk_two.Range("N14").Value & "€ TTC et une marge de " & wk_two.Range("O11").Value & "% (ou " & _
                   wk_two.Range("N11").Value & "€)"
End Sub

Public Sub change_phrases()
    wk_one.Range("B8").Value = phrase_one
    wk_one.Range("B15").Value = phrase_five
    wk_one.Range("B18").Value = phrase_seven
End Sub
