Dim wk As Worksheet

Private Sub CommandButton1_Click()
    'Insert new params
    wk.Range("personnel").Value = TextBox1.Value
    wk.Range("salaire").Value = TextBox2.Value
    wk.Range("jours_activités").Value = TextBox3.Value
    wk.Range("heures_par_jours").Value = TextBox4.Value
    wk.Range("jours_par_semaines").Value = TextBox5.Value
    Call changer_les_phrases_helper
End Sub

Private Sub CommandButton2_Click()
    '
    ' This will show an additional form to the user as
    ' a way to form him/her to not change the weeks directly
    ' on the main form
    '
    SemainesForm.Show
End Sub

Private Sub CommandButton3_Click()
    PrixForm.Show
End Sub

Private Sub UserForm_Initialize()
    '
    ' This will upload the default values from the sheet
    ' to the different textboxes
    '
    Set wk = set_worksheet_object_helper(1)
    'Set params
    Call set_Params
End Sub

Private Sub set_Params()
    '
    ' This a helper used to set parameters to
    ' the textboxes
    '
    TextBox1.Value = wk.Range("personnel").Value
    TextBox2.Value = wk.Range("salaire").Value
    TextBox3.Value = wk.Range("jours_activités").Value
    TextBox4.Value = wk.Range("heures_par_jours").Value
    TextBox5.Value = wk.Range("jours_par_semaines").Value
    
    '
    ' Prevent direct change of activity days froms userform
    '
    TextBox3.Locked = True
    TextBox3.Enabled = False
End Sub
