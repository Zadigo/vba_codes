Dim wk As Worksheet

Private Sub CommandButton1_Click()
    wk.Range("cout_achat").Value = (TextBox1.Value * 1)
    wk.Range("marge").Value = (TextBox2.Value * 1)
    Call changer_les_phrases_helper
End Sub

Private Sub set_Params()
    '
    ' This a helper used to set parameters to
    ' the textboxes
    '
    TextBox1.Value = wk.Range("cout_achat").Value
    TextBox2.Value = wk.Range("marge").Value
End Sub

Private Sub UserForm_Initialize()
    Set wk = Worksheets("Config")
    Call set_Params
End Sub
