Dim wk As Worksheet

Public Sub affecter_un_cout(class_text As String, cout As Long)
    '
    ' Use this module to affect a cost to a specific class of
    ' the PCG.
    '
    top_row = wk.Range("B3").Value
    
    If top_row = "" Then
        wk.Range("B3").Value = class_text
        wk.Range("C3").Value = cout
    Else
        wk.Range("B2").End(xlDown).Offset(1, 0).Value = class_text
        wk.Range("B2").End(xlDown).Offset(0, 1).Value = cout
    End If
End Sub

Public Sub supprimer_un_cout()
    '
    ' Use this module to delete a cost to a specific class of
    ' the PCG.
    '
End Sub

Private Sub Class_Initialize()
    '
    ' Initializes the class
    '
    Set wk = Worksheets("Tableau")
End Sub

'
'
' TEST SECTION
'
'
Public Sub test_affecter_un_cout(ByVal get_row As String, _
                                 ByVal class_text As String, _
                                 ByVal cout As Integer)
    Dim wks As Worksheet
    Set wks = Worksheets("PCG2")
    top_row_charges = wk.Range("B7").Value
    top_row_produits = wk.Range("H7").Value
    If get_row = 6 Then
        Call inscrire_le_cout(top_row_charges, 0)
    End If
    If get_row = 7 Then
        Call inscrire_le_cout(top_row_charges, 1)
    End If
End Sub

Public Sub inscrire_le_cout(ByVal top_row As String, _
                                 Optional ByVal ecriture As Integer = 0)
    Dim r, s As String
    If ecriture = 0 Then
        r = "B7"
        s = "E7"
    End
    If ecriture = 1 Then
        r = "H7"
        s = "K7"
    End If
    If top_row = "" Then
        wk.Range(r).Value = class_text
        wk.Range(s).Value = cout
    Else
        wk.Range(r).End(xlDown).Offset(1, 0).Value = class_text
        wk.Range(s).End(xlDown).Offset(0, 1).Value = cout
    End If
End Sub

