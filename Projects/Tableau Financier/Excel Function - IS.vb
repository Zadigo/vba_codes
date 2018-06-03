Public Function CALCIS(ByVal profit As Double) As Double
    '
    'Cette fonction permet de calculer l'IS
    '
    If profit < 0 Then
        CALCIS = 0
    End If
    Dim intermediate_calculation As Double
    intermediate_calculation = 38.12 * 12 / 12
    If profit < intermediate_calculation Then
        'Si le c.a est < à 38200€, 15%
        CALCIS = result * 0.15
    Else
        'Sinon, (15% de 38200€) + (bénéfice - 38200€) * 33%
        CALCIS = (38.2 * 0.15) + (profit - 38.2) * 0.3333
    End If
End Function