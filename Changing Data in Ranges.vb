Sub quickly_Change_Data_inRange()
    Dim googleCar As Range
    Set googleCar = Range("B2").CurrentRegion
    
    For Each Tesla In googleCar
        Tesla.Value = 20
    Next
End Sub

Sub change_Data_inColumn()
    Dim googleCar As Range
    Set googleCar = Range("B2").CurrentRegion
    
    For Each Tesla In googleCar.Columns(1)
        Tesla.Value = 15
    Next
End Sub

'Instead of For...Each just :
	'googleCar.Value = ...
	
Sub change_specific_Data_inRange()
    Dim googleCar As Range
    Set googleCar = Range("B2").CurrentRegion
    
    For Each Tesla In googleCar
        If Tesla = 5 Then
            googleCar.Value = 10
        End If
    Next
End Sub

