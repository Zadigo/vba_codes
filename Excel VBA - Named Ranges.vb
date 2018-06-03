Sub NameRanges()
    'Option 1:
    Dim k As Workbook
    Set k = Workbooks(ActiveWorkbook.Name)
    
    w = k.names("Kendall").RefersToRange(1, 1)
    
	'Or
	
	'Quick access to values in the named range
	w = Range("kendall")(1, 1)
	
	'Adding a name
	ActiveWorkbook.names.Add Name:="Test", RefersToR1C1:="=Sheet1!A1"
End Sub






Sub reading_multidimensional_Ranges()
  For i = 0 To 4
    For t = 0 To 4
        w = Range("kendall")(i, t)
    Next t
  Next i
End Sub

Sub reading_multidimensional_Ranges()
  For Each w In Range("kendall")
    show = w
  Next
End Sub
