Private Sub CommandButton1_Click()
    '1. Set sheets
    Dim stock_Sheet As Worksheet
    Set stock_Sheet = Worksheets("Stock")
    Dim output_Sheet As Worksheet
    Set output_Sheet = Worksheets("Output")
    
    '2. Set periods
    Dim periods As Long
    periods = TextBox1.Value
    
    '3. Set initial range
    Dim upper_Bound, lower_Bound As Long
    upper_Bound = 3
    lower_Bound = upper_Bound + periods
        'Initial range
    Dim price_Range As Range
    
    '4. Get the last row of the price list
    Dim last_Row As Long
    last_Row = stock_Sheet.Range("A2").End(xlDown).Row
    
    '4. Now we cycle through the range
    Do
        Set price_Range = Range("A" & upper_Bound, "A" & lower_Bound)
        price_Range.Select
        upper_Bound = upper_Bound + 1
        lower_Bound = lower_Bound + 1
    Loop Until lower_Bound > last_Row
    
    '5. Tell user calculation was successful
    MsgBox "SMA (" & periods & ") on price was successful", vbInformation
End Sub


'
'
'Through module
Dim get_This_Sheet As Worksheet
Dim output_Sheet As Worksheet
Dim periods As Long

Sub call_calculate_SMA() '<<<<<Add variable to receive true or false for when user toggles button "wilder"
    '1.Chear output sheet @get_This_Sheet = 'Stock, Output'
    Set get_This_Sheet = Worksheets(get_Sheet_Name(2))
    get_This_Sheet.Range("A1").CurrentRegion.ClearContents
    
    '2. @calculate_SMA = get_Sheet_Name() = 'Stock, Output'
	'Select Case wilder
	'	Case 0:
			'Call calculate_SMA(get_Sheet_Name(1))
			'msg = 0
		'case 1:
			'Call calculate_SMA(get_Sheet_Name(1))
			'Call calculate_SMA(get_Sheet_Name(2))
			'msg = 1
	'End Select
    Call calculate_SMA(get_Sheet_Name(1))
    Call calculate_SMA(get_Sheet_Name(2))
    
    '8. If user did not choose Wilder
	msg = 0
    
    '8. Tell user calculation was successful
    Dim show_Message As String
    
    '@msg = 0 OR 1
    Select Case msg
        Case 0:
            '@show_Message = ... & @periods = 5, 10, 25... & ...
            show_Message = "SMA (" & periods & ") on price was successful"
        Case 1:
            'Alternate message if Wilder was chosen
            show_Message = "SMA and Wilder (" & periods & ") on price was successful"
    End Select
    
    '@Show_Message = "..."
    MsgBox show_Message, vbInformation
End Sub

Sub calculate_SMA(this_Sheet As String)
    '1. Set sheet @this_Sheet = 'Stock OR Output'
    Set get_This_Sheet = Worksheets(this_Sheet)
	'EXPLICIT 'Output'
    Set output_Sheet = Worksheets("Output")
    
    '2. Set periods
    '@periods = 5, 10, 25...
	'Dim periods As Long
	'periods = TextBox1.Value
	periods = 2
    
    '3. Set variables for initial range
    Dim upper_Bound, lower_Bound As Long
    
    'When the incoming sheet name is stock...
    If this_Sheet = "Stock" Then
		'Upper bound of range is 3
        upper_Bound = 3
    Else
        upper_Bound = 1
    End If
    
    '@lower_Bound = 3 + periods
    lower_Bound = upper_Bound + periods
        'Create initial range
    Dim price_Range As Range
    
    '4. Get the last row of the price list
    Dim last_Row As Long
    '@last_Row = 1, 2, 3...x
    last_Row = get_This_Sheet.Range("A2").End(xlDown).Row
    
    '5. When there are no values in input sheet...
    If last_Row > 10000 Then
        '@last_Row = last row of 'get_This_Sheet stock values'
        last_Row = get_This_Sheet.Range("A1").End(xlDown).Row - periods
    End If
    
    '6. Set variable for average
    Dim avg As Double
    
    '7. Set variable to cycle in 'Output sheet'
    Dim r, c As Long
	'
	'r: row, c: column
	'
    r = 1
    c = 1
    
    '8. When the incoming sheet is 'Stock'
    If this_Sheet = "Stock" Then
        c = 1
    Else
        'Output to next column
        c = 2
    End If
    
    '9. Now we cycle through the range
    Do
        Set price_Range = get_This_Sheet.Range("A" & upper_Bound, "A" & lower_Bound)
        
        'Calculate average range
        avg = Application.WorksheetFunction.Average(price_Range)
        
        'Output to 'Output sheet' @r = 1++ / @c = 1 || 2
        output_Sheet.Range("A1").Cells(r, c).Value = avg
        
        'Increment
        r = r + 1
        upper_Bound = upper_Bound + 1
        lower_Bound = lower_Bound + 1
    Loop Until lower_Bound > last_Row
End Sub

Private Function get_Sheet_Name(index As Long) As String
    '1. When the user chooses 'Stock'
	'@index = 1 OR 2
    Select Case index
        Case 1:
            get_Sheet_Name = "Stock"
        Case 2:
            get_Sheet_Name = "Output"
    End Select
End Function
