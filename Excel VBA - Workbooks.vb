Sub workbook_Open()
    'Set variables
	Dim that_Workbook As Workbook
    Dim current_Path As String
    
	'Set path
    current_Path = ThisWorkbook.Path & "\Player Table.xlsx"
    
    'Open workbook
    Workbooks.Open current_Path
	
	'
	'
	'
	'
	'*Create object
    'Dim opened_Workbook As String
    'opened_Workbook = Workbooks(1).Name
	
	'Set that_Workbook = Workbooks(opened_Workbook)
	
	'that_Workbook.Activate
End Sub

Sub workbook_Close()
    'With saving at close
    Workbooks("Kendall.xltm").Close True
	
	'Workbooks("Kendall.xltm").Close
    
    'ActiveWorkbooks.Close
    
    'ThisWorkbook.Close
End Sub

Sub workbook_Create()
    Workbooks.Add
	
	'Option:
    'Workbooks.Add "Name"
	
	'@saveAs = xlCSV, xlOpenDocumentSpreadsheet, xlWorkbookDefault, xlXMLSpreadsheet
	Dim current_path As String
    current_path = ActiveWorkbook.Path
    
    Workbooks.Add
    
    ActiveWorkbook.SaveAs current_path & "\kendall.xlsx"
	
	Workbooks(2).close
End Sub