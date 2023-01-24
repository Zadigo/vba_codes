Public Sub concatenate_worksheets()
    Dim wk As Workbook
    Set wk = ThisWorkbook

    ' Open the current workbook
    Dim database As Worksheet
    Set database = wk.Worksheets("Base de données")
    
    ' Get the its current path
    Dim current_path As String
    current_path = wk.Path    
    Dim file_name As String

    ' Open the workbook to analyze
    Dim opened_workbook As Workbook
    file_to_open = current_path & "\" & "google" & ".xlsx"
    Set opened_workbook = Application.Workbooks.Open(file_to_open)
    
    ' Ensure that we have all expected fields
    Dim opened_worbook_sheet As Worksheet
    Set opened_worbook_sheet = opened_workbook.Worksheets(1)
    opened_worbook_sheet.Activate
    
    ' Get the file header and check that
    ' we have the expected fields
    Dim opened_worbook_sheet_values As Range
    Set opened_worbook_sheet_values = opened_worbook_sheet.Range("A1:" & opened_worbook_sheet.Range("A1").End(xlToRight).End(xlDown).Address)
    
    ' Dim expected_columns As Object
    ' Set array_list = CreateObject("System.Collections.ArrayList")
    ' array_list.Add "nom"
    
    opened_worbook_sheet_values.Select
    
    ' For i = i To opened_worbook_sheet_values.Count
    
    ' Next i

    opened_worbook_sheet_values.Copy database.Range("C16")
    
    ' Since the copy also removes the
    ' initial styling, we have to
    ' reapply everything
    Dim copied_items As Range
    Set copied_items = database.Range("A10:F20")
    
    copied_items.Interior.Color = RGB(255, 255, 255)
    copied_items.Font.Name = "Source Sans Pro Light"
    copied_items.HorizontalAlignment = xlCenter
    opened_workbook.Close
End Sub
