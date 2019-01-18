Sub create_new_workbook()
    Application.ScreenUpdating = False
    
    Dim wb As Workbook
    Dim wk As Worksheet
    
    'Create workbook
    Set wb = Workbooks.Add
    Set wk = wb.Worksheets(ActiveSheet.Name)
    
    'Do something with the sheets
    wk.Name = "Kendall Jenner"
    wk.Tab.Color = RGB(205, 92, 92)
    wk.SaveAs Filename:="C:\Users\Pende\Downloads\" + "test.xlsx"
    
    'Close workbook
    ActiveWorkbook.Close
    
    Application.ScreenUpdating = True
End Sub
