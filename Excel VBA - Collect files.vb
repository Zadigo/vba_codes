Sub collectWorkbooks()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object
    Dim fullFilePath As String
    
    ' Get the active workbook path
    Dim wbPath As String
    wbPath = wb.Path
    ' Get the folder of the current workbook
    Set folder = fso.GetFolder(wbPath)
    
    Dim openedWb As Workbook
    Dim currentSheet As Worksheet
    
    Dim newWs As Worksheet
    
    Dim sourceRange As Range
    Dim destinationRange As Range
    
    Dim currentWs As Worksheet
    
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    For Each file In folder.Files
        If InStr(1, file.Name, "Rapport", vbTextCompare) = 0 Then
            ' Build the worksheet path
            fullFilePath = openedWsPath & "\" & file.Name
            ' Open the worksheet
            Set openedWb = Workbooks.Open(file.Name)
            ' Get the first sheet
            Set currentWs = openedWb.Worksheets(1)
            Set sourceRange = currentWs.Range("A1:N300")
            
            ' Create a new sheet in this current workbook
            wb.Sheets.Add
            Set newWs = wb.Sheets(1)
            ' newWs.Select
            newWs.Name = Replace(file.Name, ".csv", "")

            ' Copy the data from the source range to the destination range
            Set destinationRange = newWs.Range("A1:N300")
            sourceRange.Copy destinationRange

            ' Close and release the opened worksheet
            openedWb.Close SaveChanges:=False
            Set openedWb = Nothing
            Set newWs = Nothing
            Set sourceRange = Nothing
        End If
    Next file
    
    Application.ScreenUpdating = originalScreenUpdating
End Sub


