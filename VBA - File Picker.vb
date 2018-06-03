Sub file_Picker ()
	'
    ' Imports table / Microsoft Office Object Library 16.0
    '
    
    'Dim
    Dim table_Name(1) As Variant
    Dim sheet_Path As String
    
    'Open file picker
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    'Open file picker
    With fd
        .AllowMultiSelect = False
        .title = "Select a file"
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx"
        
        'When user has picked file
        If .show = True Then
            'Path
            sheet_Path = fd.SelectedItems.Item(1)
            'Name
            table_Name(0) = Dir(fd.SelectedItems.Item(1))
        Else
            'TO DO
        End If
    End With
End Sub