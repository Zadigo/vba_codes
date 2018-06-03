Sub opening_Uploading_File()
    '
    ' This asks the user to upload a specific file
    '
    
    Dim user_Selected_File As Variant
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "All Excel Files", "*.xlsx"
        .Show
        'Gets the path of the file / Dir() gets the name in the path
        user_Selected_File = Dir(.SelectedItems(1), vbNormal)
    End With
    
	msgbox user_Selected_File
	
    'Dim k(0) As Variant
    'k(0) = Split(user_Selected_File, ".")
    
    'user_Selected_File = k(0)(0)
End Sub

Public Sub renaming_File(ByVal file_Name As String)
    '
    ' This gives a file a new name
    '
    
    Dim fso As FileSystemObject
    Dim fso_Folder As Folder
    Dim fso_File As file
    Dim new_File_Name As String
    
    Set fso = New FileSystemObject
    Set fso_Folder = fso.GetFolder("C:\Users\Zadig\Documents\Access Databases\")
    Set fso_File = fso_Folder.Files("Import.xlsx")
    
    'Set new file name
    new_File_Name = "Imports_" & Year(Date) & ".xlsx"
    
    'When not "imports", change name
    If file_Name <> "Imports" Then
        fso_File.Name = new_File_Name
    End If
    
    Set fso = Nothing
    Set fso_File = Nothing
    Set fso_Folder = Nothing
End Sub
