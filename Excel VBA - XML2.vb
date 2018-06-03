Option Explicit

Dim titles_Dic As Scripting.Dictionary

Sub opening_Uploading_File()
    Dim fso As FileSystemObject
    Dim file_Folder As Folder
    Dim file_Files As file
    
    Dim file_Path As String
    file_Path = "C:\Users\Zadig\Documents\Access Databases\"
    
    Set fso = New FileSystemObject
    Set file_Folder = fso.GetFolder(file_Path)
    'Set file_Files = file_Folder.Files("Import.xlsx")
    
    Set titles_Dic = New Scripting.Dictionary
    
    Dim number_OfFiles As Long
    number_OfFiles = file_Folder.Files.Count
    
    'For Each n In file_Folder.Files
        
    'Next
    
    Dim i, r As Long
    Dim n, j, file_Type(0), file_Name_Array(0), file_Name, file_Extension As Variant
    
    For Each j In file_Folder.Files
        file_Type(0) = Split(j, ".")
        file_Extension = file_Type(0)(1)
        
        If file_Extension = "xlsx" Then
            file_Name_Array(0) = Split(j, "\")
            file_Name = file_Name_Array(0)(5)
            
            r = r + 1
            
            If r > 1 Then
                'TO DO
            End If
        End If
    Next
    
    Set file_Files = file_Folder.Files(file_Name)
    
    
    
    set_Dictionnary "Title A", 1
    set_Dictionnary "Title B", 2
    
    Set fso = Nothing
    Set file_Folder = Nothing
    Set file_Files = Nothing
End Sub

Private Sub set_Dictionnary(ByVal t As String, ByVal element_Key As Long)
    titles_Dic.Add element_Key, t
End Sub


