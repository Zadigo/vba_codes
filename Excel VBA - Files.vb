http://software-solutions-online.com/list-files-and-folders-in-a-directory/
https://stackoverflow.com/questions/31414106/get-list-of-excel-files-in-a-folder-using-vba

Activate 'Microsoft Scripting Runtime'

Option Explicit

Sub get_eachFilesInFolder()
	'
	' This macro gets each files in a given folder
	'
	
    Dim o As FileSystemObject
    Dim p As Folder
    Dim f As file
    Dim a As Variant
    
    Set o = New FileSystemObject
    Set p = o.GetFolder("D:\Applications")
    
    For Each a In p.Files
        Debug.Print a
		
		'Or
		
		'Get names of files
		Debug.Print Dir(a, vbNormal)
    Next
	
	set f = Nothing
    set o = Nothing
	set p = Nothing
End Sub

Sub get_aSpecificFile()
	'
	' This macro gets each files in a given folder
	'
	
    Dim o As FileSystemObject
    Dim p As Folder
    Dim f As file
    Dim a As Variant
	
	Dim file_Path as String
    file_Path = Workbooks(ActiveBook.Path).Path
	
    Set o = New FileSystemObject
    Set p = o.GetFolder("D:\Applications")
	
	Set f = p.Files("...")
	
	'TO DO
    
	set f = Nothing
    set o = Nothing
	set p = Nothing
End Sub

Sub get_eachFilesInFolder()
    '
    ' Other method
    '
    
    Dim o As FileSystemObject
    Dim p As Folder
    Dim h As Files
    Dim f As file
    Dim a As Variant
    
    Dim open_File As String
    open_File = "Comptage" & ".xlsm"
    
    Set o = New FileSystemObject
    Set p = o.GetFolder("C:\Users\Zadig\Documents")
    Set h = p.Files
    Set f = h.Item("...")
    
    'TO DO
    
    Set o = Nothing
    Set p = Nothing
    Set f = Nothing
End Sub






'
' Import each Excel files from a folder to access
'

Sub ImportfromPath(path As String, intoTable As String, hasHeader As Boolean)
	Dim fileName As String

	'Loop through the folder & import each file
	fileName = Dir(path & "\*.xls")
	While fileName <> ""
		DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, intoTable, path & fileName, hasHeader
	   'check whether there are any more files to import
		fileName = Dir()
	Wend
End Sub