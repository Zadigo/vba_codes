Sub write_Text()
	'Set document
    Dim current_Document As Document
    Set current_Document = ActiveDocument
    
	'Get document path and append text file name
    Dim text_Document_Path As String
    text_Document_Path = ThisDocument.Path & "\DB.txt"
    
	'Create File System Object
    Dim output_File As Scripting.FileSystemObject
    Set output_File = New Scripting.FileSystemObject
    
	'Create Text Stream for writing
    Dim fso As TextStream
    Set fso = output_File.OpenTextFile(text_Document_Path, ForWriting, True)
    
	'Iterator
    Dim i As Long
    i = 0
	
	'Create array or links
    Dim link_Array() As String
    link_Array = Split(current_Document.Range)
    
	'Get array size
    Dim link_Array_Size As Long
    link_Array_Size = UBound(link_Array, 1) - LBound(link_Array, 1)

	'Iterate
    Do
		'Write links by line
        fso.WriteLine (link_Array(i))
        i = i + 1
    Loop Until i > link_Array_Size
	
	'Close file
	fso.Close
	
	'Set output_File variable to nothing
    Set output_File = Nothing
End Sub


'Reading a text file
'Do While Not txtStream.AtEndOfStream
'    txtStream.ReadLine
'Loop
'txtStream.Close

