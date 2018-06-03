https://msdn.microsoft.com/en-us/library/office/ff821396.aspx

Sub finding_Record()
    '
    ' Opening a record set and findind a record
    '
    Dim db As DAO.Database
    Dim rs As Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Tournaments", dbOpenSnapshot)
    
    rs.FindFirst "[TourCode] LIKE 'TOR'"
    MsgBox rs(2).Value, vbInformation, "Value"
    
    Set db = Nothing
    Set rs = Nothing
End Sub

Sub filter_RecordSet()
    '
    ' Opening a record set and finding a record
    '
    Dim db As DAO.Database
    Dim rs As Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * " & _
							  "FROM Tournaments " & _
							  "WHERE TourCode = 'TKY'")
    'Set rs = db.OpenRecordset("SELECT * FROM Tournaments " & _
	'						   "WHERE TourCode = 'TKY' AND/OR/NOT ... ''")
	'Set rs = db.OpenRecordset("SELECT * " & _
	'						   "FROM Tournaments " & _
	'						   "WHERE TourCode = 'TKY' ORDER BY ... DESC/ASC ")
	
    'TO DO
    
    Set db = Nothing
    Set rs = Nothing
End Sub

Sub printing_Elements()
	'
	'	Prints everything from a recordset
	'
    Dim db As DAO.Database
    Set db = CurrentDb

    Dim rs As Recordset
    Set rs = db.OpenRecordset("Google")
    
    Do While Not rs.EOF
       Debug.Print rs("ID") & " - " & rs("Cible1")
       rs.MoveNext
    Loop
End Sub