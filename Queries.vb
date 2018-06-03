Sub dynamic_Query_SQL()
	'
	' This sub dynamically programatically modifies a query
	'
	
    Dim db As DAO.Database
    Set db = CurrentDb
    
	'Sets query
    Dim qdf As QueryDef
    Set qdf = db.QueryDefs("Google")
    
	'SQL string filter query
    Dim string_SQL As String
    string_SQL = "SELECT * " & _
                 "FROM Table1 " & _
                 "WHERE [Cible1] = 'Paul'"
    
	'Run
    qdf.SQL = string_SQL
    
	'Open
    'DoCmd.OpenQuery "Google"
    
    Set qdf = Nothing
    Set db = Nothing
End Sub

Sub creating_Query()
	'
	' This sub creates a query
	'
	
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim qdf As DAO.QueryDef
    Dim newSQL As String
    
    newSQL = "Select * From [Google] WHERE [Cible2]>'2010'"
    Set qdf = db.CreateQueryDef("tempQry", newSQL)
End Sub