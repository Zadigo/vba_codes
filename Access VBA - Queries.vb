'
' Use this sub to create a query
' John PM (2017)
'
'
Sub create_Query()
    Dim qdf As QueryDef
    Set qdf = CurrentDb.CreateQueryDef("query1", "SELECT * FROM Table")
    
    On Error Resume Next
    
    DoCmd.OpenQuery "query1", acViewDesign, acEdit
    DoCmd.Save acQuery, "query1"
    DoCmd.Close acQuery, "query1"
    DoCmd.rename "new_name", acQuery, "query1"
    
    Set qdf = Nothing
End Sub

'
' Use this sub to change the source of a query
' John PM (2017)
'
'
Sub manipulate_Query()
    Dim query_to_change As QueryDef
    Set query_to_change = CurrentDb.QueryDefs("query_name")
    
    query_to_change.SQL = "SELECT * FROM Table ORDER BY ID Asc"
    query_to_change.SQL = "SELECT Field1, Field2 FROM Table ORDER BY ID Asc"
    query_to_change.SQL = "SELECT Field1, Field2 FROM Table WHERE Field LIKE Fashion"
    query_to_change.SQL = "SELECT Field1, Field2 FROM Table WHERE Field LIKE '" & something & "'"
End Sub

' "SELECT Field1, Field2 FROM Table WHERE Field1 = 'Fashion'"