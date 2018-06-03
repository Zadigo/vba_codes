Sub manipulating_Tables()
    On Error Resume Next
	'INSERT values
    DoCmd.RunSQL "INSERT INTO Facebook(OK, Field1) VALUES ('5', 'Kendall')"
	'UPDATE field
	DoCmd.RunSQL "UPDATE Facebook SET Field1 = 'Kendall' WHERE ID = 1"
	'ALTER TABLE
	DoCmd.runSQL "ALTER TABLE X "
End Sub

Sub edit_Table()
    Dim d As DAO.Database
    Dim t As TableDef
    Dim r As DAO.Recordset
    
    Set d = CurrentDb
    Set t = d.TableDefs("...")
    Set r = t.OpenRecordset(, dbOpenSnapshot)
    
    r.Edit
    r(...).Value = "..."
    r.Update
    
    Set d = Nothing
    Set t = Nothing
End Sub

'When field exists, create random number and put to name "Address"
    If Err.Number = 3380 Then
        DoCmd.RunSQL "ALTER TABLE Kendall " & _
                     "ADD COLUMN Address" & Int((25 - 10 + 1) * Rnd + 10) & " varchar(255)" & _
                     ")"
        
    End If

Tables
1. "CREATE TABLE x (...[datatype])"
2. "INSERT INTO x (...) VALUES (... [datatype])"
3. "ALTER TABLE x ADD x [datatype]"
4. "ALTER TABLE x DROP x"
5. "UPDATE x SET x = '...'"
6. "UPDATE x SET x = '...' WHERE x = '...' AND x = '...'"
7. "UPDATE x SET x = '...' WHERE x = '...' OR x = '...'"
8. "UPDATE TABLE x ALTER COLUMN x [datatype]"
9. "CREATE TABLE x (...[datatype] NOT NULL UNIQUE)"
10. "CREATE TABLE x (...[datatype] NOT NULL PRIMARY KEY)"
11. ??? Foreign Key

Queries
1. Create query - "SELECT x FROM x / SELECT x, y FROM x, y"
2. "SELECT x, y.something FROM x INNER JOIN y ON x.something = y.something"
3. "SELECT x FROM x ORDER BY x"
4. Recordset openRecordSet(sql) / openRecordSet(...)
5. .FINDFIRST, FINDLAST etc. "[fieldname] > x"
6. Record Count
7. Fields

DLookup
DLookup("[field]", "table", "[something]=x")

Filter Forms
1. .filter "[field]=x" / .filterOn = true/false

"SELECT CustID, CustName, Product.ProductName " & _
                    "FROM Customer " & _
                    "INNER JOIN Product ON Customer.CustID = Product.CustomerID"