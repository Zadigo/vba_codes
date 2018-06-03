Sub create_Table()
	'
	' Example creating a table with SQL
	'
    DoCmd.RunSQL "CREATE TABLE Kendall (" & _
                    "HerName varchar (255), " & _
                    "HerSurname varchar(255), " & _
					"HerAge int"
                    ")"
End Sub

Sub modify_Table()
    On Error Resume Next
    DoCmd.RunSQL "ALTER TABLE Kendall " & _
                 "ADD COLUMN Address varchar(255)"
End Sub

Sub modify_Table()
    DoCmd.RunSQL "UPDATE Google SET Test = '1' WHERE Nom = 'Julie'"
End Sub
