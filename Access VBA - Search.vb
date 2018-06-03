Public Sub search_DB_Static()
    ' Element to search
	q = Text3.Value	
    ' Search record number
	search = DLookup("[CustID]", "Customer", "[CustName]='" & q & "'")
    ' Goto record
    DoCmd.GoToRecord , "Form1", acGoTo, search
    ' Box blank
    Text3.Value = ""
    ' Set focus
    CustName.SetFocus
End Sub

' WILDCARDS
' search = DLookup("[CustID]", "Customer", "[CustName]='*" & q & "*'")