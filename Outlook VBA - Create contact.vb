Public Sub CreateNewContact()
	Dim objContact As ContactItem
	Set objContact = Application.CreateItem(olContactItem)
	
	With objContact 
		.BusinessAddressCity = "Halifax"
		.BusinessAddressCountry = "Canada"
		.Business2TelephoneNumber = "902123" 'the area code and local prefix
		.Display
	End With

	Set objContact = Nothing
End Sub