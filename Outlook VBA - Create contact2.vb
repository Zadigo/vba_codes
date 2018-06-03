https://www.slipstick.com/developer/code-samples/create-appointment-email-automatically/

Sub add_NewContact()
    Dim j As ContactItem
    Set j = Outlook.CreateItem(olContactItem)
    
    With j
        .Title = "Miss"
        .FirstName = "Leila"
        .MiddleName = "Goory"
        .LastName = "Lopez"
        .Gender = olFemale
        .CompanyName = "Google"
        .JobTitle = "Directrice Marketing"
        '.FileAs = "..."
        .Email1Address = "leila@gmail.com"
        .Email1AddressType = "Work"
        .WebPage = "www.google.com"
        .Anniversary = #3/10/1987#
        '.AddPicture "..."
        .Initials = "LL"
        .BusinessAddress = "Loos"
        .BusinessTelephoneNumber = "06 68 55 29 75"
        .MobileTelephoneNumber = "06 68 55 29 75"
        .MailingAddressStreet = "20 rue du Docteur Calmette"
        .MailingAddressCity = "Lille"
        .MailingAddressPostalCode = "59120"
        .Body = "Notes"
        '.Categories
        .Display
    End With
End Sub









Private Sub ListCategoryIDs()
 Dim objNameSpace As NameSpace
 Dim objCategory As Category
 Dim strOutput As String

 ' Obtain a NameSpace object reference.
 Set objNameSpace = Application.GetNamespace("MAPI")

 ' Check if the Categories collection for the Namespace
 ' contains one or more Category objects.
 If objNameSpace.Categories.Count > 0 Then

 ' Enumerate the Categories collection.
 For Each objCategory In objNameSpace.Categories

 ' Add the name and ID of the Category object to
 ' the output string.
 strOutput = strOutput & objCategory.Name & ": " & objCategory.CategoryID & vbCrLf
 Next
 End If

 ' Display the output string.
 MsgBox strOutput

 ' Clean up.
 Set objCategory = Nothing
 Set objNameSpace = Nothing

End Sub