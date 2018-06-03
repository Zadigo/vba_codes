Sub add_recepient()
    Dim new_Message As MailItem
    Set new_Message = Application.CreateItem(olMailItem)
    
    With new_Message
        .To = "x@gmail.com"
        .CC = "x@gmail.com"
        .BCC = "x@gmail.com"
        .Subject = "Google"
        '.Categories = "Test"
        '.VotingOptions = "Yes;No;Maybe"
        .BodyFormat = olFormatHTML
        .Body = "Text"
        '.Importance = olImportanceHigh
        '.Sensitivity = olConfidential
        '.Attachments.Add "..."
        '.ExpiryTime = DateAdd("m", 6, Now)
        '.DeferredDeliveryTime = #8/5/2018 6:00:00 PM#
        .Display
    End With
    
    Set new_Message = Nothing
End Sub



Public Sub CreateNewMessage()
	'
	' Sends mail to sender based on the active selection
	'

Dim objMsg As MailItem
Dim Selection As Selection
Dim obj As Object

Set Selection = ActiveExplorer.Selection

For Each obj In Selection

Set objMsg = Application.CreateItem(olMailItem)

 With objMsg
  .To = obj.SenderEmailAddress
  .Subject = "This is the subject"
  .Categories = "Test"
  .Body = "My notes" & vbCrLf & vbCrLf & obj.Body
  .Display
' use .Send to send it automatically

End With
Set objMsg = Nothing

Next

End Sub