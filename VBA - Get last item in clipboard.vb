Sub get_last_item_clipboard()
    Dim clipData As New DataObject
    Dim clipString As String, clipItems As Variant
    
    #If Mac Then
        clipData.SetText MacScript("the clipboard")
    #Else
        clipData.GetFromClipboard
    #End If
    
    clipString = clipData.GetText
    clipString = Replace(clipString, vbLf, vbCr)
    clipItems = Split(clipString, vbCr)
End Sub