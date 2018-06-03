Sub test_if_OutlookIsOpen()
    Dim oOutlook As Object

    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If oOutlook Is Nothing Then
        MsgBox "Outlook is not open, open Outlook and try again"
    Else
        
		' TO DO
		'
        MsgBox "Is Open"
    End If
End Sub