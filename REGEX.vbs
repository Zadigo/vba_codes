http://www.regular-expressions.info/quickstart.html
'
' Using REGEX to test a pattern
' Activate 'Microsoft VBScript Reguar Expressions 5.5'
'
Private Sub pattern_Test()
    With regEx
        .Global = True
        .IgnoreCase = False
        .pattern = "([a-zA-Z]*)\s+([a-zA-Z]*)"
        .Multiline = False
    End With
    
    If Not regEx.Test(string_Input) Then
        MsgBox "Error", vbCritical, "Error"
        Exit Function
    End If
End Function
