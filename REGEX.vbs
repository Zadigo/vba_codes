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

[a-z] = a, b c... or z
[^c] = everything not c
[a-z].* = a, b, c... or z + anything behind
.*[a-z] = Anything in front + a, b, c... or z
[a-zA-Z] = a, b, c... or z / A, B, C... or Z
[aA][a-zA-Z] = a or A followed by a, b, c... z / A, B, C... Z
[a-b][0-2] = Same with number 0, 1 or 2
[a-b]\\d = Same with anykind of number
[a-c[t-w]] = a, b or c / t, u, v or w
[a-c&&[^b]] = a or c / NOT b
[a-e&&[^c-e]] = a or b / NOT c, d or e

[a-c]\\d{2} = a, b or c followed by 2 numbers
[a-z]{3} = a, b AND c
[a-c].{2} = a, b or c followed by anything of length 2

[a-c]\\s[a-c] = a, b or c followed by whitespace followed by a, b or c
[a-c]\\S[a-c] = a, b or c NOT followed by whitespace followed by a, b or c

[a-c]\\d = a, b or c followed by a number

\\d+(\\.\\d+)? = number + dot + [group of (made optional due to question mark] numbers

(...) = Element as a group
(a|b) = a OR b

+ = one or many times e.g. \\s+ one or many times whitespace