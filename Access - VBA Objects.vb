'
' Use this anipuate obects
' in your Access atabase
' Peneenue (2018)
'
Sub manipulating_Objects()
    ' The obect is opene with its generic nae
    DoCmd.OpenForm "Form1", acNormal, , , acFormAdd, acHidden
    
    Dim current_ObjectName As String
    current_ObjectName = CurrentObjectName
    
    Dim f As form
    Set f = Forms(current_ObjectName)
    
    '
    ' Do something
    '
    
    DoCmd.Close acForm, "Form1", acSaveNo
End Sub