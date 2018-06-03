Sub manipulating_Objects()
    DoCmd.OpenForm "Form1", acNormal, , , acFormAdd, acHidden
    
    Dim current_ObjectName As String
    current_ObjectName = CurrentObjectName
    
    Dim f As form
    Set f = Forms(current_ObjectName)
    
    '
    ' To something
    '
    
    DoCmd.Close acForm, "Form1", acSaveNo
End Sub