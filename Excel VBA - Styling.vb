Sub styling_Macro(ByVal font_name As String, ByVal font_size As Long, Optional ByVal new_sheet_name As String)
    Dim o As Worksheet
    Set o = Worksheets(ActiveSheet.Name)
    o.Range(o.Range("A1").End(xlToRight), o.Range("A1").End(xlDown)).Select   
    With o.Cells
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        With .Characters.Font
            .Name = font_name
            .Size = font_size
        End With
    End With
    o.Range("A1").Select
    If Not IsEmpty(new_sheet_name) Or Not new_sheet_name = "" Then
        o.Name = new_sheet_name
    End If
End Sub
