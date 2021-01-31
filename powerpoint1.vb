Dim ppt As Presentation
Dim slides As slides
Dim slide As slide
Dim title, shape As shape
Dim shape_titles As String

Public Sub titles()
    Set ppt = ActivePresentation
    
    For i = 1 To ppt.slides.Count
        Set slide = ppt.slides.Item(i)
        slide.Select
        
        If slide.shapes.Count > 1 Then
            Set shape = slide.shapes.Item(1)
            slide.Select
            
            If shape.Name = "Title" Then
                With shape.TextFrame2.TextRange.Font
                    .Name = "Roboto"
                    .Size = 44
                End With
            End If
'            shape.Height = 104
        End If
    Next
End Sub

Public Sub text_boxes()
    Set ppt = ActivePresentation
    
    For i = 1 To ppt.slides.Count
        Set slide = ppt.slides.Item(i)
        slide.Select
        If slide.shapes.Count > 1 Then
            Set shape = slide.shapes.Item(2)
            If shape.Name = "TextBox" Then
                shape.ScaleWidth 1, msoFalse, msoScaleFromMiddle
            End If
            slide.shapes.Range.Align msoAlignCenters, msoFalse
        End If
    Next
End Sub

