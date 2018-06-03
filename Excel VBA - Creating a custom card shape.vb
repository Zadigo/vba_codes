Dim wk As Worksheet

Public Sub redim_shape()
    MsgBox a
End Sub

Public Sub create_new_card()
    'Application.ScreenUpdating = False
    
    Dim card_shape As Shape
    Set card_shape = wk.Shapes.AddShape(msoShapeRectangle, 10, 90, 200, 30)
    
    With card_shape
    '    With .Glow
    '        .Color = RGB(15, 12, 13)
    '        .Radius = 5
    '        .Transparency = 0.8
    '    End With
    '    .ScaleWidth 2, msoFalse
        .Name = "cards_card"
        .ZOrder msoSendToBack
        .Line.Visible = msoFalse
        With .Fill
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
        With .Shadow
            .Visible = msoTrue
            .Blur = 2
            .Style = msoShadowStyleOuterShadow
            .ForeColor.RGB = RGB(208, 206, 206)
            '.ForeColor.RGB = &;BDC3C7
            .Size = 102
            .Transparency = 0.6
            .OffsetX = 0
            .OffsetY = 0.35
        End With
        With .TextFrame
            .VerticalAlignment = xlVAlignCenter
            .HorizontalAlignment = xlHAlignCenter
        End With
        With .TextFrame2
            With .TextRange
                .Text = "Votre texte ici"
                With .Font
                    .Name = "Lato"
                    .Fill.ForeColor.RGB = RGB(0, 0, 0)
                End With
                With .ParagraphFormat
                    .Alignment = msoAlignCenter
                End With
                .Select
            End With
        End With
    End With
    
    'Application.ScreenUpdating = True
End Sub

Private Sub Class_Initialize()
    Set wk = Worksheets(ActiveSheet.Index)
End Sub
