Option Explicit

Dim this_Sheet As Worksheet

'
' This adds a new rectangular shape to an Excel sheet
'

Sub add_NewBox()
    'Set sheet
    Set this_Sheet = Worksheets(ActiveSheet.Index)
    
    'Add new shape
    this_Sheet.Shapes.AddShape(msoShapeRectangle, 150, 150, 320, 30).Visible = msoTrue
    
    'Get the last shape
    Dim last_Shape As Long
    last_Shape = count_Shapes
    
    'Style shape
    With this_Sheet.Shapes.Item(last_Shape)
        .BackgroundStyle = msoBackgroundStylePreset2
        .Line.Visible = msoFalse
        
        With .TextFrame2
            .TextRange.Text = "Google"
            
            With .TextRange.Font
                .Fill.ForeColor.RGB = RGB(1, 1, 1)
                .Size = 12
                .Name = "Open Sans"
            End With
            
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

Private Function count_Shapes() As Long
    '
	' This function counts the amount of shapes
    '
	count_Shapes = this_Sheet.Shapes.Count
End Function

'
' Same with shape shadow
'

Option Explicit

Dim this_Sheet As Worksheet

Sub add_NewBox()
    'Set sheet
    Set this_Sheet = Worksheets(ActiveSheet.Index)
    'Add new shape
    this_Sheet.Shapes.AddShape(msoShapeRectangle, 150, 150, 320, 30).Visible = msoTrue
    
    'Get the last shape
    Dim last_Shape As Long
    last_Shape = count_Shapes
    
    'Style shape
    With this_Sheet.Shapes.Item(last_Shape)
        .BackgroundStyle = msoBackgroundStylePreset2
        .Line.Visible = msoFalse
        
        With .Shadow
            .Blur = 5
            .Transparency = 0.7
            .OffsetX = -1
            .OffsetY = 1
            .Visible = msoTrue
        End With
        
        With .TextFrame2
            With .TextRange
                .ParagraphFormat.Alignment = msoAlignCenter
                
                With .Font
                    .Fill.ForeColor.RGB = RGB(1, 1, 1)
                    .Size = 12
                    .Name = "Open Sans"
                End With
                
                .Text = "-- Tapez votre texte ici --"
                .Select
            End With
        End With
    End With
End Sub

Private Function count_Shapes() As Long
    '
	' This function counts the amount of shapes
    '
    count_Shapes = this_Sheet.Shapes.Count
End Function

