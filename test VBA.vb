Sub g()
    Dim o As Range
    Set o = ActiveDocument.Range(0, 300)
    
    t = 0
    
    For Each w In o.Words
        w = Trim(w)
        If w = "you" Then
            t = t + 1
            If t > o.Words.Count Then
                Exit Sub
            End If
            o.Words.Item(t).Bold = True
        End If
        t = t + 1
    Next
End Sub

Sub g()
    Dim o As Range
    Set o = ActiveDocument.Range(0, Selection.End)
    
    o.Bold = True
End Sub

Sub h()
    Dim o As Document
    Set o = ActiveDocument
    
    Dim w As Paragraphs
    Set w = ActiveDocument.Paragraphs
    
    w.Item(2).Alignment = wdAlignParagraphCenter
End Sub
