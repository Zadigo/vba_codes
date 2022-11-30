Sub Cleaning()
    Dim wk As Worksheet
    Set wk = Worksheets("Clean")
    Dim r As Range
    Set r = wk.Range("E2:" & wk.Range("E2").End(xlDown).Address)
    Dim companyAddress As String
    Dim b As Integer
    For i = 0 To r.Count
      ' Debug.Print Application.WorksheetFunction("=LEFT(" & r(i, 1).Address & ";LEN(" & r(i, 1).Address & ") - 12)")
    Next i
End Sub
