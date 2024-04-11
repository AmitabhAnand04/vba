Option Explicit

Sub Test()
    Dim i As Integer
    Dim j As Integer
    For j = 1 To 2
        For i = 1 To 100
            Sheet1.Cells(i, j).Value = i
        Next i
    Next j
End Sub