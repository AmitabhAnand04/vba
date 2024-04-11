Option Explicit

Sub TestModule2()
    Dim i As Integer
    For i = 1 To 100
        Sheet1.Cells(i, j).Value = i
    Next i
    
    Dim j As Integer
    For j = 1 To 2
        For i = 1 To 100
            Sheet1.Cells(i, j).Value = i
        Next i
    Next j
End Sub