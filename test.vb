Option Explicit

Sub Test()
    Dim i As Integer
    For i = 1 To 100
        Sheet1.Cells(i, 1).Value = i
    Next i
End Sub