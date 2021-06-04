Function interpolate(x0 As Variant, y0 As Variant, _
    x1 As Variant, y1 As Variant, x As Variant) As Variant
    interpolate = y0 + (y1 - y0) / (x1 - x0) * (x - x0)
End Function

Function interpolateArray(arrayX As Variant, arrayY As Variant, x As Double) As Variant
        
        Dim n As Long
        n = arrayX.Cells.Count
        If n < 2 Then
            interpolateArray = "The arrays must have at least 2 values"
            Exit Function
        End If
        If n <> arrayY.Cells.Count Then
            interpolateArray = "The arrays must have equal sizes"
            Exit Function
        End If
        
        Dim x0, y0, x1, y1
        x0 = arrayX.Cells(1).Value
        y0 = arrayY.Cells(1).Value
        x1 = arrayX.Cells(2).Value
        y1 = arrayY.Cells(2).Value
        
        If x > arrayX(n) Then
            
            x0 = arrayX.Cells(n - 1).Value
            y0 = arrayY.Cells(n - 1).Value
            x1 = arrayX.Cells(n).Value
            y1 = arrayY.Cells(n).Value
                    
        ElseIf x > x1 Then
            
            Dim i As Long
            For i = 3 To arrayX.Cells.Count
                If x < arrayX.Cells(i).Value Then
                    x0 = arrayX.Cells(i - 1).Value
                    y0 = arrayY.Cells(i - 1).Value
                    x1 = arrayX.Cells(i).Value
                    y1 = arrayY.Cells(i).Value
                    Exit For
                End If
            Next i
            
        End If
        
        interpolateArray = interpolate(x0, y0, x1, y1, x)
End Function

Sub testInterpolateArray()
    Dim arrayX As Range
    Dim arrayY As Range
    Set arrayX = Range("R6:R8") 'Worksheets("EAP-Bore").Range("R6:R8")
    Set arrayY = Range("S6:S8") 'Worksheets("EAP-Bore").Range("S6:S8")
    MsgBox interpolateArray(arrayX, arrayY, 20)
End Sub
