Function Interpolate(val As Double, x_range As Range, y_range As Range)
'Function written by Shushanth M S Prabhu
' Date - 04 December 2010
                
    If x_range.Cells(1, 1).Value < x_range.Cells(2, 1).Value Then
        Order = 1
    Else
        Order = -1
        
    End If
    r = Application.Match(val, x_range, Order)
    x1 = x_range.Cells(r, 1)
    x2 = x_range.Cells(r + 1, 1)
    
    y1 = y_range(r, 1)
    y2 = y_range(r + 1, 1)
    
    Interpolate = ((y2 - y1) / (x2 - x1) * (val - x1)) + y1
    
End Function


