Function findval(val As Double, range1 As Range, range2 As Range)
    ' Finds "val" from "Range1" and Returns Value from "Range2"
    ' Primarily used for getting data at a particaular time from CSV dump from a Engineering Softwared    
    Dim Loc As Integer
    Loc = Application.Match(val, range1, 0)
    a = range2(Loc, 1)
    findval = a
End Function
