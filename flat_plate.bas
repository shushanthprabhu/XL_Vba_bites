Function POW(x As Double, y As Double)
'
' Function to return the power of a variable
' Funtion return x to the power y
' Function written by Shushanth Prabhu on 27-Feb-2016

POW = Application.WorksheetFunction.Power(x, y)
End Function


Function NCFLP(Ts As Double, Tf As Double, L As Double, d As Integer)
'
' Function to Calculate Flate Plate Heat Transfer Correlation.
' Macro written by Shushanth Prabhu on 28.01.2019
' UNITS  : SI
' Ts Temperature of Surface (K)
' Tf Inlet Temperature of Fluid (K)
' L  Length of the plate (m)
' v  Inlet velocity of plate (m/s)

'   PHYSICAL PROPERTIES
    Dim mu As Double, cp As Double, k As Double, ro As Double
    Dim Tm As Double, Ra As Double, Nu As Double, Pr As Double
    Dim n As Double
    Dim g As Double, b As Double
    
    
    Tm = (Tf + Ts) / 2
    mu = (-0.00000000003 * POW(Tm, 2)) + (0.00000006 * Tm) + 0.000003
    cp = (-0.00000000008 * POW(Tm, 3)) + (0.0000002 * POW(Tm, 2)) + (0.00002 * Tm) + 0.983
    k = (-0.00000000003 * POW(Tm, 2)) + (0.00000009 * Tm) + 0.0000008
    ro = (-0.000000001 * POW(Tm, 3)) + (0.000005 * POW(Tm, 2)) - (0.005 * Tm) + 2.587
    
    g = 9.81
    b = 0.0007
    
    Pr = mu * cp / k
    Gr = (POW(ro, 2) * g * b * cp * (Ts - Tf) * POW(L, 3)) / POW(mu, 2)
    'Gr = (POW(ro, 2) * g * b * cp * (Ts - Tf) * POW(L, 3))
     Ra = Gr * Pr
    
    If (d = 0) Then
    ' D = 0 VERTICAL PLATE
        If (Ra < 1000000000) Then
        ' FLOW IS LAMINAR
            c = 0.59
            n = 1 / 4
        ElseIf (Ra >= 1000000000) Then
        ' FLOW IS TURBULENT
            c = 0.15
            n = 1 / 3
    End If
    End If
        
    If (d = 1) Then
    ' D = 1 UPPER SURFACE OF HOT PLATE
        If (Ra < 10000000) Then
        ' FLOW IS LAMINAR
            c = 0.54
            n = 1 / 4
        ElseIf (Ra >= 10000000) Then
        ' FLOW IS TURBULENT
            c = 0.15
            n = 1 / 3
    End If
    End If
  
    If (d = -1) Then
    ' D = -1 LOWER SURFACE OF HOT PLATE
        c = 0.27
        n = 1 / 4
    End If

    If Not ((Abs(d) = 1) Or (d = 0)) Then
    ' ERROR IN D Value
        c = 0
        n = 1
    End If
    'Nu = (Gr * Pr)
    Nu = c * POW((Gr * Pr), n)
    NCFLP = Nu * k / L
End Function
