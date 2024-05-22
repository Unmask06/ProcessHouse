'Horizontal KOD Sizing based on API 521 Section C.3
Option Explicit
Dim g As Double

Function DropoutVelocity(D As Double, rho_l As Double, rho_v As Double, C As Double) As Variant
    g = 9.81
    DropoutVelocity = 1.15 * Sqr((g * D * (rho_l - rho_v)) / (rho_v * C))
End Function

Function CRe_sqr(rho_v As Double, D As Double, rho_l As Double, vis_g As Double) As Variant
    CRe_sqr = (0.13 * 10 ^ 8 * rho_v * D ^ 3 * (rho_l - rho_v)) / (vis_g ^ 2)
End Function

Function DragCoefficient(CRe_sqr As Double) As Variant
    If CRe_sqr = 0 Then
        DragCoefficient = 0
    Else
        If CRe_sqr >= 10 Then
            If CRe_sqr < 180 Then
                DragCoefficient = 336.62 * CRe_sqr ^ -0.7638
            ElseIf CRe_sqr <= 1000 Then
                DragCoefficient = 120.87 * CRe_sqr ^ -0.5668
            ElseIf CRe_sqr <= 10500 Then
                DragCoefficient = 50.746 * CRe_sqr ^ -0.42747
            ElseIf CRe_sqr <= 200000 Then
                DragCoefficient = 7.9273 * CRe_sqr ^ -0.2268
            Else
                DragCoefficient = 2.4562 * CRe_sqr ^ -0.1277
            End If
        Else
            DragCoefficient = 60
        End If
    End If
End Function