Attribute VB_Name = "Seperator"
Option Explicit
'Password "PetroleumSep"
Function SepCriticalVelocity(Dl, Dg As Double) As Double

SepCriticalVelocity = 0.048 * Sqr((Dl - Dg) / Dg)

End Function

Function SepDMIN(Qg, Vv As Double) As Double

Qg = Qg / 3600

SepDMIN = Sqr((4 * Qg) / ((22 / 7) * Vv))


End Function
