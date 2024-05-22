Attribute VB_Name = "AreaVolume"
Option Explicit

Function AreaCylinder(d, h As Double) As Double

AreaCylinder = (22 / 7) * d * h

End Function


Function AreaConeRoof(d As Double) As Double

Dim r, h, l As Double

r = d / 2

'Slope of 1:6 is considered

h = r / 6
l = Sqr((r ^ 2) + (h ^ 2))

AreaConeRoof = (22 / 7) * r * l

End Function


Function AreaHemiSphericalRoof(d As Double) As Double

Dim r As Double
r = d / 2
AreaHemiSphericalRoof = 2 * (22 / 7) * (r ^ 2)

End Function

Function VolumeHemiSphericalHead(d As Double) As Double

Dim r As Double
r = d / 2
VolumeHemiSphericalHead = (2 / 3) * (22 / 7) * (r ^ 3)

End Function

Function VolumeCylinder(d, h As Double) As Double

VolumeCylinder = (22 / 7) * d * d * h / 4

End Function

Function VolumeConeRoof(d As Double) As Double

Dim r, h, l As Double

r = d / 2

'Slope of 1:6 is considered

h = r / 6
l = Sqr((r ^ 2) + (h ^ 2))

VolumeConeRoof = (1 / 3) * (22 / 7) * (r ^ 2) * h

End Function


Function AreaEllipticalHead(d As Double) As Double

AreaEllipticalHead = 1.084 * d * d

End Function


Function VolumeEllipticalHead(d As Double) As Double

VolumeEllipticalHead = (22 / 7) * (d ^ 3) / 24

End Function


Function VolumeTorisphericalHead(d As Double) As Double

VolumeTorisphericalHead = 0.1694 * (d ^ 3)

End Function


Function WettedAreaHorizontalCylinder(d, h, l As Double) As Double

Dim r As Double
r = d / 2

WettedAreaHorizontalCylinder = 2 * l * r * Application.WorksheetFunction.Acos((r - h) / r)

End Function

Function PartialVolumeHorizontalCylinder(d, h, l As Double) As Double

Dim x, i, j As Double
x = h / d

i = Application.WorksheetFunction.Acos(1 - (2 * x))
j = (2 - (4 * x)) * (Sqr(x * (1 - x)))

PartialVolumeHorizontalCylinder = (l * d * d / 4) * (i - j)

End Function

Function PartialVolumeEllipticalHead(d, h As Double) As Double

Dim x As Double
x = h / d

PartialVolumeEllipticalHead = (22 / 7) * (d ^ 3) * (x * x / 24) * (3 - (2 * x))


End Function
Function AreaCrossSection(d As Double) As Double

d = d / 1000
AreaCrossSection = (22 / 7) * d * d / 4

End Function
