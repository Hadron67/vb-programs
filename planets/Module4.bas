Attribute VB_Name = "Module4"
Type Planet
    x As Double
    y As Double
    z As Double
    vx As Double
    vy As Double
    vz As Double
    m As Double
    radius As Double
End Type
Global selectball As Integer
Function hitTest(p1 As Planet, p2 As Planet) As Boolean
Dim l As Double
l = Lenth(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)
If l < p1.radius + p2.radius Then hitTest = True Else hitTest = False
End Function
Function forcex(p1 As Planet, p2 As Planet, k As Double, kk As Double)
If hitTest(p1, p2) And Form1.Check2.Value Then
    forcex = k * p2.m * (p2.x - p1.x) / l(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z) - kk * (p2.x - p1.x) * (p1.radius + p2.radius - Lenth(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)) / Lenth(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z) / p1.m
Else
    forcex = k * p2.m * (p2.x - p1.x) / l(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)
End If
End Function
Function forcey(p1 As Planet, p2 As Planet, k As Double, kk As Double)
If hitTest(p1, p2) And Form1.Check2.Value Then
    forcey = k * p2.m * (p2.y - p1.y) / l(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z) - kk * (p2.y - p1.y) * (p1.radius + p2.radius - Lenth(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)) / Lenth(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z) / p1.m
Else
    forcey = k * p2.m * (p2.y - p1.y) / l(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)
End If
End Function

Function forcez(p1 As Planet, p2 As Planet, k As Double, kk As Double)
If hitTest(p1, p2) And Form1.Check2.Value Then
    forcez = k * p2.m * (p2.z - p1.z) / l(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z) - kk * (p2.z - p1.z) * (p1.radius + p2.radius - Lenth(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)) / Lenth(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z) / p1.m
Else
    forcez = k * p2.m * (p2.z - p1.z) / l(p1.x, p1.y, p1.z, p2.x, p2.y, p2.z)
End If
End Function

