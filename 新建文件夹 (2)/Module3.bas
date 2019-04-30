Attribute VB_Name = "Module3"

Global g1x As Double, g1y As Double, g2x As Double, g2y As Double, X As Double, Y As Double, Z As Double
Sub g1(ax As Double, ay As Double, bx As Double, by As Double)
Dim vx As Double, vy As Double, t1 As Integer, t2 As Integer, a As Double, b As Double, r As Double
g1x = 0
g1y = 0
For t1 = -10 To 10 Step 1
    For t2 = 1 To 10 Step 1
        vx = ax * t1 + bx * t2
        vy = ay * t1 + by * t2
        a = (vx * vx - 6 * vy * vy) * vx * vx + vy * vy * vy * vy
        b = 4 * vx * vy * (vy + vx) * (vy - vx)
        r = a * a + b * b
        g1x = g1x + 120 * a / r
        g1y = g1y + 120 * b / r
    Next t2
Next t1
For t1 = 1 To 10 Step 1
    
        vx = ax * t1
        vy = ay * t1
        a = (vx * vx - 6 * vy * vy) * vx * vx + vy * vy * vy * vy
        b = 4 * vx * vy * (vy + vx) * (vy - vx)
        r = a * a + b * b
        g1x = g1x + 120 * a / r
        g1y = g1y + 120 * b / r
    
Next t1

End Sub
Sub g2(ax As Double, ay As Double, bx As Double, by As Double)
Dim vx As Double, vy As Double, t1 As Integer, t2 As Integer, a As Double, b As Double, r As Double
g2x = 0
g2y = 0
For t1 = -10 To 10 Step 1
    For t2 = 1 To 10 Step 1
        vx = ax * t1 + bx * t2
        vy = ay * t1 + by * t2
        a = ((vx * vx - 15 * vy * vy) * vx * vx + 15 * vy * vy * vy * vy) * vx * vx - vy * vy * vy * vy * vy * vy
        b = -2 * vx * vy * ((3 * vx * vx - 10 * vy * vy) * vx * vx + 3 * vy * vy * vy * vy)
        r = a * a + b * b
        g2x = g2x + 280 * a / r
        g2y = g2y + 280 * b / r
    Next t2
Next t1

For t1 = 1 To 10 Step 1
  
        vx = ax * t1
        vy = ay * t1
        a = ((vx * vx - 15 * vy * vy) * vx * vx + 15 * vy * vy * vy * vy) * vx * vx - vy * vy * vy * vy * vy * vy
        b = -2 * vx * vy * ((3 * vx * vx - 10 * vy * vy) * vx * vx + 3 * vy * vy * vy * vy)
        r = a * a + b * b
        g2x = g2x + 280 * a / r
        g2y = g2y + 280 * b / r
    
Next t1

End Sub
Sub reduce()
Dim r As Double
r = Sqr(g1x * g1x + g1y * g1y + g2x * g2x + g2y * g2y)
g1x = g1x / r
g1y = g1y / r
g2x = g2x / r
g2y = g2y / r
End Sub
Sub rotatey(a1 As Double)
Dim x1 As Double, y1 As Double, z1 As Double, t1 As Double
x1 = g1x: y1 = g1y: z1 = g2x: t1 = g2y

g1y = y1 * Cos(a1) - t1 * Sin(a1)
g2y = y1 * Sin(a1) + t1 * Cos(a1)
End Sub
Sub rotatez(a1 As Double)
Dim x1 As Double, y1 As Double, z1 As Double, t1 As Double
x1 = g1x: y1 = g1y: z1 = g2x: t1 = g2y

g2x = z1 * Cos(a1) - t1 * Sin(a1)
g2y = z1 * Sin(a1) + t1 * Cos(a1)
End Sub
Sub rotatet(a1 As Double)
Dim x1 As Double, y1 As Double, z1 As Double, t1 As Double
x1 = g1x: y1 = g1y: z1 = g2x: t1 = g2y

g1x = x1 * Cos(a1) - t1 * Sin(a1)
g2y = x1 * Sin(a1) + t1 * Cos(a1)
End Sub
Sub projection(a As Boolean)
If a Then
X = 2 * g1x / (1 - g2y)
Y = 2 * g1y / (1 - g2y)
Z = 2 * g2x / (1 - g2y)
Else
X = 2 * g1x
Y = 2 * g1y
Z = 2 * g2x
End If
End Sub
Sub getpoint(ax As Double, ay As Double, bx As Double, by As Double, a1 As Double, a2 As Double, a3 As Double)
Call g1(ax, ay, bx, by)
Call g2(ax, ay, bx, by)
Call reduce
Call rotatey(a1)
Call rotatez(a2)
Call rotatet(a3)
Call projection(True)
End Sub
Sub getpoint2(ax As Double, ay As Double, bx As Double, by As Double, a1 As Double, a2 As Double, a3 As Double)
g1x = ax
g1y = ay
g2x = bx
g2y = by
Call reduce
Call rotatey(a1)
Call rotatez(a2)
Call rotatet(a3)
Call projection(True)








End Sub

