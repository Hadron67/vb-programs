Attribute VB_Name = "Module2"
Global balls(1000, 3), balls2(1000, 3) As Double, enable, enable2 As Boolean



Sub placeballs(pointx As Double, pointy As Double, pointz As Double)
On Error Resume Next
Dim tx, ty, tz As Double, t As Integer
t = 1
For tx = pointx To pointx + 60 Step 6
For ty = pointy To pointy + 60 Step 6
For tz = pointz To pointz + 60 Step 6
    balls(t, 1) = tx
    balls(t, 2) = ty
    balls(t, 3) = tz
    t = t + 1
Next tz
Next ty
Next tx
End Sub
Sub copy()
Dim x As Integer
x = 1
For x = 1 To 1000 Step 1
    balls2(x, 1) = balls(x, 1)
    balls2(x, 2) = balls(x, 2)
    balls2(x, 3) = balls(x, 3)
Next x

End Sub
Sub displayballs(a As Double, b As Double, c As Double, dt As Double)
On Error Resume Next
Dim x As Integer, vx, vy, vz As Double
For x = 1 To 1000 Step 1
    glTranslatef balls2(x, 1), balls2(x, 2), balls2(x, 3)
    glutSolidSphere 1, 10, 10
    glTranslatef -balls2(x, 1), -balls2(x, 2), -balls2(x, 3)
    If enable Then
        vx = a * (balls2(x, 2) - balls2(x, 1))
        vy = balls2(x, 1) * (c - balls2(x, 3)) - balls2(x, 2)
        vz = balls2(x, 2) * balls2(x, 1) - b * balls2(x, 3)
        balls2(x, 1) = balls2(x, 1) + vx * dt
        balls2(x, 2) = balls2(x, 2) + vy * dt
        balls2(x, 3) = balls2(x, 3) + vz * dt
    End If
Next x
End Sub
Sub displayattractor(x As Double, y As Double, z As Double, a As Double, b As Double, c As Double)
On Error Resume Next
Dim t As Double
If enable2 Then
For t = 1 To 10 Step 0.001
    vx = -a * (x - y)
    vy = x * (c - z) - y
    vz = x * y - b * z
    x = x + vx / 100
    y = y + vy / 100
    z = z + vz / 100
    Call Lineto(x, y, z)
Next t
End If
End Sub
Sub displayballs2(a As Double, dt As Double)
On Error Resume Next
Dim x As Integer, vx, vy, vz As Double
For x = 1 To 1000 Step 1
    glTranslatef balls2(x, 1), balls2(x, 2), balls2(x, 3)
    glutSolidSphere 1, 10, 10
    glTranslatef -balls2(x, 1), -balls2(x, 2), -balls2(x, 3)
    If enable Then
        vx = -balls2(x, 2) / 10 - balls2(x, 3) / 10
        vy = balls2(x, 1) / 10 + balls2(x, 2) / 10 * a
        vz = 2 + balls2(x, 3) / 10 * (balls2(x, 1) / 10 - 4)
        balls2(x, 1) = balls2(x, 1) + vx * dt
        balls2(x, 2) = balls2(x, 2) + vy * dt
        balls2(x, 3) = balls2(x, 3) + vz * dt
    End If
Next x
End Sub
Sub displayattractor2(x As Double, y As Double, z As Double, a As Double)
On Error Resume Next
Dim t As Double
If enable2 Then
For t = 1 To 10 Step 0.001
    vx = -y - z
    vy = x + a * y
    vz = 2 + z * (x - 4)
    x = x + vx / 100
    y = y + vy / 100
    z = z + vz / 100
    Call Lineto(x * 10, y * 10, z * 10)
Next t
End If
End Sub

