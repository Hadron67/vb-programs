Attribute VB_Name = "Module2"
Global ball(100) As Planet, ball2(100) As Planet, balln As Integer, balln2 As Integer, ax(100) As Double, ay(100) As Double, az(100) As Double, t As Double, zero As Planet

Global start As Boolean

Sub react(k As Double, dt As Double, kk As Double)
On Error Resume Next
Dim m As Integer, n As Integer
For m = 1 To balln Step 1
    ax(m) = 0
    ay(m) = 0
    az(m) = 0
    For n = 1 To balln Step 1
        ax(m) = ax(m) + forcex(ball(m), ball(n), k, kk)
        ay(m) = ay(m) + forcey(ball(m), ball(n), k, kk)
        az(m) = az(m) + forcez(ball(m), ball(n), k, kk)
    Next n
Next m

For m = 1 To balln Step 1
    ball(m).vx = ball(m).vx + ax(m) * dt
    ball(m).vy = ball(m).vy + ay(m) * dt
    ball(m).vz = ball(m).vz + az(m) * dt
    ball(m).x = ball(m).x + ball(m).vx * dt + 0.5 * ax(m) * dt * dt
    ball(m).y = ball(m).y + ball(m).vy * dt + 0.5 * ay(m) * dt * dt
    ball(m).z = ball(m).z + ball(m).vz * dt + 0.5 * az(m) * dt * dt
Next m
t = t + dt
End Sub
Function l(fx1, fy1, fz1, fx2, fy2, fz2) As Double
    l = ((fx1 - fx2) * (fx1 - fx2) + (fy1 - fy2) * (fy1 - fy2) + (fz1 - fz2) * (fz1 - fz2)) ^ (3 / 2)
End Function
Sub atstart()
start = False
Dim m As Integer
balln = 1
For m = 1 To 100 Step 1
    ball(m).radius = 5
    ball(m).m = 10
Next m
selectball = 1
zero.m = 10
zero.radius = 5
End Sub
Function Lenth(fx1, fy1, fz1, fx2, fy2, fz2) As Double
    Lenth = ((fx1 - fx2) * (fx1 - fx2) + (fy1 - fy2) * (fy1 - fy2) + (fz1 - fz2) * (fz1 - fz2)) ^ (1 / 2)
End Function

Function Sign(x As Double) As Double
If x > 0 Then
Sign = 1
ElseIf x < 0 Then
Sign = -1
ElseIf x = 0 Then
Sign = 0
End If
End Function
