Attribute VB_Name = "Module2"
Global x1, y1, z1, x2, y2, z2, x3, y3, z3 As Double, vx1, vy1, vz1, vx2, vy2, vz2, vx3, vy3, vz3 As Double, m1, m2, m3 As Double, start As Boolean, t As Double
Global x(3), y(3), z(3), vx(3), vy(3), vz(3) As Double
Sub react(k As Double, dt As Double, kk As Double)
Dim ax1, ay1, az1, ax2, ay2, az2, ax3, ay3, az3 As Double
    ax1 = k * m2 * (x2 - x1) / L(x1, y1, z1, x2, y2, z2) + k * m3 * (x3 - x1) / L(x1, y1, z1, x3, y3, z3)
    ay1 = k * m2 * (y2 - y1) / L(x1, y1, z1, x2, y2, z2) + k * m3 * (y3 - y1) / L(x1, y1, z1, x3, y3, z3)
    az1 = k * m2 * (z2 - z1) / L(x1, y1, z1, x2, y2, z2) + k * m3 * (z3 - z1) / L(x1, y1, z1, x3, y3, z3)
    
    ax2 = k * m1 * (x1 - x2) / L(x2, y2, z2, x1, y1, z1) + k * m3 * (x3 - x2) / L(x2, y2, z2, x3, y3, z3)
    ay2 = k * m1 * (y1 - y2) / L(x2, y2, z2, x1, y1, z1) + k * m3 * (y3 - y2) / L(x2, y2, z2, x3, y3, z3)
    az2 = k * m1 * (z1 - z2) / L(x2, y2, z2, x1, y1, z1) + k * m3 * (z3 - z2) / L(x2, y2, z2, x3, y3, z3)
    
    ax3 = k * m1 * (x1 - x3) / L(x3, y3, z3, x1, y1, z1) + k * m2 * (x2 - x3) / L(x3, y3, z3, x2, y2, z2)
    ay3 = k * m1 * (y1 - y3) / L(x3, y3, z3, x1, y1, z1) + k * m2 * (y2 - y3) / L(x3, y3, z3, x2, y2, z2)
    az3 = k * m1 * (z1 - z3) / L(x3, y3, z3, x1, y1, z1) + k * m2 * (z2 - z3) / L(x3, y3, z3, x2, y2, z2)
    
    If Form1.Check2.Value Then
        If hitTest(x1, y1, z1, x2, y2, z2) Then
            ax1 = ax1 - kk * (x2 - x1) * (10 - Lenth(x1, y1, z1, x2, y2, z2)) / Lenth(x1, y1, z1, x2, y2, z2) / m1
            ay1 = ay1 - kk * (y2 - y1) * (10 - Lenth(x1, y1, z1, x2, y2, z2)) / Lenth(x1, y1, z1, x2, y2, z2) / m1
            az1 = az1 - kk * (z2 - z1) * (10 - Lenth(x1, y1, z1, x2, y2, z2)) / Lenth(x1, y1, z1, x2, y2, z2) / m1
        
            ax2 = ax2 - kk * (x1 - x2) * (10 - Lenth(x1, y1, z1, x2, y2, z2)) / Lenth(x1, y1, z1, x2, y2, z2) / m2
            ay2 = ay2 - kk * (y1 - y2) * (10 - Lenth(x1, y1, z1, x2, y2, z2)) / Lenth(x1, y1, z1, x2, y2, z2) / m2
            az2 = az2 - kk * (z1 - z2) * (10 - Lenth(x1, y1, z1, x2, y2, z2)) / Lenth(x1, y1, z1, x2, y2, z2) / m2
        End If
        
        If hitTest(x1, y1, z1, x3, y3, z3) Then
            ax1 = ax1 - kk * (x3 - x1) * (10 - Lenth(x1, y1, z1, x3, y3, z3)) / Lenth(x1, y1, z1, x3, y3, z3) / m1
            ay1 = ay1 - kk * (y3 - y1) * (10 - Lenth(x1, y1, z1, x3, y3, z3)) / Lenth(x1, y1, z1, x3, y3, z3) / m1
            az1 = az1 - kk * (z3 - z1) * (10 - Lenth(x1, y1, z1, x3, y3, z3)) / Lenth(x1, y1, z1, x3, y3, z3) / m1
            
            ax3 = ax3 - kk * (x1 - x3) * (10 - Lenth(x1, y1, z1, x3, y3, z3)) / Lenth(x1, y1, z1, x3, y3, z3) / m3
            ay3 = ay3 - kk * (y1 - y3) * (10 - Lenth(x1, y1, z1, x3, y3, z3)) / Lenth(x1, y1, z1, x3, y3, z3) / m3
            az3 = az3 - kk * (z1 - z3) * (10 - Lenth(x1, y1, z1, x3, y3, z3)) / Lenth(x1, y1, z1, x3, y3, z3) / m3
            
        End If
        
        If hitTest(x3, y3, z3, x2, y2, z2) Then
            ax2 = ax2 - kk * (x3 - x2) * (10 - Lenth(x3, y3, z3, x2, y2, z2)) / Lenth(x3, y3, z3, x2, y2, z2) / m2
            ay2 = ay2 - kk * (y3 - y2) * (10 - Lenth(x2, y2, z2, x3, y3, z3)) / Lenth(x3, y3, z3, x2, y2, z2) / m2
            az2 = az2 - kk * (z3 - z2) * (10 - Lenth(x2, y2, z2, x3, y3, z3)) / Lenth(x3, y3, z3, x2, y2, z2) / m2
            
            ax3 = ax3 - kk * (x2 - x3) * (10 - Lenth(x3, y3, z3, x2, y2, z2)) / Lenth(x3, y3, z3, x2, y2, z2) / m3
            ay3 = ay3 - kk * (y2 - y3) * (10 - Lenth(x3, y3, z3, x2, y2, z2)) / Lenth(x3, y3, z3, x2, y2, z2) / m3
            az3 = az3 - kk * (z2 - z3) * (10 - Lenth(x3, y3, z3, x2, y2, z2)) / Lenth(x3, y3, z3, x2, y2, z2) / m3
        End If
    
    End If
    vx1 = vx1 + ax1 * dt
    vy1 = vy1 + ay1 * dt
    vz1 = vz1 + az1 * dt
    
    vx2 = vx2 + ax2 * dt
    vy2 = vy2 + ay2 * dt
    vz2 = vz2 + az2 * dt
    
    vx3 = vx3 + ax3 * dt
    vy3 = vy3 + ay3 * dt
    vz3 = vz3 + az3 * dt
    
    x1 = x1 + vx1 * dt + 0.5 * ax1 * dt * dt
    y1 = y1 + vy1 * dt + 0.5 * ay1 * dt * dt
    z1 = z1 + vz1 * dt + 0.5 * az1 * dt * dt
    
    x2 = x2 + vx2 * dt + 0.5 * ax2 * dt * dt
    y2 = y2 + vy2 * dt + 0.5 * ay2 * dt * dt
    z2 = z2 + vz2 * dt + 0.5 * az2 * dt * dt
    
    x3 = x3 + vx3 * dt + 0.5 * ax3 * dt * dt
    y3 = y3 + vy3 * dt + 0.5 * ay3 * dt * dt
    z3 = z3 + vz3 * dt + 0.5 * az3 * dt * dt
    
    t = t + dt
End Sub
Function L(fx1, fy1, fz1, fx2, fy2, fz2) As Double
    L = ((fx1 - fx2) * (fx1 - fx2) + (fy1 - fy2) * (fy1 - fy2) + (fz1 - fz2) * (fz1 - fz2)) ^ (3 / 2)
End Function
Sub atstart()
start = False
x1 = 0
y1 = 0
z1 = 0

x2 = 30
y2 = 30
z2 = 30

x3 = -30
y3 = -30
z3 = 30
m1 = 10: m2 = 30: m3 = 20
End Sub
Function Lenth(fx1, fy1, fz1, fx2, fy2, fz2) As Double
    Lenth = ((fx1 - fx2) * (fx1 - fx2) + (fy1 - fy2) * (fy1 - fy2) + (fz1 - fz2) * (fz1 - fz2)) ^ (1 / 2)
End Function
Function hitTest(fx1, fy1, fz1, fx2, fy2, fz2) As Boolean
Dim L As Double
L = Lenth(fx1, fy1, fz1, fx2, fy2, fz2)
If L < 10 Then hitTest = True Else hitTest = False
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
