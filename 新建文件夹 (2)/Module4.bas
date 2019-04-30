Attribute VB_Name = "Module4"
Global sx(10000) As Double, sy(10000) As Double, sz(10000) As Double
Global n As Double
Sub draw1()
On Error Resume Next
Dim th As Double, t1 As Integer, t2 As Integer
Dim A As Double, B As Double, C As Double, d As Double, thita As Double
th = 0
 thita = Form1.HScroll3.Value / 1000
For th = 0 To 6.28 Step 0.01
    A = Form1.HScroll1.Value / 100 * Cos(th)
    B = Form1.HScroll1.Value / 100 * Sin(th)
    C = Form1.HScroll2.Value / 100 * Cos(th + thita)
    d = Form1.HScroll2.Value / 100 * Sin(th + thita)

    Call getpoint(A, B, C, d, Form1.VScroll2.Value / 100, Form1.VScroll3.Value / 100, Form1.VScroll1.Value / 100)
    glTranslatef x * 5, y * 5, z * 5
    glColor3f 0, 0.5, 1
    glutSolidSphere 1, 5, 5
    glTranslatef -x * 5, -y * 5, -z * 5
Next th
Form1.Picture2.Cls
For t1 = -10 To 10 Step 1
    For t2 = -10 To 10 Step 1
        vx = A * t1 + B * t2
        vy = C * t1 + d * t2
       Form1.Picture2.PSet (vx, vy)
        
    Next t2
Next t1

End Sub
Sub draw2()
On Error Resume Next
Dim th As Double
Dim A As Double, B As Double, C As Double, d As Double, thita As Double
Dim x1 As Double, y1 As Double, z1 As Double, t1 As Double
Dim ax As Double, ay As Double, bx As Double, by As Double
Dim h1 As Integer, h2 As Integer

th = 0
 thita = Form1.HScroll3.Value / 1000
     A = Form1.HScroll1.Value / 100 * Cos(th)
    B = Form1.HScroll1.Value / 100 * Sin(th)
    C = Form1.HScroll2.Value / 100 * Cos(th + thita)
    d = Form1.HScroll2.Value / 100 * Sin(th + thita)
    Call g1(A, B, C, d)
    Call g2(A, B, C, d)
    ax = g1x
    ay = g1y
    bx = g2x
    by = g2y

For th = 0 To 12.56 Step 0.01
    x1 = ax * Cos(4 * th) - ay * Sin(4 * th)
    y1 = ax * Sin(4 * th) + ay * Cos(4 * th)
    z1 = bx * Cos(6 * th) - by * Sin(6 * th)
    t1 = bx * Sin(6 * th) + by * Cos(6 * th)

    Call getpoint2(x1, y1, z1, t1, Form1.VScroll2.Value / 100, Form1.VScroll3.Value / 100, Form1.VScroll1.Value / 100)
    glTranslatef x * 5, y * 5, z * 5
    
    glColor3f 0, 0.5, 1
    glutSolidSphere 1, 5, 5
    glTranslatef -x * 5, -y * 5, -z * 5
Next th
Form1.Picture2.Cls
For h1 = -10 To 10 Step 1
    For h2 = -10 To 10 Step 1
        vx = A * h1 + B * h2
        vy = C * h1 + d * h2
       Form1.Picture2.PSet (vx, vy)
        
    Next h2
Next h1

End Sub
Sub seifert()
On Error Resume Next
Dim x As Double, y As Double, z As Double, t As Double
Dim x1 As Double, y1 As Double, z1 As Double
Dim A As Double, B As Double, C As Double
For x1 = -5 To 5 Step 0.1
    For y1 = -5 To 5 Step 0.1
        For z1 = -5 To 5 Step 0.1
            A = x1 * x1 + y1 * y1 + z1 * z1 + 4
            x = 4 * x1 / A
            y = 4 * y1 / A
            z = 4 * z1 / A
            t = (A - 8) / A
            B = x * x * x - 3 * x * y * y - 27 * z * z + 27 * t * t
            C = 3 * x * x * y - 3 * y * y * y - 54 * z * t
            DoEvents
            If (C > 0) And (Abs(B) < 0.2) Then
            Call putinto(x1, y1, z1)
            End If
        Next z1
    Next y1
Next x1
End Sub
Sub putinto(x1, y1, z1)
sx(n + 1) = x1
sy(n + 1) = y1
sz(n + 1) = z1
n = n + 1
End Sub
Sub draw3()
Dim t As Integer
For t = 1 To n Step 1
glTranslatef sx(t) * 5, sy(t) * 5, sz(t) * 5
glColor3f 0, 0.5, 1
glutSolidSphere 1, 5, 5


glTranslatef -sx(t) * 5, -sy(t) * 5, -sz(t) * 5

Next t









End Sub
