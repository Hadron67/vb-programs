Attribute VB_Name = "Module3"
Function wave(x, y)
On Error Resume Next
Dim r As Double, s As Double, c As Double
s = 0
c = 0

r = Sqr(x * x + y * y)
s = y / r
c = x / r
wave = r * (s) * Exp(-r / 5) * 30
End Function
Sub color(x As Double)
If x >= 0 Then
glColor3f 255, 255 * Exp(-x * 5), 255 * Exp(-x * 5)
ElseIf x < 0 Then
glColor3f 255 * Exp(x * 5), 255 * Exp(x * 5), 255
End If
End Sub
Sub draw()
Dim x As Double, y As Double, dt As Double
dt = 0.5
For x = -40 To 40 Step 1
    For y = -40 To 40 Step 1
       
       glColor3f 0, 0, 0
        glBegin bmLines
       
        glVertex3f x, y, wave(x, y) + dt
        
        glVertex3f x, y + 1, wave(x, y + 1) + dt
        
        
        glVertex3f x, y, wave(x, y) + dt
        
        glVertex3f x + 1, y, wave(x + 1, y) + dt
        
        
        glVertex3f x, y, wave(x, y) - dt
        
        glVertex3f x, y + 1, wave(x, y + 1) - dt
        
        
        glVertex3f x, y, wave(x, y) - dt
        
        glVertex3f x + 1, y, wave(x + 1, y) - dt
        glEnd
        
        glBegin bmTriangles
        color (wave(x, y))
        glVertex3f x, y, wave(x, y)
        
        color (wave(x, y + 1))
        glVertex3f x, y + 1, wave(x, y + 1)
        
        color (wave(x + 1, y))
        glVertex3f x + 1, y, wave(x + 1, y)
        
        
        color (wave(x, y + 1))
        glVertex3f x, y + 1, wave(x, y + 1)
        
        color (wave(x + 1, y))
        glVertex3f x + 1, y, wave(x + 1, y)
        
        color (wave(x + 1, y + 1))
        glVertex3f x + 1, y + 1, wave(x + 1, y + 1)
        glEnd
        
        glBegin bmTriangles
        color (wave(x, y))
        glVertex3f x, y, -70
        
        color (wave(x, y + 1))
        glVertex3f x, y + 1, -70
        
        color (wave(x + 1, y))
        glVertex3f x + 1, y, -70
        
        
        color (wave(x, y + 1))
        glVertex3f x, y + 1, -70
        
        color (wave(x + 1, y))
        glVertex3f x + 1, y, -70
        
        color (wave(x + 1, y + 1))
        glVertex3f x + 1, y + 1, -70
        glEnd
    Next y
Next x
End Sub
