Attribute VB_Name = "Module3"
Public Sub Display()
    On Error Resume Next
    glClearColor 0#, 0#, 0#, 1#   '清空颜色缓存的RGBA颜色值
    glClear clrColorBufferBit Or clrDepthBufferBit    '为绘下帧曲面清除缓冲区
    glColor3f 0.8, 0.3, 0.5       '设置显示的字体颜色
    If keys(65) Then glRotatef 2, 0, 0, 1
    If keys(68) Then glRotatef -2, 0, 0, 1
    If keys(87) Then glRotatef 2, 1, 0, 0
    If keys(83) Then glRotatef -2, 1, 0, 0
    If keys(38) Then gluLookAt 0, 0, 1, 0, 0, 0, 0, 1, 0
    If keys(40) Then
        glRotatef 180, 1, 0, 0
        gluLookAt 0, 0, 1, 0, 0, 0, 0, 1, 0
        glRotatef -180, 1, 0, 0
    End If
    If keys(39) Then
        glRotatef 90, 0, 1, 0
        gluLookAt 0, 0, 1, 0, 0, 0, 0, 1, 0
        glRotatef -90, 0, 1, 0
    End If
    If keys(37) Then
        glRotatef -90, 0, 1, 0
        gluLookAt 0, 0, 1, 0, 0, 0, 0, 1, 0
        glRotatef 90, 0, 1, 0
    End If
    
    Call getval
    glPushMatrix                  '依据当前模式（模式-视图矩阵）使矩阵入栈
    Call axis
    Call disballs
    If Form1.Check1.Value Then viewvelocity (5)
    If start Then Call react(Val(Form1.Text4.Text), Val(Form1.Text5.Text), Val(Form1.Text6.Text))
    'add code here
    Form1.Label6.Caption = "T=" & Int(t * 100) / 100
    glPopMatrix                           '依据当前模式（模式-视图矩阵）使矩阵出栈
    SwapBuffers Form1.Picture1.hDC             '切换缓存
    End Sub
Private Sub disballs()
glColor3f 1, 0, 0
glTranslatef x1, y1, z1
glutSolidSphere 5, 10, 10
glTranslatef -x1, -y1, -z1

glColor3f 0, 1, 0
glTranslatef x2, y2, z2
glutSolidSphere 5, 10, 10
glTranslatef -x2, -y2, -z2

glColor3f 0, 1, 1
glTranslatef x3, y3, z3
glutSolidSphere 5, 10, 10
glTranslatef -x3, -y3, -z3
End Sub
Private Sub viewvelocity(k As Double)
glColor3f 1, 0, 0
glBegin bmLines
    glVertex3f x1, y1, z1
    glVertex3f x1 + vx1 * k, y1 + vy1 * k, z1 + vz1 * k
glEnd

glColor3f 0, 1, 0
glBegin bmLines
    glVertex3f x2, y2, z2
    glVertex3f x2 + vx2 * k, y2 + vy2 * k, z2 + vz2 * k
glEnd

glColor3f 0, 1, 1
glBegin bmLines
    glVertex3f x3, y3, z3
    glVertex3f x3 + vx3 * k, y3 + vy3 * k, z3 + vz3 * k
glEnd
End Sub
Private Sub getval()
 If Not start Then
        If Form1.Option1.Value Then
            If keys(17) Then
                If keys(102) Then vx1 = vx1 + 0.1
                If keys(104) Then vy1 = vy1 + 0.1
                If keys(101) Then vz1 = vz1 + 0.1
                If keys(100) Then vx1 = vx1 - 0.1
                If keys(98) Then vy1 = vy1 - 0.1
                If keys(97) Then vz1 = vz1 - 0.1
            Else
                If keys(102) Then x1 = x1 + 0.1
                If keys(104) Then y1 = y1 + 0.1
                If keys(101) Then z1 = z1 + 0.1
                If keys(100) Then x1 = x1 - 0.1
                If keys(98) Then y1 = y1 - 0.1
                If keys(97) Then z1 = z1 - 0.1
            
            End If
        End If
        
        If Form1.Option2.Value Then
            If keys(17) Then
                If keys(102) Then vx2 = vx2 + 0.1
                If keys(104) Then vy2 = vy2 + 0.1
                If keys(101) Then vz2 = vz2 + 0.1
                If keys(100) Then vx2 = vx2 - 0.1
                If keys(98) Then vy2 = vy2 - 0.1
                If keys(97) Then vz2 = vz2 - 0.1
            Else
                If keys(102) Then x2 = x2 + 0.1
                If keys(104) Then y2 = y2 + 0.1
                If keys(101) Then z2 = z2 + 0.1
                If keys(100) Then x2 = x2 - 0.1
                If keys(98) Then y2 = y2 - 0.1
                If keys(97) Then z2 = z2 - 0.1
            End If
        End If
        
        If Form1.Option3.Value Then
            If keys(17) Then
                If keys(102) Then vx3 = vx3 + 0.1
                If keys(104) Then vy3 = vy3 + 0.1
                If keys(101) Then vz3 = vz3 + 0.1
                If keys(100) Then vx3 = vx3 - 0.1
                If keys(98) Then vy3 = vy3 - 0.1
                If keys(97) Then vz3 = vz3 - 0.1
            Else
                If keys(102) Then x3 = x3 + 0.1
                If keys(104) Then y3 = y3 + 0.1
                If keys(101) Then z3 = z3 + 0.1
                If keys(100) Then x3 = x3 - 0.1
                If keys(98) Then y3 = y3 - 0.1
                If keys(97) Then z3 = z3 - 0.1
            End If
        End If
       m1 = Val(Form1.Text1.Text)
       m2 = Val(Form1.Text2.Text)
       m3 = Val(Form1.Text3.Text)
       
    End If
End Sub
Sub save()
x(1) = x1
x(2) = y1
x(3) = z1

y(1) = x2
y(2) = y2
y(3) = z2

z(1) = x3
z(2) = y3
z(3) = z3

vx(1) = vx1
vx(2) = vy1
vx(3) = vz1

vy(1) = vx2
vy(2) = vy2
vy(3) = vz2

vz(1) = vx3
vz(2) = vy3
vz(3) = vz3
t = 0
End Sub
Sub load()
x1 = x(1)
y1 = x(2)
z1 = x(3)

x2 = y(1)
y2 = y(2)
z2 = y(3)

x3 = z(1)
y3 = z(2)
z3 = z(3)

vx1 = vx(1)
vy1 = vx(2)
vz1 = vx(3)

vx2 = vy(1)
vy2 = vy(2)
vz2 = vy(3)

vx3 = vz(1)
vy3 = vz(2)
vz3 = vz(3)
t = 0

End Sub


