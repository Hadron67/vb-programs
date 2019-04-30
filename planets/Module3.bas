Attribute VB_Name = "Module3"
Public Sub Display()
    On Error Resume Next
    glClearColor 0#, 0#, 0#, 1#   '清空颜色缓存的RGBA颜色值
    glClear clrColorBufferBit Or clrDepthBufferBit    '为绘下帧曲面清除缓冲区
    glColor3f 0.8, 0.3, 0.5      '设置显示的字体颜色
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
    Form1.Label6.Caption = "T=" & Int(t * 100) / 100
    
    glPopMatrix                           '依据当前模式（模式-视图矩阵）使矩阵出栈
    SwapBuffers Form1.Picture1.hDC             '切换缓存
    End Sub
Private Sub disballs()
Dim m As Integer
For m = 1 To balln Step 1
    If m = selectball Then glColor3f 0, 1, 1 Else glColor3f 0, 1, 0
    glTranslatef ball(m).x, ball(m).y, ball(m).z
    glutSolidSphere ball(m).radius, 20, 20
    glTranslatef -ball(m).x, -ball(m).y, -ball(m).z
Next m
End Sub
Private Sub viewvelocity(k As Double)
Dim m As Double
Dim vvx As Double, vvy As Double, vvz As Double

For m = 1 To balln Step 1
vvx = ball(m).vx * ball(m).radius / Lenth(0, 0, 0, ball(m).vx, ball(m).vy, ball(m).vz)
vvy = ball(m).vy * ball(m).radius / Lenth(0, 0, 0, ball(m).vx, ball(m).vy, ball(m).vz)
vvz = ball(m).vz * ball(m).radius / Lenth(0, 0, 0, ball(m).vx, ball(m).vy, ball(m).vz)


    glBegin bmLines
    glVertex3f ball(m).x, ball(m).y, ball(m).z
    glVertex3f ball(m).x + k * ball(m).vx + vvx, ball(m).y + k * ball(m).vy + vvy, ball(m).z + k * ball(m).vz + vvz
    glEnd
Next m
End Sub
Private Sub getval()
 If Not start Then
       
            If keys(17) Then
                If keys(102) Then ball(selectball).vx = ball(selectball).vx + 0.1
                If keys(104) Then ball(selectball).vy = ball(selectball).vy + 0.1
                If keys(101) Then ball(selectball).vz = ball(selectball).vz + 0.1
                If keys(100) Then ball(selectball).vx = ball(selectball).vx - 0.1
                If keys(98) Then ball(selectball).vy = ball(selectball).vy - 0.1
                If keys(97) Then ball(selectball).vz = ball(selectball).vz - 0.1
            Else
                If keys(102) Then ball(selectball).x = ball(selectball).x + 0.1
                If keys(104) Then ball(selectball).y = ball(selectball).y + 0.1
                If keys(101) Then ball(selectball).z = ball(selectball).z + 0.1
                If keys(100) Then ball(selectball).x = ball(selectball).x - 0.1
                If keys(98) Then ball(selectball).y = ball(selectball).y - 0.1
                If keys(97) Then ball(selectball).z = ball(selectball).z - 0.1
            
            End If
        
       
       
    End If
End Sub
Sub save()
Dim m As Integer
For m = 1 To 100 Step 1
ball2(m) = ball(m)
Next m
balln2 = balln
t = 0
End Sub
Sub load()
Dim m As Integer
For m = 1 To 100 Step 1
ball(m) = ball2(m)
Next m
t = 0
balln = balln2

End Sub


