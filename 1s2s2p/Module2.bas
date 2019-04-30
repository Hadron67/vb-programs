Attribute VB_Name = "Module2"
Public Sub Display()

    
    glClearColor 1#, 1#, 1#, 1#   '清空颜色缓存的RGBA颜色值
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
    
    
    
    
    
    
    glPushMatrix                  '依据当前模式（模式-视图矩阵）使矩阵入栈
    
    
    If Form1.Check1.Value Then Call axis
    
    
    draw
    'add code here
   
    
    
    
    
    
    glPopMatrix                           '依据当前模式（模式-视图矩阵）使矩阵出栈
    SwapBuffers Form1.Picture1.hDC             '切换缓存
    'Picture1.AutoRedraw = True
    
    End Sub
