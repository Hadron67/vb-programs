Attribute VB_Name = "Module2"
Public Sub Display()

    
    glClearColor 1#, 1#, 1#, 1#   '�����ɫ�����RGBA��ɫֵ
    glClear clrColorBufferBit Or clrDepthBufferBit    'Ϊ����֡�������������
    glColor3f 0.8, 0.3, 0.5       '������ʾ��������ɫ
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
    
    
    
    
    
    
    glPushMatrix                  '���ݵ�ǰģʽ��ģʽ-��ͼ����ʹ������ջ
    
    
    If Form1.Check1.Value Then Call axis
    
    
    draw
    'add code here
   
    
    
    
    
    
    glPopMatrix                           '���ݵ�ǰģʽ��ģʽ-��ͼ����ʹ�����ջ
    SwapBuffers Form1.Picture1.hDC             '�л�����
    'Picture1.AutoRedraw = True
    
    End Sub
