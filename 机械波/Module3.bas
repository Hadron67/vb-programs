Attribute VB_Name = "Module3"
Sub Display()
    Dim y As Double
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
    
    
    If Form1.Check1.Value Then wave(Val(Form1.Text1.Text), Val(Form1.Text2.Text)) = Val(Form1.Text3.Text) * Sin(t * Val(Form1.Text4.Text))
    If Form1.Check2.Value Then wave(Val(Form1.Text8.Text), Val(Form1.Text7.Text)) = Val(Form1.Text6.Text) * Sin(t * Val(Form1.Text5.Text))
    If keys(80) Then wave(Val(Form1.Text1.Text), Val(Form1.Text2.Text)) = Val(Form1.Text3.Text)
    If Form1.Check3.Value Then
        For y = 0 To 100 Step 1
            If y <> 24 And y <> 25 And y <> 26 And y <> 74 And y <> 75 And y <> 76 Then wave(50, y) = 0
        Next y
    End If
            
    t = t + 0.1
    Call axis
    react 0, Val(Form1.Text9.Text), 0.1
    glColor3f 0, 0.5, 1
    disball
    
    'add code here
   
    
    
    
    
    
    glPopMatrix                           '���ݵ�ǰģʽ��ģʽ-��ͼ����ʹ�����ջ
    SwapBuffers Form1.Picture1.hDC             '�л�����
    
    End Sub

