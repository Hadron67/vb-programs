Attribute VB_Name = "Module1"
Dim hRC     As Long 'ȫ�ֱ���
Global i As Double
Global x, y, z, vx, vy, vz, lx, ly, lz As Double
Global keys(256) As Boolean
Global LightAmbient(3) As GLfloat
                                                        
Global LightDiffuse(3) As GLfloat
                                                        
Global LightPosition(3) As GLfloat

Global lightSpecular(3) As GLfloat

    '����OGL
    Sub EnableOpenGL(ghDC As Long)

    On Error GoTo Err

    Dim pfd As PIXELFORMATDESCRIPTOR                'pfd���ظ�ʽ����.
    ZeroMemory pfd, Len(pfd)
    pfd.nSize = Len(pfd)                        '��С
    pfd.nVersion = 1                            '�汾
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER '��־
    pfd.iPixelType = PFD_TYPE_RGBA              '��������
    pfd.cColorBits = 24                         '��ɫλ��
    pfd.cDepthBits = 32                         'λ��
    pfd.iLayerType = PFD_MAIN_PLANE             'ͼ������

    Dim PixFormat As Long
    PixFormat = ChoosePixelFormat(ghDC, pfd)    'ѡ���豸����ƥ�����������õ�����
    SetPixelFormat ghDC, PixFormat, pfd         '���óɵ�ǰ������
    hRC = wglCreateContext(ghDC)                '��������������
    wglMakeCurrent ghDC, hRC                    '�������ķ�������������Ϊ��ǰ
    Exit Sub

Err:
    MsgBox "Can't   create   OpenGL   context!", vbCritical, "Error"
    End
    End Sub
Sub setlight()
    LightAmbient(0) = 0.5
    LightAmbient(1) = 0.5
    LightAmbient(2) = 0.5
    LightAmbient(3) = 1#

    LightDiffuse(0) = 1#
    LightDiffuse(1) = 1#
    LightDiffuse(2) = 1#
    LightDiffuse(3) = 1#
    
    LightPosition(0) = 10#
    LightPosition(1) = 10#
    LightPosition(2) = 10#
    LightPosition(3) = 1#
    
    lightSpecular(0) = 1
    lightSpecular(1) = 1
    lightSpecular(2) = 1
    lightSpecular(3) = 1
End Sub
Sub axis()
Dim t As Integer
For t = -100 To 100 Step 10
    glColor3f 0, 1, 0
    glBegin GL_LINES
        glVertex3f t, 100, 0
        glVertex3f t, -100, 0
        
        glVertex3f 100, t, 0
        glVertex3f -100, t, 0
    glEnd
Next t
glColor3f 0.3, 0.22, 0.236
glBegin GL_LINES
        glVertex3f 0, 0, 100
        glVertex3f 0, 0, -100
        
        glVertex3f 100, 0, 0
        glVertex3f -100, 0, 0
        
glEnd
End Sub
