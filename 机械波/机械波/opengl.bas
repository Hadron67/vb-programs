Attribute VB_Name = "Module1"
Dim hRC     As Long 'ȫ�ֱ���
Global i As Double
Global lx, ly, lz As Double
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
    Sub DisableOpenGL()
    wglMakeCurrent 0, 0
    wglDeleteContext hRC
    End Sub
Sub setlight()
    LightAmbient(0) = 0.5
    LightAmbient(1) = 0.5
    LightAmbient(2) = 0.5
    LightAmbient(3) = 1#

    LightDiffuse(0) = 0#
    LightDiffuse(1) = 1#
    LightDiffuse(2) = 1#
    LightDiffuse(3) = 1#
    
    LightPosition(0) = 50#
    LightPosition(1) = 50#
    LightPosition(2) = 10#
    LightPosition(3) = 2
    
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
        If t <> 0 Then
        glVertex3f 100, t, 0
        glVertex3f -100, t, 0
        End If
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
Sub init()
Call setlight
    glMatrixMode (mmProjection)
    glLoadIdentity
    glOrtho -100#, 100#, -100#, 100#, -100#, 100#
    gluPerspective 0.6, 1, 1, 150
    
    glMatrixMode (GL_MODELVIEW)
    glLoadIdentity
    gluLookAt 0, 0, 100, 0, 0, 0, 0, 1, 0
    
    glShadeModel smSmooth
    glClearDepth -10#                     ' Depth Buffer Setup
    glEnable glcDepthTest               ' Enables Depth Testing
    glDepthFunc (cfGreater)                ' The Type Of Depth Test To Do

    glHint htPerspectiveCorrectionHint, hmNicest
    
    Call glLightfv(GL_LIGHT0, GL_AMBIENT, LightAmbient(0))
    Call glLightfv(GL_LIGHT0, GL_DIFFUSE, LightDiffuse(0))
    Call glLightfv(GL_LIGHT0, GL_POSITION, LightPosition(0))
    Call glLightfv(GL_LIGHT0, GL_SPECULAR, lightSpecular(0))
    glEnable (GL_LIGHTING)
    glEnable (GL_LIGHT0)
    glEnable (GL_COLOR_MATERIAL)
    
    lx = ly = lz = 0
End Sub
Private Sub Lineto(cx As Double, cy As Double, cz As Double)
    glColor3f 0, 0, 1
    glBegin bmLines
        glVertex3f lx, ly, lz
        glVertex3f cx, cy, cz
    glEnd
    lx = cx
    ly = cy
    lz = cz
End Sub

