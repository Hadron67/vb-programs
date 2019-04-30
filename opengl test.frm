VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   8835
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8280
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
keys(KeyCode) = True
End Sub
Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
keys(KeyCode) = False
End Sub

Private Sub Form_Load()
Call EnableOpenGL(Picture1.hDC)
Call setlight
i = 0

glMatrixMode (mmProjection)
    glLoadIdentity
    glOrtho -100#, 100#, -100#, 100#, -100#, 100#
    gluPerspective 0.6, 1, 0.1, 100
    
    glMatrixMode (GL_MODELVIEW)
    glLoadIdentity
    gluLookAt 0, 0, 100, 0, 0, 0, 0, 1, 0
    glShadeModel (GL_FLAT)
    Call glLightfv(GL_LIGHT0, GL_AMBIENT, LightAmbient(0))
    Call glLightfv(GL_LIGHT0, GL_DIFFUSE, LightDiffuse(0))
    Call glLightfv(GL_LIGHT0, GL_POSITION, LightPosition(0))
    Call glLightfv(GL_LIGHT0, GL_SPECULAR, lightSpecular(0))
    glEnable (GL_LIGHTING)
    glEnable (GL_LIGHT0)
    glEnable (GL_COLOR_MATERIAL)
    x = 0
    y = 1
    z = 0
    lx = ly = lz = 0
End Sub
Private Sub Display(angle As Double)
    Picture1.AutoRedraw = False
    glClearColor 1#, 1#, 1#, 1#   '清空颜色缓存的RGBA颜色值
    glClear clrColorBufferBit     '为绘下帧曲面清除缓冲区
    glColor3f 0.8, 0.3, 0.5       '设置显示的字体颜色
    If keys(65) Then glRotatef 2, 0, 0, 1
    If keys(68) Then glRotatef -2, 0, 0, 1
    If keys(87) Then glRotatef 2, 1, 0, 0
    If keys(83) Then glRotatef -2, 1, 0, 0
    x = 0
    z = 0
    y = 1
    t = 0
    
    
    glPushMatrix                  '依据当前模式（模式-视图矩阵）使矩阵入栈
    'glRotatef angle, 0, 1, 1
    glColor3f 1, 0, 0
    'glutSolidSphere 10, 100, 100
    Call axis
    
    
    
    
     For t = 0 To 10 Step 0.001
    vx = -10 * (x - y)
    vy = -x * z + 28 * x - y * 2
    vz = x * y - (8 / 3) * z
    x = x + vx / 100
    y = y + vy / 100
    z = z + vz / 100
    Lineto x * 1, y * 1, z * 1
    Next t
    
    
    
    'glRotatef angle, 0, 1, 1
    glPopMatrix                           '依据当前模式（模式-视图矩阵）使矩阵出栈
    SwapBuffers Picture1.hDC             '切换缓存
    Picture1.AutoRedraw = True
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


Private Sub Picture1_Click()
Call Display(90)
End Sub

Private Sub Timer1_Timer()
Call Display(i)
i = i + 1
'Picture1.Cls
End Sub




