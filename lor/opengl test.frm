VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "lorentz attractor demo (opengl)"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9075
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option2 
      Caption         =   "attractor2"
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   6360
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "attractor1"
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   5880
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "重启"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "显示/隐藏吸引子(O)"
      Height          =   615
      Left            =   7200
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   7200
      TabIndex        =   9
      Text            =   "0.001"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始/暂停(P)"
      Height          =   615
      Left            =   7200
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   7200
      TabIndex        =   7
      Text            =   "28"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   7200
      TabIndex        =   6
      Text            =   "2.66666666667"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   7200
      TabIndex        =   5
      Text            =   "10"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "复位(R)"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   7200
      TabIndex        =   3
      Text            =   "10"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   7200
      TabIndex        =   2
      Text            =   "10"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   7200
      TabIndex        =   1
      Text            =   "10"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8640
      Top             =   240
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
   Begin VB.Label Label1 
      Caption         =   "By CFY"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call placeballs(Val(Text1.Text), Val(Text2.Text), Val(Text3.Text))
Call copy
End Sub

Private Sub Command2_Click()
enable = Not enable
End Sub

Private Sub Command3_Click()
enable2 = Not enable2
End Sub

Private Sub Command4_Click()
ret = Shell(App.Path & "\" & App.EXEName & ".exe")
End
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
keys(KeyCode) = True
End Sub
Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
keys(KeyCode) = False
End Sub

Private Sub Form_Load()
lx = 0
ly = 0
lz = 0
Call EnableOpenGL(Picture1.hDC)
Call init
Call placeballs(10, 10, 10)
Call copy
Form1.Show
End Sub
Private Sub display(angle As Double)
On Error Resume Next
    Picture1.AutoRedraw = False
    glClearColor 1#, 1#, 1#, 1#   '清空颜色缓存的RGBA颜色值
    glClear clrColorBufferBit Or clrDepthBufferBit     '为绘下帧曲面清除缓冲区
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
    
    If keys(102) And Not enable Then
        Text1.Text = Text1.Text + 1
        Command1_Click
    End If
    If keys(100) And Not enable Then
        Text1.Text = Text1.Text - 1
        Command1_Click
    End If
    If keys(104) And Not enable Then
        Text2.Text = Text2.Text + 1
        Command1_Click
    End If
    If keys(98) And Not enable Then
        Text2.Text = Text2.Text - 1
        Command1_Click
    End If
    If keys(101) And Not enable Then
        Text3.Text = Text3.Text + 1
        Command1_Click
    End If
    If keys(97) And Not enable Then
        Text3.Text = Text3.Text - 1
        Command1_Click
    End If
    If keys(80) Then
        keys(80) = False
        enable = Not enable
    End If
    If keys(79) Then
        keys(79) = False
        enable2 = Not enable2
    End If
    If keys(82) Then
        keys(82) = False
        Command1_Click
    End If
    
    
    
    
    
    
    
    
    
    
    
    glPushMatrix                  '依据当前模式（模式-视图矩阵）使矩阵入栈
    
    
    
    Call axis
    
    
    glColor3f 0, 1, 1
    'add code here
    If Option1.Value Then
    Call displayballs(Val(Text4.Text), Val(Text5.Text), Val(Text6.Text), Val(Text7.Text))
    Call displayattractor(0, 1, 0, Val(Text4.Text), Val(Text5.Text), Val(Text6.Text))
    ElseIf Option2.Value Then
    Call displayballs2(Val(Text4.Text), Val(Text7.Text))
    Call displayattractor2(0, 1, 0, Val(Text4.Text))
    End If
    
    
    
    
    glPopMatrix                           '依据当前模式（模式-视图矩阵）使矩阵出栈
    SwapBuffers Picture1.hDC             '切换缓存
    Picture1.AutoRedraw = True
    
    End Sub



Private Sub Timer1_Timer()
Call display(i)

End Sub




