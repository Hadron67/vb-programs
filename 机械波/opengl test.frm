VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mechanic Wave Demo(OpenGL)"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   8835
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "说明"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "复位"
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   7320
      TabIndex        =   12
      Text            =   "1"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CheckBox Check3 
      Caption         =   "干涉板"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "波源2"
      Height          =   495
      Left            =   7200
      TabIndex        =   10
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Text            =   "25"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Text            =   "25"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Text            =   "50"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Text            =   "1"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Text            =   "1"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Text            =   "50"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Text            =   "50"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Text            =   "50"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "波源1"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
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

Private Sub Command1_Click()
initb
End Sub

Private Sub Command2_Click()

    MsgBox "操作说明" & vbCrLf & "AD-SW: 旋转视图" & vbCrLf & "方向键：平移"
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
t = 0
angle1 = 0
angle2 = 0
Call EnableOpenGL(Picture1.hDC)
Call init

End Sub



Private Sub Timer1_Timer()
Call Display

End Sub





