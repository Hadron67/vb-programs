VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "单摆"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5385
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "参数输入"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "形象情景"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "函数图像"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相空间图形"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Made by CFY"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
dt = 0.001

Scale (-5, 5)-(5, -5)
Cls
For t = 0 To 1000 Step 0.01
ax = -10 / l * Sin(x) + f * Sin(w * t) - k * vx
vx = vx + ax * dt
x = x + vx * dt + 0.5 * ax * dt * dt
PSet (x, vx)
Next t

End Sub
Public Sub lineto(x As Double, y As Double)
Line (x1, x2)-(x, y)
x1 = x
x2 = y
End Sub

Private Sub Command2_Click()
dt = 0.001
Cls
Scale (0, -10)-(300, 10)
For t = 0 To 1000 Step 0.01
ax = -10 / l * Sin(x) + f * Sin(w * t) - k * vx
vx = vx + ax * dt
x = x + vx * dt + 0.5 * ax * dt * dt
PSet (t, x)
Next t

End Sub

Private Sub Command3_Click()
If Timer1.Enabled = True Then Timer1.Enabled = False Else Timer1.Enabled = True
t = 0
End Sub

Private Sub Command4_Click()
l = Val(InputBox("输入摆长l=", "参数输入", 1))
x = Val(InputBox("输入初位置x0=", "参数输入", 1))
vx = Val(InputBox("输入初速度v0=", "参数输入", 0))
f = Val(InputBox("输入驱动力F=", "参数输入", 0))
k = Val(InputBox("输入阻力系数k=", "参数输入", 0))
w = Val(InputBox("输入驱动力角速度w=", "参数输入", 0.5))
End Sub

Private Sub Form_Load()
l = 1
vx = 0
x = 1
ax = 0
w = 0.5
k = 0

End Sub

Private Sub Timer1_Timer()
dt = 0.01
Cls
Scale (-l * 3, -l * 3)-(l * 3, l * 3)

ax = -10 / l * Sin(x) + f * Sin(w * t) - k * vx
vx = vx + ax * dt
x = x + vx * dt + 0.5 * ax * dt * dt
Line (0, 0)-(l * 2 * Sin(x), l * 2 * Cos(x))
t = t + 0.01


End Sub
