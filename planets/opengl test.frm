VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo of Planets Movements(OpenGL)"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   DrawStyle       =   1  'Dash
   Icon            =   "opengl test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   8835
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "Help"
      Height          =   375
      Left            =   7200
      TabIndex        =   21
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "set"
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   7560
      TabIndex        =   19
      Text            =   "5"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   7560
      TabIndex        =   17
      Text            =   "10"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "delect"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "add"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "select"
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   7560
      TabIndex        =   12
      Text            =   "1000"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Collision"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   7560
      TabIndex        =   6
      Text            =   "0.01"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   7560
      TabIndex        =   5
      Text            =   "10"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "load"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Velocity"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start/pause"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
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
   Begin VB.Label Label2 
      Caption         =   "R="
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "M="
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "k="
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "T=0"
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "dt="
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "G="
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   5040
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
start = Not start
End Sub

Private Sub Command2_Click()
Call save
End Sub

Private Sub Command3_Click()
Call load
End Sub

Private Sub Command4_Click()
selectball = selectball + 1
If selectball > balln Then selectball = 1
Text1.Text = ball(selectball).m
Text2.Text = ball(selectball).radius

End Sub

Private Sub Command5_Click()
If balln >= 100 Then
MsgBox "only allowed 100 planets."
Exit Sub
End If
balln = balln + 1


End Sub

Private Sub Command6_Click()
If balln <= 1 Then Exit Sub
ball(selectball) = ball(balln)
ball(balln) = zero
balln = balln - 1
End Sub

Private Sub Command7_Click()
ball(selectball).m = Val(Text1.Text)
ball(selectball).radius = Val(Text2.Text)

End Sub

Private Sub Command8_Click()
MsgBox "操作说明" & vbCrLf & "AD-SW: 旋转视图" & vbCrLf & "方向键：平移" & vbCrLf & "数字方向键：平移球" & vbCrLf & "ctrl+数字方向键：设置初速度"
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
angle1 = 0
angle2 = 0
Call EnableOpenGL(Picture1.hDC)
Call init
Call atstart
save
End Sub




Private Sub Timer1_Timer()
Call Display

End Sub




