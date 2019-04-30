VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo of Three Planets(OpenGL)"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   DrawStyle       =   1  'Dash
   Icon            =   "opengl test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   8835
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   7560
      TabIndex        =   21
      Text            =   "1000"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Collision"
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   7560
      TabIndex        =   15
      Text            =   "0.01"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   7560
      TabIndex        =   14
      Text            =   "10"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   7560
      TabIndex        =   10
      Text            =   "20"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   7560
      TabIndex        =   9
      Text            =   "30"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   7560
      TabIndex        =   8
      Text            =   "10"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "load"
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "m3"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "m2"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "m1"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Velocity"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   1560
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
   Begin VB.Label Label7 
      Caption         =   "k="
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "T=0"
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "dt="
      Height          =   255
      Left            =   7200
      TabIndex        =   17
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "G="
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "m3="
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "m2="
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "m1="
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   3840
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




