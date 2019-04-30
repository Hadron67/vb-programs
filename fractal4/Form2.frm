VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "动画生成"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4470
   ScaleWidth      =   7590
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command9 
      Caption         =   "取点"
      Height          =   375
      Left            =   4800
      TabIndex        =   35
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "取点"
      Height          =   375
      Left            =   4800
      TabIndex        =   34
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "动态M集合"
      Height          =   495
      Left            =   6480
      TabIndex        =   33
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Text            =   "1"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   6480
      TabIndex        =   30
      Text            =   "2"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   2640
      TabIndex        =   26
      Text            =   "2"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Mandelbrot集合"
      Height          =   495
      Left            =   4680
      TabIndex        =   25
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   2640
      TabIndex        =   23
      Text            =   "2"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "动态Julia集合"
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Text            =   "0"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "取点"
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Text            =   "0"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Text            =   "0"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取点"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Text            =   "0"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Text            =   "0"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Julia集合"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Text            =   "2"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Text            =   "1.2"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "1.2"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "2"
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mandelbrot集合"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   32
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   0
      Y2              =   4440
   End
   Begin VB.Label Label10 
      Caption         =   "总张数"
      Height          =   375
      Left            =   2040
      TabIndex        =   29
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "总张数"
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   27
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   24
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   1920
      X2              =   1920
      Y1              =   2160
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   1920
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Label Label6 
      Caption         =   "倍数"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "总张数"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "倍数"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "总张数"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Call disable
Dim t, zs As Integer, db As Double
zs = Val(Text1.Text)
db = Val(Text2.Text)
Label2.Caption = "0/" & zs
For t = 1 To zs Step 1
Form1.Text5.Text = Form1.Text5.Text * db
Call Form1.Command2_Click
Call Form1.Command4_Click
Label2.Caption = t & "/" & zs
DoEvents
Next t
Call able
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim t, zs As Integer, db As Double
Call disable
zs = Val(Text4.Text)
db = Val(Text3.Text)
Label4.Caption = "0/" & zs
For t = 1 To zs Step 1
Form1.Text7.Text = Form1.Text7.Text * db
Call Form1.Command1_Click
Call Form1.Command6_Click
Label4.Caption = t & "/" & zs
DoEvents
Next t
Call able
End Sub

Private Sub Command3_Click()
Text7.Text = h2 / 1000
Text8.Text = v2 / 1000

End Sub

Private Sub Command4_Click()
Text11.Text = h2 / 1000
Text12.Text = v2 / 1000
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim x1, x2, y1, y2 As Double, zs As Integer, t As Double, c As Integer
Call disable
c = 0
x1 = Val(Text7.Text)
y1 = Val(Text8.Text)
x2 = Val(Text5.Text)
y2 = Val(Text6.Text)
zs = Val(Text13.Text) - 1
If zs <= 0 Then
    MsgBox ("张数>=2")
Else
    Label7.Caption = "0/" & zs + 1
    For t = 0 To 1 Step 1 / zs
        Form1.Text1.Text = (1 - t) * x1 + t * x2
        Form1.Text3.Text = (1 - t) * y1 + t * y2
        Call Form1.Command1_Click
        Call Form1.Command6_Click
        c = c + 1
        Label7.Caption = c & "/" & zs + 1
        DoEvents
    Next t
End If
Call able
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim x1, x2, y1, y2 As Double, zs As Integer, t As Double, c As Integer
Call disable
c = 0
x1 = Val(Text11.Text)
y1 = Val(Text13.Text)
x2 = Val(Text9.Text)
y2 = Val(Text10.Text)
zs = Val(Text14.Text) - 1
If zs <= 0 Then
    MsgBox ("张数>=2")
Else
    Label7.Caption = "0/" & zs + 1
    For t = 0 To 1 Step 1 / zs
        Form1.Text2.Text = (1 - t) * x1 + t * x2
        Form1.Text4.Text = (1 - t) * y1 + t * y2
        Call Form1.Command2_Click
        Call Form1.Command4_Click
        c = c + 1
        Label8.Caption = c & "/" & zs + 1
        DoEvents
    Next t
End If
Call able
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim zs As Integer, dp As Double, t As Integer
Call disable
zs = Val(Text15.Text)
dp = Val(Text16.Text)
Label11.Caption = "0/" & zs
For t = 1 To zs Step 1
    
    Call Form1.Command2_Click
    Call Form1.Command4_Click
    Form1.Text11.Text = Form1.Text11.Text + dp
    Label11.Caption = t & "/" & zs
    DoEvents
Next t
Call able
End Sub

Private Sub Command8_Click()
Text5.Text = h2 / 1000
Text6.Text = v2 / 1000

End Sub

Private Sub Command9_Click()
Text9.Text = h2 / 1000
Text10.Text = v2 / 1000

End Sub

Private Sub able()
Form1.Text1.Enabled = True
Form1.Text3.Enabled = True
Form1.Text7.Enabled = True
Form1.Text2.Enabled = True
Form1.Text4.Enabled = True
Form1.Text5.Enabled = True
Form1.Text6.Enabled = True
Form1.Text9.Enabled = True
Form1.Command4.Enabled = True
Command2.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command4.Enabled = True
Command1.Enabled = True
Form1.Text11.Enabled = True
Form1.Text9.Enabled = True
Command7.Enabled = True
Form1.Command4.Enabled = True
End Sub
Private Sub disable()
Form1.Text1.Enabled = False
Form1.Text3.Enabled = False
Form1.Text7.Enabled = False
Form1.Text6.Enabled = False
Form1.Text8.Enabled = False
Form1.Text11.Enabled = False
Form1.Command6.Enabled = False
Form1.Command4.Enabled = False
Command2.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command4.Enabled = False
Command1.Enabled = False
Command7.Enabled = False
Form1.Text9.Enabled = False
Form1.Text1.Enabled = False
Form1.Text3.Enabled = False
Form1.Text7.Enabled = False
Form1.Text2.Enabled = False
Form1.Text4.Enabled = False
Form1.Text5.Enabled = False

End Sub



