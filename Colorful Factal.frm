VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Colorful Factal"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   Icon            =   "Colorful Factal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8809.189
   ScaleMode       =   0  'User
   ScaleWidth      =   9690
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check2 
      Caption         =   "mandelbrot"
      Height          =   180
      Left            =   4320
      TabIndex        =   41
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9240
      Top             =   6840
   End
   Begin VB.PictureBox Picture6 
      Height          =   495
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   40
      Top             =   7080
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "julia"
      Height          =   255
      Left            =   2640
      TabIndex        =   39
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   7080
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "全部停止"
      Height          =   495
      Left            =   8280
      TabIndex        =   33
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   840
      TabIndex        =   31
      Text            =   "2"
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6480
      TabIndex        =   29
      Text            =   "Text10"
      Top             =   7920
      Width           =   3135
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7320
      TabIndex        =   28
      Text            =   "1"
      Top             =   7200
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command6 
      Caption         =   "保存1"
      Height          =   495
      Left            =   6480
      TabIndex        =   23
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "动画生成"
      Height          =   375
      Left            =   8280
      TabIndex        =   26
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6480
      TabIndex        =   25
      Text            =   "1"
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "保存2"
      Height          =   495
      Left            =   7320
      TabIndex        =   24
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "说明"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5400
      TabIndex        =   18
      Text            =   "1"
      Top             =   5760
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      Height          =   4950
      Left            =   4920
      Picture         =   "Colorful Factal.frx":1CCA
      ScaleHeight     =   4890
      ScaleWidth      =   4515
      TabIndex        =   15
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   7320
      TabIndex        =   14
      Text            =   "200"
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Text            =   "1"
      Top             =   6600
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      Height          =   4935
      Left            =   120
      Picture         =   "Colorful Factal.frx":91D0C
      ScaleHeight     =   4875
      ScaleMode       =   0  'User
      ScaleWidth      =   4575
      TabIndex        =   12
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Text            =   "0"
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Text            =   "0"
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Text            =   "0"
      Top             =   6600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Text            =   "0"
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mandelbrot集合"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Julia集合"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      Height          =   4935
      Left            =   120
      Picture         =   "Colorful Factal.frx":121D4E
      ScaleHeight     =   4875
      ScaleMode       =   0  'User
      ScaleWidth      =   4575
      TabIndex        =   22
      Top             =   120
      Width           =   4575
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      Height          =   4950
      Left            =   4920
      Picture         =   "Colorful Factal.frx":1B1D90
      ScaleHeight     =   4890
      ScaleWidth      =   4515
      TabIndex        =   21
      Top             =   120
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8280
      Top             =   6840
   End
   Begin VB.Label Label15 
      Caption         =   "Y="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   37
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "X="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   35
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   34
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Power="
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   7080
      Width           =   615
   End
   Begin VB.Line Line4 
      X1              =   6360
      X2              =   6360
      Y1              =   5978.8
      Y2              =   6742.051
   End
   Begin VB.Label Label7 
      Caption         =   "保存路径"
      Height          =   255
      Left            =   6480
      TabIndex        =   30
      Top             =   7680
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   9840
      Y1              =   5978.8
      Y2              =   5978.8
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6360
      Y1              =   6742.051
      Y2              =   8904.596
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   6742.051
      Y2              =   6742.051
   End
   Begin VB.Label Label13 
      Caption         =   "放大倍数"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "迭代次数"
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "放大倍数"
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "*i"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "*i"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "+"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "+"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "S="
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "C="
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   5760
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()

On Error Resume Next

Dim se As Double
Dim xe As Double
Dim sc As Double
Dim xc As Double
Dim m As Double
Dim d As Double
Dim t As Double
Dim X As Double, Y As Double, a As Double, movex As Double, movey As Double, bs As Double, tmax As Double
Dim arg, pow As Double
pow = Val(Text11.Text)
d = 0.01
tmax = Val(Text6.Text)
sc = Val(Text1.Text)
xc = Val(Text3.Text)
bs = Val(Text7.Text)
If bs <= 0 Then
MsgBox ("放大倍数不允许为非正数！")
Else
Command1.Enabled = False
Command2.Enabled = False
Picture1.Visible = False
Picture1.Cls
movex = h1 / 1000
movey = v1 / 1000
Picture1.Scale (movex - 2 / bs, movey + 2 / bs)-(movex + 2 / bs, movey - 2 / bs)
For se = (movex - (2 / bs)) To (movex + (2 / bs)) Step 0.01 / bs
    For xe = (movey - (2 / bs)) To (movey + (2 / bs)) Step 0.01 / bs
        m = 0
        t = 0
        X = se: Y = xe
        While m < 10 And t < tmax
            m = Sqr(X * X + Y * Y)
            arg = 0
            If pow = 2 Then
                a = X
                X = X * X - Y * Y + sc
                Y = 2 * a * Y + xc
                t = t + 1
            Else
                arg = arctan(X, Y)
                arg = arg * pow
                
                m = m ^ pow
                X = m * Cos(arg) + sc
                Y = m * Sin(arg) + xc
                t = t + 1
            End If
            
        Wend
       
           Picture1.PSet (se, xe), RGB(255 - t / tmax * 255, Abs(255 - 2 * Abs(127 - t / tmax * 255)), t / tmax * 255)
    
        ProgressBar1.Value = (se - (movex - (2 / bs))) / (4 / bs) * 100
        DoEvents
    Next xe
Next se
Command1.Enabled = True
Command2.Enabled = True
Picture1.Visible = True
ProgressBar1.Value = 0

End If
End Sub

Public Sub Command2_Click()

On Error Resume Next
Dim se As Double
Dim xe As Double
Dim sc As Double
Dim xc As Double
Dim m As Double
Dim d As Double
Dim t As Double
Dim X As Double, Y As Double, a As Double, movex As Double, movey As Double, bs As Double, tmax As Double
Dim arg, pow As Double
pow = Val(Text11.Text)
d = 0.01
tmax = Val(Text6.Text)
movex = h2 / 1000
movey = v2 / 1000
bs = Val(Text5.Text)
If bs <= 0 Then
MsgBox ("放大倍数不允许为非正数！")
Else
Command1.Enabled = False
Command2.Enabled = False
Picture2.Visible = False
Picture2.Cls
Picture2.Scale (movex - 2 / bs, movey + 2 / bs)-(movex + 2 / bs, movey - 2 / bs)
For sc = (movex - (2 / bs)) To (movex + (2 / bs)) Step 0.01 / bs
    For xc = (movey - (2 / bs)) To (movey + (2 / bs)) Step 0.01 / bs
        m = 0
        t = 0
        X = Val(Text2.Text): Y = Val(Text4.Text)
        While m < 10 And t < tmax
            m = Sqr(X * X + Y * Y)
            arg = 0
            If pow = 2 Then
                a = X
                X = X * X - Y * Y + sc
                Y = 2 * a * Y + xc
                t = t + 1
            Else
                arg = arctan(X, Y)
                
                arg = arg * pow
                
                m = m ^ pow
                X = m * Cos(arg) + sc
                Y = m * Sin(arg) + xc
                t = t + 1
            End If
            
        Wend
      
            Picture2.PSet (sc, xc), RGB(255 - t / tmax * 255, Abs(255 - 2 * Abs(127 - t / tmax * 255)), t / tmax * 255)
        
ProgressBar1.Value = (sc - (movex - (2 / bs))) / (4 / bs) * 100
DoEvents
    Next xc
Next sc
Command1.Enabled = True
Command2.Enabled = True
Picture2.Visible = True
ProgressBar1.Value = 0

End If
End Sub
Private Function arctan(X As Double, Y As Double) As Double
Dim pi As Double, a As Double
pi = 3.14159265358979
a = Atn(Y / X)
If Y > 0 And X < 0 Then a = a + pi
If Y < 0 And X < 0 Then a = a + pi
If Y < 0 And X > 0 Then a = a + 2 * pi
arctan = a
End Function





Private Sub Command3_Click()
MsgBox ("在Mandelbrot集合里点一下，可计算其对应点的Julia集合。")
MsgBox ("如需放大某一点，点它再改放大倍数即可。")
End Sub

Public Sub Command4_Click()
On Error Resume Next

SavePicture Picture2.Image, Text10.Text & "\Mandelbrot" & Text9.Text & ".bmp"
Text9.Text = Text9.Text + 1
End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Public Sub Command6_Click()
On Error Resume Next
SavePicture Picture1.Image, Text10.Text & "\Julia" & Text8.Text & ".bmp"

Text8.Text = Text8.Text + 1
End Sub

Private Sub Command7_Click()
re = Shell(App.Path & "\" & App.EXEName & ".exe")
End
End Sub

Private Sub Form_Load()
Form1.Show
Picture1.Scale (-2, 2)-(2, -2)
Picture2.Scale (-2, 2)-(2, -2)
v1 = 0
h1 = 0
v2 = 0
h2 = 0
Text10.Text = App.Path

End Sub



Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
h1 = X * 1000
v1 = Y * 1000

End Sub

Private Sub picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = X
Text3.Text = Y
h2 = X * 1000
v2 = Y * 1000
End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.Caption = X
Label12.Caption = Y
Timer1.Enabled = True
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Check1.Value = Checked Then
Dim se As Double
Dim xe As Double
Dim sc As Double
Dim xc As Double
Dim m As Double
Dim d As Double
Dim t As Double
Dim X As Double, Y As Double, a As Double, movex As Double, movey As Double, bs As Double, tmax As Double
Dim arg, pow As Double
pow = Val(Text11.Text)
d = 0.01
tmax = Val(Text6.Text)
sc = Label11.Caption
xc = Label12.Caption
bs = Val(Text7.Text)
If bs <= 0 Then

Else
Picture5.Cls
movex = h1 / 1000
movey = v1 / 1000
Picture5.Scale (movex - 2 / bs, movey + 2 / bs)-(movex + 2 / bs, movey - 2 / bs)
For se = (movex - (2 / bs)) To (movex + (2 / bs)) Step 0.13 / bs
    For xe = (movey - (2 / bs)) To (movey + (2 / bs)) Step 0.13 / bs
        m = 0
        t = 0
        X = se: Y = xe
        While m < 10 And t < tmax
            m = Sqr(X * X + Y * Y)
            arg = 0
            If pow = 2 Then
                a = X
                X = X * X - Y * Y + sc
                Y = 2 * a * Y + xc
                t = t + 1
            Else
                arg = arctan(X, Y)
                arg = arg * pow
                
                m = m ^ pow
                X = m * Cos(arg) + sc
                Y = m * Sin(arg) + xc
                t = t + 1
            End If
            
        Wend
       
           Picture5.PSet (se, xe), RGB(255 - t / tmax * 255, Abs(255 - 2 * Abs(127 - t / tmax * 255)), t / tmax * 255)
    
       
        DoEvents
    Next xe
Next se

End If
End If
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Check2.Value = Checked Then
Dim se As Double
Dim xe As Double
Dim sc As Double
Dim xc As Double
Dim m As Double
Dim d As Double
Dim t As Double
Dim X As Double, Y As Double, a As Double, movex As Double, movey As Double, bs As Double, tmax As Double, x1 As Double, y1 As Double
Dim arg, pow As Double
pow = Val(Text11.Text)
d = 0.01
tmax = Val(Text6.Text)
movex = h2 / 1000
movey = v2 / 1000
bs = Val(Text5.Text)
x1 = Label11.Caption
y1 = Label12.Caption
If bs <= 0 Then

Else
Picture6.Cls
Picture6.Scale (movex - 2 / bs, movey + 2 / bs)-(movex + 2 / bs, movey - 2 / bs)
For sc = (movex - (2 / bs)) To (movex + (2 / bs)) Step 0.13 / bs
    For xc = (movey - (2 / bs)) To (movey + (2 / bs)) Step 0.13 / bs
        m = 0
        t = 0
        X = x1: Y = y1
        While m < 10 And t < tmax
            m = Sqr(X * X + Y * Y)
            arg = 0
            If pow = 2 Then
                a = X
                X = X * X - Y * Y + sc
                Y = 2 * a * Y + xc
                t = t + 1
            Else
                arg = arctan(X, Y)
                
                arg = arg * pow
                
                m = m ^ pow
                X = m * Cos(arg) + sc
                Y = m * Sin(arg) + xc
                t = t + 1
            End If
            
        Wend
      
            Picture6.PSet (sc, xc), RGB(255 - t / tmax * 255, Abs(255 - 2 * Abs(127 - t / tmax * 255)), t / tmax * 255)
        
DoEvents
    Next xc
Next sc

End If
End If
Timer2.Enabled = False
End Sub
