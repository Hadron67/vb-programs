VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Complex Number Factal"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9120
   StartUpPosition =   3  '窗口缺省
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   480
      Max             =   3000
      Min             =   -3000
      TabIndex        =   20
      Top             =   7200
      Width           =   8295
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7095
      LargeChange     =   100
      Left            =   120
      Max             =   3000
      Min             =   -3000
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   3240
      TabIndex        =   8
      Text            =   "0"
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   2160
      TabIndex        =   7
      Text            =   "0"
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mandelbrot集"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   6240
      TabIndex        =   5
      Text            =   "201"
      Top             =   8505
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   7440
      TabIndex        =   4
      Text            =   "3"
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   5760
      TabIndex        =   3
      Text            =   "0.01"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3240
      TabIndex        =   2
      Text            =   "0"
      Top             =   7665
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2160
      TabIndex        =   1
      Text            =   "0"
      Top             =   7680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Julia集"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "迭代次数="
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "h="
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
      Left            =   7080
      TabIndex        =   16
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "fe="
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
      Left            =   5280
      TabIndex        =   15
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "i"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "+"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "S="
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "i"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "+"
      Height          =   255
      Left            =   3000
      TabIndex        =   10
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "C="
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   7680
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
h = Text4.Text
f = Text6.Text
Cls
Dim se As Double
Dim xe As Double
Dim sc As Double
Dim xc As Double
Dim m As Double
Dim d As Double
Dim t As Double
Dim a As Double
Dim b As Double
Dim c As Double, x As Double, y As Double
Scale (-h, h)-(h, -h)
x = (VScroll1.Value)
y = (HScroll1.Value)
d = Text3.Text
sc = Text1.Text
xc = Text2.Text
For se = -h + x To h + x Step d
For xe = -h + y To h + y Step d
a = se
b = xe
m = Sqr(a * a + b * b)
t = 0
While m < 2000 And t < f
c = a
a = a * a - b * b + sc
b = 2 * c * b + xc
m = Sqr(a * a + b * b)
t = t + 1
Wend
If m < 1000 And t > f - 2 Then
PSet (se, xe)
End If
Next xe
Next se
End Sub


Private Sub Command2_Click()
h = Text4.Text
f = Text6.Text
Scale (-h, h)-(h, -h)
Cls
Dim se As Double
Dim xe As Double
Dim sc As Double
Dim xc As Double
Dim m As Double
Dim d As Double
Dim t As Double
Dim a As Double
Dim b As Double
Dim c As Double
d = Text3.Text

For sc = -h To h Step d
For xc = -h To h Step d
m = Sqr(a * a + b * b)
t = 0
a = Text7.Text
b = Text8.Text
While m < 2 And t < f
c = a
a = a * a - b * b + sc
b = 2 * c * b + xc
m = Sqr(a * a + b * b)
t = t + 1
Wend
If m < 10 And t > f - 2 Then
PSet (sc, xc)
End If
Next xc
Next sc
End Sub

Private Sub Command3_Click()
MsgBox ("该程序可以计算z的平方的填充Julia集和Mandelbrot集。")
End Sub

