VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   5430
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2640
      TabIndex        =   8
      Text            =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1560
      TabIndex        =   7
      Text            =   "1"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   6
      Text            =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   120
      Max             =   10
      Min             =   -10
      TabIndex        =   5
      Top             =   5040
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   120
      Max             =   10
      Min             =   -10
      TabIndex        =   4
      Top             =   4680
      Width           =   3495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   10
      Min             =   -10
      TabIndex        =   3
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Scale (-150, 250)-(150, -250)
x = Text1.Text
y = Text2.Text
z = Text3.Text

For t = 0 To 10 Step 0.001
vx = -10 * (x - y)
vy = -x * z + 28 * x - y * 2
vz = x * y - (8 / 3) * z
x = x + vx / 100
y = y + vy / 100
z = z + vz / 100
Call lineto((x) * 6, (z - 20) * 6)
Next t
End Sub
Private Sub lineto(x As Double, y As Double)
Line (Label1.Caption, Label2.Caption)-(x, y)
Label1.Caption = x
Label2.Caption = y
End Sub

Private Sub Timer1_Timer()
Scale (-150, 250)-(150, -250)
x = Text1.Text
y = Text2.Text
z = Text3.Text
a = HScroll1.Value
b = HScroll2.Value
c = HScroll3.Value
For t = 0 To 10 Step 0.001
vx = a * (x - y)
vy = x * (b - z) - y
vz = x * y - c * z
x = x + vx / 100
y = y + vy / 100
z = z + vz / 100
Call lineto((x) * 6, (z - 20) * 6)
Next t
End Sub
