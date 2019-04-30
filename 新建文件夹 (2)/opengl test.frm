VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10500
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   7320
      Max             =   3140
      TabIndex        =   7
      Top             =   840
      Value           =   1570
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   7320
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   7320
      Max             =   1000
      TabIndex        =   5
      Top             =   480
      Value           =   100
      Width           =   3135
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   7320
      Max             =   1000
      TabIndex        =   4
      Top             =   120
      Value           =   100
      Width           =   3135
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   3015
      Left            =   9600
      Max             =   628
      TabIndex        =   3
      Top             =   3960
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   3015
      Left            =   8520
      Max             =   628
      TabIndex        =   2
      Top             =   3960
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      Left            =   7320
      Max             =   628
      TabIndex        =   1
      Top             =   3960
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9960
      Top             =   2160
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
      Caption         =   "Label1"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   4320
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
ax = 1
ay = 0
bx = 0
by = 1
Call g1(1, 0, 0, 1)
Text1.Text = g1x
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
Call Display
Picture2.Scale (-5, 5)-(5, -5)
n = 0
 Call seifert
End Sub




Private Sub Timer1_Timer()
Call Display
End Sub
