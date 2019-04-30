VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5925
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.VScrollBar VScroll1 
      Height          =   2655
      Left            =   5400
      Max             =   30000
      TabIndex        =   1
      Top             =   120
      Value           =   600
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5280
      Top             =   3720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Line Line8 
      X1              =   5280
      X2              =   3360
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   2280
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.Line Line6 
      X1              =   5160
      X2              =   5280
      Y1              =   3120
      Y2              =   3000
   End
   Begin VB.Line Line5 
      X1              =   5160
      X2              =   5280
      Y1              =   2880
      Y2              =   3000
   End
   Begin VB.Line Line4 
      X1              =   5280
      X2              =   5280
      Y1              =   2760
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   120
      Y1              =   1560
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   120
      Y1              =   3120
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   240
      Y1              =   3000
      Y2              =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Picture1.Scale (0, 2)-(1001, -0.1)
Picture1.Line (0, 0)-(100, 0)
i = 0
Call init
End Sub

Private Sub lineto(x As Double, y As Double)
Picture1.Line (px, py)-(x, y), RGB(255, 0, 0)
px = x
py = y
End Sub
Private Sub moveto(x As Double, y As Double)
px = x
py = y
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()

Dim t As Double
For t = 0 To 1 Step 1
    Call evaluate(0.1, 0.03, 1)
    DoEvents
Next t
Call moveto(0, 0)
Picture1.Cls
Picture1.Line (0, 0)-(1000, 0)
For t = 0 To 1000 Step 1
    Call lineto(t, (fia(t) * fia(t) + fib(t) * fib(t)) * VScroll1.Value)
Next t
Label1.Caption = 0.03 * 1000

End Sub
