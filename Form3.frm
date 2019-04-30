VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   6105
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      LargeChange     =   100
      Left            =   240
      Max             =   1000
      Min             =   -1000
      TabIndex        =   6
      Top             =   5040
      Width           =   5535
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   100
      Left            =   240
      Max             =   4000
      Min             =   -4000
      TabIndex        =   5
      Top             =   4560
      Width           =   5535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   5520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Double, r As Double, t As Double
Scale (0, 0)-(4, 3)
For r = 0 To 4 Step 0.01
    a = 0.3333
    For t = 0 To 200 Step 1
        a = a * r * (1 - a)
    Next t
    For t = 0 To 100 Step 1
        a = a * r * (1 - a)
        PSet (r, a)
    Next t
Next r
End Sub

Private Sub Command2_Click()
Label3.Caption = Label3.Caption + 1

End Sub

Private Sub lineto(x As Double, y As Double)
Line (Label1.Caption, Label2.Caption)-(x, y)
Label1.Caption = x
Label2.Caption = y
End Sub
Private Sub moveto(x As Double, y As Double)
Label1.Caption = x
Label2.Caption = y
End Sub
Private Function ram(a As Double, b As Double)
Randomize
ram = Int((b - a + 1) * Rnd + a)
End Function

Private Sub Timer1_Timer()
On Error Resume Next
Scale (0, -1)-(50, 2)
Dim a As Double, t As Double, r As Double
r = HScroll1.Value / 1000
a = HScroll2.Value / 1000
If Label3.Caption Mod 2 = 0 Then
   Cls
   For t = 0 To 50 Step 1
   a = r * a * (1 - a)
   Call lineto(t, a)
   Next t
End If
End Sub
