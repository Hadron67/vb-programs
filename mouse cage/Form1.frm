VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "mouse cage"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   5625
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "NaN"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 71 Then Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Form1.Hide
MsgBox "Warning! this program will control your mouse"
Form1.Show
Timer1.Enabled = True
Timer2.Enabled = True
t = 500
End Sub

Private Sub Timer1_Timer()
Picture1.Drag
Form1.Show
End Sub

Private Sub Timer2_Timer()
Label1.Caption = "You have to wait for " & t & " seconds"
t = t - 1
If t <= 0 Then Timer1.Enabled = False
End Sub
