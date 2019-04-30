VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6465
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   3000
      ScaleHeight     =   2595
      ScaleWidth      =   2235
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim z1 As New Complex, z2 As New Complex, x As Double, y As Double, a As New cfun, c As New Complex
For x = -25 To 25 Step 0.1
For y = -25 To 25 Step 0.1
Call c.cget(5, 0)
Call z1.cget(x, y)
Set z2 = a.pow(z1, 1)
Picture2.PSet (z1.cpx(1), z1.cpx(2)), Picture1.Point(z2.cpx(1), z2.cpx(2))
DoEvents
Next y
Next x

End Sub

Private Sub Form_Load()
Picture1.Scale (0, 5)-(5, 0)
Picture2.Scale (-20, 20)-(20, -20)
End Sub
