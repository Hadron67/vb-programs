VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   5400
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Text            =   "1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3165
      Left            =   120
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   3165
      Left            =   120
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3045
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Load()
Dim x As Double
t = 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim x As Double, tt As Double

For x = 10 To 120 Step 1
wave(x, 30) = 200 * Sin(t)
Next x
t = t + 0.1

    react 0.5, 1, 0.1
    DoEvents


For x = 0 To 200 Step 1
For y = 0 To 200 Step 1
Picture2.PSet (x, y), RGB(128, 128 + 128 * wave(x, y) / 200, 128)

DoEvents
Next y
Next x
Picture1.Picture = Picture2.Image
SavePicture Picture2.Image, "F:\f\flash\vb Programs\wave\" & "wave" & Text1.Text & ".bmp"
Text1.Text = Text1.Text + 1
End Sub
