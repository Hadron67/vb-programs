VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Transformer"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   14055
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command5 
      Caption         =   "Clear Grid"
      Height          =   495
      Left            =   1440
      TabIndex        =   26
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Test(Z->Z)"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Savepicture"
      Height          =   495
      Left            =   1440
      TabIndex        =   24
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Text            =   "Text10"
      Top             =   1320
      Width           =   5055
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Z->ln(Z)"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paint Grid"
      Height          =   495
      Left            =   1440
      TabIndex        =   21
      Top             =   3000
      Width           =   1335
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Z->exp(Z)"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Z->1/Z"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Z->Z*Z"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Z->sqrt(Z)"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5280
      TabIndex        =   16
      Top             =   9120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   3720
      TabIndex        =   14
      Text            =   "-2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   3000
      TabIndex        =   13
      Text            =   "2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1680
      TabIndex        =   12
      Text            =   "2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   960
      TabIndex        =   11
      Text            =   "-2"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   3720
      TabIndex        =   9
      Text            =   "-5"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   3000
      TabIndex        =   8
      Text            =   "5"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1680
      TabIndex        =   7
      Text            =   "5"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Text            =   "-5"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   840
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Loadpicture"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   8895
      Left            =   5280
      ScaleHeight     =   589
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   573
      TabIndex        =   2
      Top             =   120
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   4440
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   8895
      Left            =   5280
      ScaleHeight     =   589
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   573
      TabIndex        =   15
      Top             =   120
      Width           =   8655
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   3000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Caption         =   "scale2:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   3120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Caption         =   "scale1:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo err
ProgressBar1.Value = 0
Command1.Enabled = False
Picture2.Visible = False
Picture2.Cls
Picture1.Scale (Text2.Text, Text3.Text)-(Text4.Text, Text5.Text)
'Picture2.Scale (-5, 5)-(5, -5)
Dim x0 As Double, y0 As Double, x As Double, y As Double, x1 As Double, y1 As Double
For x0 = 0 To Picture2.ScaleWidth Step 1
    For y0 = 0 To Picture2.ScaleHeight Step 1
        If x0 <> 0 And y0 <> 0 Then
        x1 = (x0 / Picture2.ScaleWidth) * (Text8.Text - Text6.Text) + Text6.Text
        y1 = (y0 / Picture2.ScaleHeight) * (Text9.Text - Text7.Text) + Text7.Text
        x = tran(x1, y1, 1)
        y = tran(x1, y1, 2)
        Picture2.PSet (x0, y0), Picture1.Point(x, y)
        DoEvents
        End If
        ProgressBar1.Value = 100 * (x0 / Picture2.ScaleWidth)
    Next y0
Next x0
Command1.Enabled = True
ProgressBar1.Value = 0
Picture2.Visible = True
Picture1.Cls
Exit Sub
err:
    MsgBox ("Êä´íÁË£¡")
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
On Error GoTo err
Dim a As Picture
Set a = LoadPicture(Text1.Text)
Picture1.Height = a.Height * (3255 / 5715)
Picture1.Width = a.Width * (2655 / 4657)
Picture1.Picture = a
MsgBox ("Loading Successful")
Exit Sub
err:
    MsgBox ("cannot load picture")
End Sub

Private Sub Command3_Click()
Picture1.ScaleMode = 3 - pixel
Picture1.Line (0, 0)-(0, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth / 8, 0)-(Picture1.ScaleWidth / 8, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth / 4, 0)-(Picture1.ScaleWidth / 4, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth * 3 / 8, 0)-(Picture1.ScaleWidth * 3 / 8, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth / 2, 0)-(Picture1.ScaleWidth / 2, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth * 3 / 4, 0)-(Picture1.ScaleWidth * 3 / 4, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth * 5 / 8, 0)-(Picture1.ScaleWidth * 5 / 8, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth * 7 / 8, 0)-(Picture1.ScaleWidth * 7 / 8, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (Picture1.ScaleWidth, 0)-(Picture1.ScaleWidth, Picture1.ScaleHeight), RGB(0, 255, 0)
Picture1.Line (0, 0)-(Picture1.ScaleWidth, 0), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight / 8)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 8), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight / 4)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 4), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight * 3 / 8)-(Picture1.ScaleWidth, Picture1.ScaleHeight * 3 / 8), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight / 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight * 5 / 8)-(Picture1.ScaleWidth, Picture1.ScaleHeight * 5 / 8), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight * 3 / 4)-(Picture1.ScaleWidth, Picture1.ScaleHeight * 3 / 4), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight * 7 / 8)-(Picture1.ScaleWidth, Picture1.ScaleHeight * 7 / 8), RGB(0, 0, 255)
Picture1.Line (0, Picture1.ScaleHeight)-(Picture1.ScaleWidth, Picture1.ScaleHeight), RGB(0, 0, 255)
MsgBox ("Paiting Complete")
End Sub

Private Sub Command4_Click()
On Error GoTo err
SavePicture Picture2.Image, Text10.Text
Exit Sub
err:
MsgBox ("cannot save picture")
End Sub

Private Sub Command5_Click()
    Picture1.Cls
    MsgBox ("Grid Cleared")
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Text1.Text = App.Path & "\test.jpg"
Text10.Text = App.Path & "\result.bmp"

End Sub
Private Function tran(x As Double, y As Double, co As Integer)
On Error Resume Next
Dim e As Double
e = 2.718281828
If co = 1 Then
    If Option1.Value Then
        tran = x * x - y * y
    ElseIf Option2.Value Then
        tran = Sqr((x + Sqr(x * x + y * y)) / 2)
    ElseIf Option3.Value Then
        tran = x / (x * x + y * y)
    ElseIf Option4.Value Then
        tran = Log(x * x + y * y) / 2
    ElseIf Option5.Value Then
        tran = Exp(x) * Cos(y)
    ElseIf Option6.Value Then
        tran = x
    End If
ElseIf co = 2 Then
    If Option1.Value Then
        tran = 2 * x * y
    ElseIf Option2.Value Then
        tran = (y / 2) / (Sqr((x + Sqr(x * x + y * y)) / 2))
    ElseIf Option3.Value Then
        tran = -y / (x * x + y * y)
    ElseIf Option4.Value Then
        tran = arctan(x, y)
    ElseIf Option5.Value Then
        tran = Exp(x) * Sin(y)
    ElseIf Option6.Value Then
        tran = y
    End If
End If




End Function
Private Function arctan(x As Double, y As Double) As Double
Dim pi As Double, a As Double
pi = 3.14159265358979
a = Atn(y / x)
If y > 0 And x < 0 Then a = a + pi
If y < 0 And x < 0 Then a = a + pi
If y < 0 And x > 0 Then a = a + 2 * pi
arctan = a
End Function

