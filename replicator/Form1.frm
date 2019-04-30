VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   4725
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


End Sub

Private Sub Form_Load()
Form1.Hide
Dim n As Integer, b As String, c As String
While Dir(App.Path & "\replicator" & n)
n = n + 1
Wend
MkDir (App.Path & "\replicator" & n)
Dim a

a = Shell(cmd.exe)

End Sub
