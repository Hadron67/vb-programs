VERSION 5.00
Begin VB.Form Hutton32 
   Caption         =   "Hutton32"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Hutton32"
   ScaleHeight     =   3960
   ScaleWidth      =   5115
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Nobili"
      Height          =   855
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   2520
      Picture         =   "golly.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   2475
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清除数据"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "输出"
      Height          =   855
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "逆翻译"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Hutton32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub puts()
Dim a As String, x As Integer, t As Integer
a = Text1.Text
x = 1: t = 1
Do
    
    N(x) = Right(Left(a, t), 1)
    If N(x) = "p" Then x = x + 1: t = t + 2 Else x = x + 1: t = t + 1
    If N(x - 1) <> "I" And N(x - 1) <> "J" And N(x - 1) <> "K" And N(x - 1) <> "L" And N(x - 1) <> "Q" And N(x - 1) <> "R" And N(x - 1) <> "S" And N(x - 1) <> "T" And N(x - 1) <> "1" And N(x - 1) <> "2" And N(x - 1) <> "3" And N(x - 1) <> "4" And N(x - 1) <> "5" And N(x - 1) <> "6" And N(x - 1) <> "7" And N(x - 1) <> "8" And N(x - 1) <> "9" And N(x - 1) <> "0" And N(x - 1) <> "$" And N(x - 1) <> "!" And N(x - 1) <> "p" And N(x - 1) <> "." Then x = x - 1
Loop Until N(x - 1) = "!"
End Sub

Private Sub Command1_Click()
On Error Resume Next
Text2.Text = ""
If Text1.Text <> "" Then
Call puts
Dim x As Integer, t As Integer, ti As Integer, s As Integer, t1 As Integer, sign As Integer, m As Integer
x = 0
While N(x) <> "!"
    x = x + 1
   
    For t = 1 To 1000 Step 1
        d1(t) = "": d2(t) = 0
    Next t
    m = 0: s = 0
    While N(x) <> "$" And N(x) <> "!"
        If N(x) = "1" Or N(x) = "2" Or N(x) = "3" Or N(x) = "4" Or N(x) = "5" Or N(x) = "6" Or N(x) = "7" Or N(x) = "8" Or N(x) = "9" Or N(x) = "0" Then
            ti = 1: m = m + 1
            While N(x + ti) = "1" Or N(x + ti) = "2" Or N(x + ti) = "3" Or N(x + ti) = "4" Or N(x + ti) = "5" Or N(x + ti) = "6" Or N(x + ti) = "7" Or N(x + ti) = "8" Or N(x + ti) = "9" Or N(x + ti) = "0"
                ti = ti + 1
            Wend
            For t = 1 To ti Step 1
                d2(m) = d2(m) + Val(N(x + t - 1)) * 10 ^ (ti - t)
            Next t
            x = x + ti + 1
            d1(m) = N(x - 1)
        Else
        If N(x) <> "" Then
            m = m + 1
            d1(m) = N(x)
            d2(m) = 1
            x = x + 1
            End If
        End If
    Wend
    For t = 1 To m + 1 Step 1
        
        s = s + d2(t)
    Next t
    Call down(1)
    Call movefor(s - 1)
    For t = m + 1 To 1 Step -1
        If d1(t) = "I" Then Call bar1(d2(t))
        If d1(t) = "J" Then Call bau1(d2(t))
        If d1(t) = "K" Then Call bal1(d2(t))
        If d1(t) = "L" Then Call bad1(d2(t))
        If d1(t) = "Q" Then Call rar1(d2(t))
        If d1(t) = "R" Then Call rau1(d2(t))
        If d1(t) = "S" Then Call ral1(d2(t))
        If d1(t) = "T" Then Call rad1(d2(t))
        If d1(t) = "p" Then Call pa1(d2(t))
        If d1(t) = "." Then Call back(d2(t))
    Next t
    Call back(1)
Wend
End If

        


End Sub

Private Sub Command2_Click()
Text2.Text = "x = 0, y = 0, rule = Hutton32" & vbCrLf
For t = 1 To 10000 Step 1
    Text2.Text = Text2.Text & out(t)
    DoEvents
Next t
MsgBox ("Finished!")
End Sub
Private Sub bau1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "O3K2O"
tn = tn + 1
Next tt
End Sub
Private Sub bal1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "O2K2OK"
tn = tn + 1
Next tt
End Sub
Private Sub bad1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "OKOKOK"
tn = tn + 1
Next tt
End Sub
Private Sub bar1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "O4KO"
tn = tn + 1
Next tt
End Sub
Private Sub rau1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "2O2KOK"
tn = tn + 1
Next tt
End Sub
Private Sub rar1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "OK2OKO"
tn = tn + 1
Next tt
End Sub
Private Sub rad1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "3OKOK"
tn = tn + 1
Next tt
End Sub
Private Sub ral1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "2OKOKO"
tn = tn + 1
Next tt
End Sub
Private Sub pa1(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "4OK"
tn = tn + 1
Next tt
End Sub

Private Sub movefor(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "O5K"
tn = tn + 1
Next tt
End Sub
Private Sub back(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "OK4O"
tn = tn + 1
Next tt
End Sub
Private Sub down(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "OKO2K"
tn = tn + 1
Next tt
End Sub

Private Sub Command3_Click()
For t = 0 To 10000 Step 1
N(t) = "": out(t) = ""
Next t
For t = 1 To 100 Step 1
d1(t) = "": d2(t) = 0
Text1.Text = "": Text2.Text = ""
tn = 1
Next t
End Sub

Private Sub Command4_Click()
On Error Resume Next
Text2.Text = ""
If Text1.Text <> "" Then
Call puts
Dim x As Integer, t As Integer, ti As Integer, s As Integer, t1 As Integer, sign As Integer, m As Integer
x = 0
While N(x) <> "!"
    x = x + 1
   
    For t = 1 To 1000 Step 1
        d1(t) = "": d2(t) = 0
    Next t
    m = 0: s = 0
    While N(x) <> "$" And N(x) <> "!"
        If N(x) = "1" Or N(x) = "2" Or N(x) = "3" Or N(x) = "4" Or N(x) = "5" Or N(x) = "6" Or N(x) = "7" Or N(x) = "8" Or N(x) = "9" Or N(x) = "0" Then
            ti = 1: m = m + 1
            While N(x + ti) = "1" Or N(x + ti) = "2" Or N(x + ti) = "3" Or N(x + ti) = "4" Or N(x + ti) = "5" Or N(x + ti) = "6" Or N(x + ti) = "7" Or N(x + ti) = "8" Or N(x + ti) = "9" Or N(x + ti) = "0"
                ti = ti + 1
            Wend
            For t = 1 To ti Step 1
                d2(m) = d2(m) + Val(N(x + t - 1)) * 10 ^ (ti - t)
            Next t
            x = x + ti + 1
            d1(m) = N(x - 1)
        Else
        If N(x) <> "" Then
            m = m + 1
            d1(m) = N(x)
            d2(m) = 1
            x = x + 1
            End If
        End If
    Wend
    For t = 1 To m + 1 Step 1
        
        s = s + d2(t)
    Next t
    
    Call movefor2(s)
    For t = m + 1 To 1 Step -1
        If d1(t) = "I" Then Call bar2(d2(t))
        If d1(t) = "J" Then Call bau2(d2(t))
        If d1(t) = "K" Then Call bal2(d2(t))
        If d1(t) = "L" Then Call bad2(d2(t))
        If d1(t) = "Q" Then Call rar2(d2(t))
        If d1(t) = "R" Then Call rau2(d2(t))
        If d1(t) = "S" Then Call ral2(d2(t))
        If d1(t) = "T" Then Call rad2(d2(t))
        If d1(t) = "p" Then Call pa2(d2(t))
        If d1(t) = "." Then Call back2(d2(t))
    Next t
    Call down2(1)
Wend
End If

End Sub

Private Sub Form_Load()
tn = 1
End Sub
Private Sub bau2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "2.2L"
tn = tn + 1
Next tt
End Sub
Private Sub bal2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = ".L2."
tn = tn + 1
Next tt
End Sub
Private Sub bad2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "3L."
tn = tn + 1
Next tt
End Sub
Private Sub bar2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "2.L."
tn = tn + 1
Next tt
End Sub
Private Sub rau2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "L2.L"
tn = tn + 1
Next tt
End Sub
Private Sub rar2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = ".3L"
tn = tn + 1
Next tt
End Sub
Private Sub rad2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "2L2."
tn = tn + 1
Next tt
End Sub
Private Sub ral2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "L.2L"
tn = tn + 1
Next tt
End Sub
Private Sub pa2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "2L.L"
tn = tn + 1
Next tt
End Sub

Private Sub movefor2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = ".2L."
tn = tn + 1
Next tt
End Sub
Private Sub back2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = ".L.L"
tn = tn + 1
Next tt
End Sub
Private Sub down2(t As Integer)
Dim tt As Integer
For tt = 1 To t Step 1
out(tn) = "4L"
tn = tn + 1
Next tt
End Sub


