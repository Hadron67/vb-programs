VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11220
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7275
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Picture1.Scale (-2, 10)-(12, 0)

End Sub

Private Sub Picture1_Click()
Dim t As Double
For t = 1 To 10 Step 0.1
    
    
        Picture1.Line (10 * Log(t) / Log(10), 5)-(10 * Log(t) / Log(10), 5.5)

    
Next t


For t = 1 To 10 Step 1
    
    
        Picture1.Line (10 * Log(t) / Log(10), 5)-(10 * Log(t) / Log(10), 6)

    
Next t
End Sub
