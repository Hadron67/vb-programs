Attribute VB_Name = "Module1"
Global c1 As Double, c2 As Double, c3 As Double
Public Sub getcolor(x As Double, y As Double)
Dim sj As Long
sj = Form1.Picture1.Point(x, y)
c1 = (sj Mod 65536) Mod 256 'Red
c2 = (sj Mod 65536) \ 256 'Green
c3 = sj \ 65536 'Blue
End Sub
