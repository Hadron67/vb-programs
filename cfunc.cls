VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function plus(z1 As Complex, z2 As Complex) As Complex
Dim a As Double, b As Double, c As New Complex
a = z1.cpx(1) + z2.cpx(1)
b = z1.cpx(2) + z2.cpx(2)
Call c.cget(a, b)
Set plus = c
End Function
Public Function minu(z1 As Complex, z2 As Complex) As Complex
Dim a As Double, b As Double, c As New Complex
a = z1.cpx(1) - z2.cpx(1)
b = z1.cpx(2) - z2.cpx(2)
Call c.cget(a, b)
Set minu = c
End Function
Public Function multi(z1 As Complex, z2 As Complex) As Complex
Dim a As Double, b As Double, c As New Complex
a = z1.cpx(1) * z2.cpx(1) - z1.cpx(2) * z2.cpx(2)
b = z1.cpx(1) * z2.cpx(2) + z1.cpx(2) * z2.cpx(1)
Call c.cget(a, b)
Set multi = c
End Function
Public Function devi(z1 As Complex, z2 As Complex) As Complex
Dim a As Double, b As Double, c As New Complex, d As Double
a = z2.cpx(1) * z2.cpx(1) - z2.cpx(2) * z2.cpx(2)
b = (z1.cpx(1) * z2.cpx(1) + z1.cpx(2) * z2.cpx(2)) / a
d = (z1.cpx(2) * z2.cpx(1) - z1.cpx(1) * z2.cpx(2)) / a
Call c.cget(b, d)
Set devi = c
End Function
Public Function pow(z1 As Complex, a As Double) As Complex
Dim arc As Double, arg As Double, arc1 As Double, arg1 As Double, d As Double, b As Double, c As New Complex
arg = Sqr(z1.cpx(1) * z1.cpx(1) + z1.cpx(2) * z1.cpx(2))
arc = arctan(z1.cpx(2), z1.cpx(1))
arg1 = arg ^ a
arc1 = arc * a
d = arg1 * Cos(arc1)
b = arg1 * Sin(arc1)
Call c.cget(d, b)
Set pow = c
End Function
Private Function arctan(X As Double, Y As Double) As Double
Dim pi As Double, a As Double
pi = 3.14159265358979
a = Atn(Y / X)
If Y > 0 And X < 0 Then a = a + pi
If Y < 0 And X < 0 Then a = a + pi
If Y < 0 And X > 0 Then a = a + 2 * pi
arctan = a
End Function


