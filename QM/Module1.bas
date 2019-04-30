Attribute VB_Name = "Module1"
Global fia(1000) As Double, fib(1000) As Double
Sub evaluate(dt As Double, dx As Double, h As Double)
On Error Resume Next
Dim t As Integer
Dim v1 As Double, v2 As Double, vol As Double
For t = 0 To 1000 Step 1
    vol = 0
    v1 = -h * (fib(t + 1) + fib(t - 1) - 2 * fib(t)) / dx * dx + vol / h
    v2 = h * (fia(t + 1) + fia(t - 1) - 2 * fia(t)) / dx * dx - vol / h
    fia(t) = fia(t) + v1 * dt
    fib(t) = fib(t) + v2 * dt
Next t
fia(0) = 0
fia(1000) = 0
 fib(0) = 0
fib(1000) = 0
Call reduce(dx)
End Sub
Sub init()
Dim t As Long
For t = 0 To 1000 Step 1
    fia(t) = Exp(-(t - 500) * (t - 500) / 100)
    fib(t) = 0
Next t
fia(0) = 0
fia(1000) = 0
End Sub
Sub reduce(dx As Double)
Dim t As Integer, s As Double
s = 0
For t = 0 To 1000 Step 1
s = s + (fia(t) * fia(t) + fib(t) * fib(t)) * 1000 * dx
Next t
For t = 0 To 1000 Step 1
    fia(t) = fia(t) / Sqr(s)
    fib(t) = fib(t) / Sqr(s)
Next t
End Sub
