Attribute VB_Name = "Module1"
Global wave(300, 300), vt(300, 300), a(300, 300) As Double, t As Double


Sub react(k As Double, v As Double, dt As Double)
On Error Resume Next
Dim x As Double, y As Double, ax, ay, at As Double
For x = 0 To 200 Step 1
    For y = 0 To 200 Step 1
        a(x, y) = vv(x, y) * (wave(x + 1, y + 1) + wave(x, y + 1) + wave(x + 1, y) + wave(x - 1, y - 1) + wave(x - 1, y) + wave(x, y - 1) + wave(x - 1, y + 1) + wave(x + 1, y - 1) - 8 * wave(x, y))
        vt(x, y) = vt(x, y) + a(x, y) * dt
        DoEvents
        Next y
Next x
For x = 0 To 200 Step 1
    For y = 0 To 200 Step 1
        wave(x, y) = wave(x, y) + vt(x, y) * dt + 0.5 * a(x, y) * dt * dt
        DoEvents
    Next y
Next x

End Sub
Private Function vv(x As Double, y As Double)
If x + y <= 200 Then
    vv = 1
ElseIf x + y > 200 Then
    vv = 0.5
End If
End Function
