Attribute VB_Name = "Module2"
Global wave(100, 100), vt(100, 100), a(100, 100) As Double, t As Double
Sub react(k As Double, v As Double, dt As Double)
On Error Resume Next
Dim x, y, ax, ay, at As Double
For x = 0 To 100 Step 1
    For y = 0 To 100 Step 1
        a(x, y) = v * (wave(x + 1, y + 1) + wave(x, y + 1) + wave(x + 1, y) + wave(x - 1, y - 1) + wave(x - 1, y) + wave(x, y - 1) + wave(x - 1, y + 1) + wave(x + 1, y - 1) - 8 * wave(x, y))
        vt(x, y) = vt(x, y) + a(x, y) * dt
        DoEvents
        Next y
Next x
For x = 0 To 100 Step 1
    For y = 0 To 100 Step 1
        wave(x, y) = wave(x, y) + vt(x, y) * dt + 0.5 * a(x, y) * dt * dt
        DoEvents
    Next y
Next x

End Sub
Sub disball()
On Error GoTo er
Dim x, y As Double
For x = 0 To 100 Step 1
    For y = 0 To 100 Step 1
    glTranslatef x - 50, y - 50, wave(x, y)
    glutSolidSphere 1, 5, 5
    glTranslatef -x + 50, -y + 50, -wave(x, y)
    DoEvents
    Next y
Next x
Exit Sub
er:
    MsgBox "Overflow!", vbCritical, "Error"
    Form1.Text9.Text = 1
    initb
End Sub
Sub initb()
Dim x, y As Double
For x = 0 To 100 Step 1
    For y = 0 To 100 Step 1
        wave(x, y) = 0
        vt(x, y) = 0
        a(x, y) = 0
        Next y
Next x
End Sub

