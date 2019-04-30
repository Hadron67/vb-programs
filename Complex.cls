VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Complex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private x As Double
Private y As Double
Public Sub cget(M As Double, n As Double)
x = M: y = n
End Sub
Public Function cpx(t As Integer) As Double
If t = 1 Then cpx = x
If t = 2 Then cpx = y
End Function




