Attribute VB_Name = "PossiUtil"
Option Explicit

' 记录一个格子里可能出现的数字（已经确定的数字这里全false）
Public Type Possibility
    p(8) As Boolean
End Type

Public Sub PossibilitySetAll(ByRef Possibility As Possibility, state As Boolean)
    Dim i%
    For i = 0 To 8
        Possibility.p(i) = state
    Next i
End Sub

