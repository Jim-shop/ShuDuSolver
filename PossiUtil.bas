Attribute VB_Name = "PossiUtil"
Option Explicit

' ��¼һ����������ܳ��ֵ����֣��Ѿ�ȷ������������ȫfalse��
Public Type Possibility
    p(8) As Boolean
End Type

Public Sub PossibilitySetAll(ByRef Possibility As Possibility, state As Boolean)
    Dim i%
    For i = 0 To 8
        Possibility.p(i) = state
    Next i
End Sub

