Attribute VB_Name = "PossiUtil"
Option Explicit

' ��¼һ����������ܳ��ֵ����֣��Ѿ�ȷ������������ȫfalse��
Public Type possibility
    p(8) As Boolean
End Type

Public Sub PossibilitySetAll(ByRef possibility As possibility, state As Boolean)
    Dim i%
    For i = 0 To 8
        possibility.p(i) = state
    Next i
End Sub

