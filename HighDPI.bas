Attribute VB_Name = "HighDPI"
Option Explicit

Private Declare Function SetProcessDpiAwareness Lib "shcore.dll" (ByVal DPImode As Long) As Long

Private Declare Function GetScaleFactorForDevice Lib "shcore.dll" (ByVal DeviceType As Long) As Long

Public Sub EnableHighDPI(ByRef frm As Form)
    SetProcessDpiAwareness 2&
End Sub
