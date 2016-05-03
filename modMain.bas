Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crCloor As Long) As Long

Public gbolNeedOutput As Boolean

Sub main()
    gbolNeedOutput = False
    frmOcr.Show

End Sub

