VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private RunCountPerMS As Double, RunCountPerS As Currency
Private CPUv1 As Currency, CPUv2 As Currency
Private CPUv3 As Currency, CPUv4 As Currency
Private TimePass As Currency '几毫秒，后面有4位数
Private V1 As Currency, V2 As Currency

Public Sub StartTimer()

    QueryPerformanceCounter CPUv1
End Sub
Public Function ShowMS() As Currency

    QueryPerformanceCounter CPUv2
    TimePass = (CPUv2 - CPUv1) * 1000 * 1000# / RunCountPerS
    ShowMS = TimePass
End Function

Private Sub Class_Initialize()
    If QueryPerformanceFrequency(RunCountPerS) Then '得到每秒的CPU运算次数
        RunCountPerMS = RunCountPerS / 1000
    Else
        MsgBox "不支持高精度计数器!"
    End If

End Sub
