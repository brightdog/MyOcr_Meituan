VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iTop As Long
Private iLeft As Long

Private iWidth As Long
Private iHight As Long

Private iContent As Variant

Public Property Get Top() As Long
    Top = iTop
End Property

Public Property Let Top(ByVal Top As Long)
    iTop = Top
End Property

Public Property Get Left() As Long
    Top = iLeft
End Property

Public Property Let Left(ByVal Left As Long)
    iLeft = Left
End Property

Public Property Get width() As Long
    width = iWidth
End Property

Public Property Let width(ByVal width As Long)
    iWidth = width
End Property

Public Property Get Hight() As Long
    Hight = iHight
End Property

Public Property Let Hight(ByVal Hight As Long)
    iHight = Hight
End Property

Public Property Get Content() As Variant
    Content = iContent
End Property

Public Property Let Content(ByVal Content As Variant)
    iContent = Content
End Property
