VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOCRDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim iBlankCount As Integer
Dim iPixelCount As Integer

Dim iChar As String

Dim iContent As String

Public Property Get BlankCount() As Integer
    BlankCount = iBlankCount
End Property

Public Property Let BlankCount(ByVal BlankCount As Integer)
    iBlankCount = BlankCount
End Property

Public Property Get PixelCount() As Integer
    PixelCount = iPixelCount
End Property

Public Property Let PixelCount(ByVal PixelCount As Integer)
    iPixelCount = PixelCount
End Property

Public Property Get Word() As String
    Word = iChar
End Property

Public Property Let Word(ByVal Word As String)
    iChar = Word
End Property

Public Property Get Content() As String
    Content = iContent
End Property

Public Property Let Content(ByVal Content As String)
    iContent = Content
End Property

