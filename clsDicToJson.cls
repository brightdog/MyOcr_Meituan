VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDicToJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DicToJson(ByRef dic As Scripting.Dictionary) As String

    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    SB.value = ""
    
    Dim v As Variant
    
    For Each v In dic.keys

        If Not IsEmpty(v) Then
            SB.Append """" & CStr(v) & """:"
            Dim ValType As String
        
            ValType = VBA.TypeName(dic.Item(v))

            Select Case ValType
            
                Case "Collection"
               
                    SB.Append "[" & ParseCollection(dic.Item(v)) & "],"
                
                Case "Dictionary"
            
                    SB.Append "{" & ParseDictionary(dic.Item(v)) & "},"
            
                Case Else
                    Call SetOtherCaseValue(SB, dic.Item(v))
            End Select
    
        End If

    Next

    SB.value = "{" & CutLastComma(SB.ToString) & "}"
    DicToJson = SB.ToString
End Function

Private Function ParseCollection(ByRef Col As VBA.Collection) As String

    Dim i As Integer
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    SB.value = ""
    
    For i = 1 To Col.Count
    
        Dim ValType As String
        
        ValType = VBA.TypeName(Col.Item(i))
    
        Select Case ValType
        
            Case "Collection"
                SB.Append "[" & ParseCollection(Col.Item(i)) & "],"

            Case "Dictionary"
            
                SB.Append "{" & ParseDictionary(Col.Item(i)) & "},"

            Case Else

                Call SetOtherCaseValue(SB, Col.Item(i))

        End Select
    
    Next

    ParseCollection = CutLastComma(SB.ToString)
    
    Set SB = Nothing
End Function

Private Function ParseDictionary(ByRef dic As Scripting.Dictionary) As String
    Dim v As Variant
    Dim SB As clsStringBuilder
    Set SB = New clsStringBuilder
    SB.value = ""
    
    For Each v In dic.keys
        SB.Append """" & v & """:"
        Dim ValType As String
        
        ValType = VBA.TypeName(dic.Item(v))
    
        Select Case ValType
        
            Case "Collection"
                SB.Append "[" & ParseCollection(dic.Item(v)) & "],"

            Case "Dictionary"
                SB.Append "{" & ParseDictionary(dic.Item(v)) & "},"

            Case Else
                Call SetOtherCaseValue(SB, dic.Item(v))

        End Select
    
    Next

    ParseDictionary = CutLastComma(SB.ToString)
    
    Set SB = Nothing
End Function

Private Function CutLastComma(ByVal s As String) As String

    If s <> "" Then
        Dim strResult As String
        strResult = VBA.Left(s, Len(s) - 1)
        CutLastComma = strResult
    Else
        CutLastComma = ""
    End If

End Function

Private Sub SetOtherCaseValue(ByRef SB As clsStringBuilder, ByVal s As Variant)

    If TypeName(s) = "Null" Then
        SB.Append "null,"
    Else

        If Not VBA.IsNumeric(s) Or TypeName(s) = "String" Then
            SB.Append """" & s & ""","
        Else
            SB.Append s & ","
        End If
    End If

End Sub
