Attribute VB_Name = "modCommon"
Option Explicit

Public Function getPartOfArray(ByRef arr As Variant, Optional ByVal xStart As Long = 0, Optional ByVal xEnd As Long = 0, Optional ByVal yStart As Long = 0, Optional ByVal yEnd As Long = 0) As String()
    '从一个2维数组里，截取特定的部分，作为新的数组，返回

    Dim i, j As Long
    
    Dim uBoundx, uBoundy As Long
    uBoundx = UBound(arr, 1)
    uBoundy = UBound(arr, 2)
    
    If xEnd = 0 Then
        xEnd = uBoundx
    End If
    
    If yEnd = 0 Then
        yEnd = uBoundy
    End If
    
    Dim arrResult() As String
    ReDim arrResult(xEnd - xStart, yEnd - yStart) As String
    
    For i = 1 To uBoundx
    
        For j = 1 To uBoundy
        
            If i >= xStart And i <= xEnd And j >= yStart And j <= yEnd Then
            
                arrResult(i - xStart, j - yStart) = arr(i, j)
        
            End If

        Next
    
    Next

    getPartOfArray = arrResult
End Function

Private Function MakeSymbol(ByVal strRaw As String, ByVal intWidth As Integer)

    Dim i As Integer
    
    Dim strResult As String
    
    For i = 1 To Len(strRaw)
    
        strResult = strResult & Mid(strRaw, i, 1)

        If i Mod intWidth = 0 Then
            strResult = strResult & vbCrLf
        End If
    
    Next

    MakeSymbol = strResult
End Function

Public Sub WriteLocalDic(ByVal strWord As String, ByVal dicData As Scripting.Dictionary, ByVal dicOCRES As Scripting.Dictionary, ByVal OCRFile As String)

    '要把dicData里的内容通过strWord，合并（新增）到dicOCRES里去，然后把dicOCRES重新序列化之后，写入文件中。

    Dim v As Variant
    Dim bolFind As Boolean
    bolFind = False
    Dim objDicToJson As clsDicToJson
    Set objDicToJson = New clsDicToJson

    For Each v In dicOCRES.keys

        If v = strWord Then
            bolFind = True
            dicOCRES.Item(v).Add dicData
            Exit For
        Else
            
        End If

    Next

    If Not bolFind Then
    
        Dim col As VBA.Collection
        Set col = New VBA.Collection
        col.Add dicData
        dicOCRES.Add strWord, col
        Set col = Nothing
    
    End If
    
    Dim strResult As String
    
    strResult = objDicToJson.DicToJson(dicOCRES)
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    Fso.OpenTextFile(App.Path & "\OCRES\" & OCRFile, ForWriting, True, TristateFalse).Write (strResult)
    Set Fso = Nothing
End Sub

Public Function JoinRawData(ByRef arr As Variant) As String
    Dim x As Long
    Dim y As Long
    Dim iWidth As Long
    Dim iHeight As Long

    iWidth = UBound(arr, 1)
    iHeight = UBound(arr, 2)
    Dim strResult As String

    For y = 1 To iHeight
    
        For x = 1 To iWidth
    
            strResult = strResult & arr(x, y)
    
        Next
        
        strResult = strResult & "&"
        
    Next

    JoinRawData = strResult
End Function

Public Function parseResultToJson(ByRef strResult As String) As String
    '传地址，然后直接改了算了。
    Dim arrChar() As String
    
    arrChar = VBA.Split(strResult, vbCrLf, -1, vbBinaryCompare)
    
    Dim i As Integer
    strResult = ""
    
    For i = 0 To UBound(arrChar)
    
        If arrChar(i) <> "" Then
        
            Dim col() As String
            col = VBA.Split(arrChar(i), ":")
            If UBound(col) = 1 Then
            
                Dim Pos() As String
                
                Pos = Split(col(0), "_")
                strResult = strResult & "{""x"":""" & Pos(0) & """,""y"":""" & Pos(1) & """,""Char"":""" & col(1) & """},"
            
            
            Else
            
            End If
        
        
        
        End If
    
    
    Next
    
    If strResult <> "" Then
    
        strResult = "{""result"":[" & Left(strResult, Len(strResult) - 1) & "]}"
    
    End If
parseResultToJson = strResult

End Function
