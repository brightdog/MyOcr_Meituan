Attribute VB_Name = "modOCR"
Option Explicit

Public Function getOCRResult(ByRef dic As Scripting.Dictionary, ByRef dicOCRES As Scripting.Dictionary, Optional ByVal distence As Integer = 10) As String
    '第一个是当前待识别的对象
    '第二个是先前load好的OCR字库
    '第三个是ocr允许的误差LD值

    Dim v As Variant
    
    'Dim objDicToJson As New clsDicToJson 'for debug only
    'Debug.Print objDicToJson.DicToJson(dicOCRES)
    
    For Each v In dicOCRES.keys

        If Not IsEmpty(v) Then
            Dim i As Integer
        
            For i = 1 To dicOCRES.Item(v).Count
        
                Dim dicLocal As Scripting.Dictionary
                Set dicLocal = dicOCRES.Item(v).Item(i)

                If dic.Item("RAW") = dicLocal.Item("RAW") Then
                    Debug.Print "DIRECT"
                    getOCRResult = v
                    Exit Function
                        
                ElseIf VBA.Abs(dicLocal.Item("Pixel") - dic.Item("Pixel")) + VBA.Abs(dicLocal.Item("Blank") - dic.Item("Blank")) < (distence / 2) Then
                    Dim strUnknow As String
                    Dim strDic As String
                    strUnknow = Replace(dic.Item("RAW"), "&", "")
                    strDic = Replace(dicLocal.Item("RAW"), "&", "")
                    'Debug.Print "当前疑似匹配：" & v
                    Dim iLen As Integer
                    iLen = VBA.CInt(Len(dic.Item("RAW")) & 60)

                    If VBA.Left(dic.Item("RAW"), iLen) = VBA.Left(dicLocal.Item("RAW"), iLen) Then
                        Debug.Print "LEFT60%"
                        getOCRResult = v
                        Exit Function
                    
                    ElseIf modLD.LD(dic.Item("RAW"), dicLocal.Item("RAW")) < distence Then
                        Debug.Print "HOLEMATCH"
                        getOCRResult = v
                        Exit Function
                    End If
            
                End If
            
            Next

        End If

    Next

    getOCRResult = ""
End Function

Public Function PreOcr(ByVal RAW As Variant) As Scripting.Dictionary

    Dim x, y As Long
    
    Dim intBlankCount, intPixelCount As Integer
    intPixelCount = 0
    intBlankCount = 0
    Dim strSerial As String
    
    strSerial = JoinRawData(RAW)
    '预先把图片里的像素和空白的数量都取出来，后期匹配的时候可以用

    intPixelCount = GetSerialPixCount(strSerial, 1)

    intBlankCount = GetSerialPixCount(strSerial, 0)
    
    Dim strResult As String
    strResult = "{""Blank"":""" & intBlankCount & """,""Pixel"":""" & intPixelCount & """,""RAW"":""" & JoinRawData(RAW) & """}"
    
    Set PreOcr = JSON.Parse(strResult)

End Function
