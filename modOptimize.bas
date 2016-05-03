Attribute VB_Name = "modOptimize"
Option Explicit

Public Function GetSerialPixCount(ByVal strData As String, Optional ByVal Pix As Integer = 1) As Integer
    
    Dim intCount As Integer
    Dim i As Integer
    Dim iLen As Long
    iLen = Len(strData)
    For i = 1 To iLen
    
        If Mid(strData, i, 1) = Pix Then
            intCount = intCount + 1
        End If
    
    Next


    GetSerialPixCount = intCount

End Function



Public Function CompressSerial(ByVal strData As String, Optional isShrink As Boolean = False) As String
    '混淆格式，压缩本地字典存储长度。
    '格式：1x2,0x18,...
    '如果原始相同的一段小于4个像素（还要加一个逗号分割？），则可能会把体积变大一点点，在考虑是否类似的情况不压缩。但是是否会被人家猜出意图呢？
    '如果原始相同的一段等于4个像素，则压与不压体积相同，可能会增加一点点后期处理的CPU消耗。。。
    '先全部压缩一遍，看看效果再说了。--》测试下来，不行，还是要分别处理！！！否则体积没明显变小，性价比太低！
    Dim iCount As Integer
    Dim CurrentPix As Integer
    Dim LastPix As Integer
    Dim i As Long
    Dim strResult As String

    For i = 1 To Len(strData)
    
        CurrentPix = Mid(strData, i, 1)

        If i > 1 Then
            If CurrentPix = LastPix Then
                iCount = iCount + 1
            Else
            
                If iCount > 4 Then
                    If Not isShrink Then
                        strResult = strResult & LastPix & "." & iCount & ","
                    Else
                        strResult = strResult & LastPix & iCount
                    End If

                Else
                    strResult = strResult & Mid(strData, i - iCount, iCount)
                End If

                LastPix = CurrentPix
                iCount = 1
            End If

        Else
            iCount = 1
        End If

    Next

    strResult = strResult & LastPix & "." & iCount
    CompressSerial = strResult
End Function

Public Function DecompressSerial(ByVal strData As String) As String

    Dim arrPart() As String
    
    arrPart = Split(strData & ",", ",", -1, vbBinaryCompare)
    
    Dim i As Integer
    Dim strResult As String

    For i = 0 To UBound(arrPart)
        Dim iDotPos As Integer
        iDotPos = InStr(1, arrPart(i), ".", vbBinaryCompare)

        If iDotPos > 1 Then '必须是从当前部分的第2个字节开始算起，否则格式出错！！
        
            Dim arrtmp() As String
            
            arrtmp = Split(arrPart(i), ".", 2, vbBinaryCompare) '只有左右2部分，并且左边部分需要排除那些未经压缩的串，只要留最后一个1或0的就可以和后半部分结合起来了。
            Dim strFlatPart As String
            strFlatPart = Left(arrtmp(0), Len(arrtmp(0)) - 1)
            strResult = strResult & strFlatPart
            Dim strFlag As String
            strFlag = Right(arrtmp(0), 1)

            If strFlag = 1 Or strFlag = 0 Then
            
                Dim iCount As Integer
                
                For iCount = 1 To arrtmp(1)
                
                    strResult = strResult & strFlag
                
                Next
            
            Else
            
                '非法啊~~~~
            
            End If
        
        Else

            '需要排除全部数字的，因为那是未经压缩的，合法的串！
            If iDotPos < 1 Then
                strResult = strResult & arrPart(i)
            Else
                '起始第一个字节就是"."无法识别
                '格式出错！
            End If
        End If
    
    Next

    DecompressSerial = strResult
End Function
