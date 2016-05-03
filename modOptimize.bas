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
    '������ʽ��ѹ�������ֵ�洢���ȡ�
    '��ʽ��1x2,0x18,...
    '���ԭʼ��ͬ��һ��С��4�����أ���Ҫ��һ�����ŷָ��������ܻ��������һ��㣬�ڿ����Ƿ����Ƶ������ѹ���������Ƿ�ᱻ�˼Ҳ³���ͼ�أ�
    '���ԭʼ��ͬ��һ�ε���4�����أ���ѹ�벻ѹ�����ͬ�����ܻ�����һ�����ڴ����CPU���ġ�����
    '��ȫ��ѹ��һ�飬����Ч����˵�ˡ�--���������������У�����Ҫ�ֱ��������������û���Ա�С���Լ۱�̫�ͣ�
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

        If iDotPos > 1 Then '�����Ǵӵ�ǰ���ֵĵ�2���ֽڿ�ʼ���𣬷����ʽ������
        
            Dim arrtmp() As String
            
            arrtmp = Split(arrPart(i), ".", 2, vbBinaryCompare) 'ֻ������2���֣�������߲�����Ҫ�ų���Щδ��ѹ���Ĵ���ֻҪ�����һ��1��0�ľͿ��Ժͺ�벿�ֽ�������ˡ�
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
            
                '�Ƿ���~~~~
            
            End If
        
        Else

            '��Ҫ�ų�ȫ�����ֵģ���Ϊ����δ��ѹ���ģ��Ϸ��Ĵ���
            If iDotPos < 1 Then
                strResult = strResult & arrPart(i)
            Else
                '��ʼ��һ���ֽھ���"."�޷�ʶ��
                '��ʽ����
            End If
        End If
    
    Next

    DecompressSerial = strResult
End Function
