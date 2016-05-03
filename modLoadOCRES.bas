Attribute VB_Name = "modLoadOCRES"
Option Explicit

Public Function LoadOCRES(ByVal OCRESpath As String) As Scripting.Dictionary

    '�°��ļ���ʽ��
    '{"Word":"�ַ�","Config":[{"Blank":"0�ĸ���","Pixel":"1�ĸ���","RAW":"ԭʼ���ݣ��ϲ���1�У�"},{"Zero":"0�ĸ���","Pixel":"1�ĸ���","RAW":"ԭʼ���ݣ��ϲ���1�У�"}]}
    '���ڿ��Կ��������Ӹ����ά�ȣ�����Ч��
    'һ���ַ����Զ�Ӧ���ģʽ����ֹ��Щ��վ�ֿ����ݸ���
    
    Dim Fso As Scripting.FileSystemObject
    Set Fso = New Scripting.FileSystemObject
    Dim TS As Scripting.TextStream
    
    Dim dicResult As Scripting.Dictionary
    Set dicResult = New Scripting.Dictionary
    
    Dim strResult As String
    Set TS = Fso.OpenTextFile(App.Path & "\OCRES\" & OCRESpath, ForReading, True, TristateFalse)

    If Not TS.AtEndOfStream Then
        strResult = TS.ReadLine
    
        If strResult <> "" Then
            Dim dicSingleWord As Scripting.Dictionary
            Set dicSingleWord = JSON.Parse(strResult)
            Dim v As Variant

            For Each v In dicSingleWord.keys
            
                If Not dicResult.Exists(VBA.CStr(v)) Then
        
                    dicResult.Add VBA.CStr(v), dicSingleWord.Item(v)
        
                Else
        
                End If
            
            Next
            
        End If

    Else
    
    End If

    Set Fso = Nothing
    
    Set LoadOCRES = dicResult
    
End Function
