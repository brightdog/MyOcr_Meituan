Attribute VB_Name = "modLoadOCRES"
Option Explicit

Public Function LoadOCRES(ByVal OCRESpath As String) As Scripting.Dictionary

    '新版文件格式：
    '{"Word":"字符","Config":[{"Blank":"0的个数","Pixel":"1的个数","RAW":"原始数据（合并成1行）"},{"Zero":"0的个数","Pixel":"1的个数","RAW":"原始数据（合并成1行）"}]}
    '后期可以考虑再增加更多的维度，增加效率
    '一个字符可以对应多个模式，防止有些网站字库内容复杂
    
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
