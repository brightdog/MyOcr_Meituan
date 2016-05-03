Attribute VB_Name = "modSplitImg"
Option Explicit

Public Function horizontalSplitImg(ByRef arrImage As Variant) As Scripting.Dictionary
        
        Dim dicLine As Scripting.Dictionary
        Set dicLine = New Scripting.Dictionary
        
        Dim iWidth As Long
        Dim iHeight As Long
        
        Dim x, y As Long
        
        iWidth = UBound(arrImage, 1)
        iHeight = UBound(arrImage, 2)

        '先把所有颜色出现的次数拿出来，并做一个统计
118
        Dim bolFindPix As Boolean
        bolFindPix = False
        Dim StartYaxis As Long
        StartYaxis = 0
        
120     For y = 1 To iHeight
            '锁定Y轴，一行一行扫描数组（图像）,若发现某一行全部为0（没有颜色），则认为这一行是可分割的。
            '直到扫描到某一行出现至少1个像素了，才认为以下有内容。
            bolFindPix = False

            For x = 1 To iWidth

122             If arrImage(x, y) > 0 Then
                    If StartYaxis = 0 Then
                        StartYaxis = y
                    End If

                    bolFindPix = True
                    Exit For
                    
                End If
            
            Next
            
            If Not bolFindPix Then

                '发现一个空行
                If StartYaxis > 0 Then
                    '如果 StartYaxis 不为0，则认为先前都是有内容的，需要把先前的内容给保留下来。
                
                    Dim objPicLine As clsPic
                    Set objPicLine = New clsPic
                    objPicLine.Top = StartYaxis
                    objPicLine.Left = 0
                    objPicLine.width = UBound(arrImage, 1)
                    objPicLine.Hight = (y - 1) - StartYaxis + 1
                    objPicLine.Content = getPartOfArray(arrImage, , , StartYaxis, y - 1)
                    dicLine.Add StartYaxis, objPicLine
                    StartYaxis = 0
                Else
                    '如果是0的话，那说明是连续的几个空行，不管它，继续扫描下去。
                    StartYaxis = 0
                End If
            End If
            
        Next

        Set horizontalSplitImg = dicLine
        
        'debug输出信息
        If gbolNeedOutput Then
            Dim v As Variant
            Dim iFreeFile As Integer

            For Each v In dicLine.keys
            
                iFreeFile = VBA.FreeFile
                Open App.Path & "\tmp_" & v & ".txt" For Append As #iFreeFile
                Dim arrtmp As Variant
                arrtmp = dicLine.Item(v).Content

                For y = 1 To UBound(arrtmp, 2)
                    Dim Tmp As String
                    Tmp = ""

                    For x = 1 To UBound(arrtmp, 1)
                        Tmp = Tmp & arrtmp(x, y)
                    Next

                    Print #iFreeFile, Tmp
                Next

                Close #iFreeFile
            Next

        End If

End Function

Public Function verticalSplitImg(ByRef arrImage As Variant) As Scripting.Dictionary
        
        Dim dicChar As Scripting.Dictionary
        Set dicChar = New Scripting.Dictionary
        
        Dim iWidth As Long
        Dim iHeight As Long
        
        Dim x, y As Long
        
        iWidth = UBound(arrImage, 1)
        iHeight = UBound(arrImage, 2)

        '先把所有颜色出现的次数拿出来，并做一个统计
118
        Dim bolFindPix As Boolean
        bolFindPix = False
        Dim StartXaxis As Long
        StartXaxis = 0
        
120     For x = 1 To iWidth
            '锁定x轴，一行一行扫描数组（图像）,若发现某一行全部为0（没有颜色），则认为这一行是可分割的。
            '直到扫描到某一行出现至少1个像素了，才认为以下有内容。
            bolFindPix = False

            For y = 1 To iHeight

122             If arrImage(x, y) > 0 Then
                    If StartXaxis = 0 Then
                        StartXaxis = x
                    End If

                    bolFindPix = True
                    Exit For
                    
                End If
            
            Next
            
            If Not bolFindPix Then

                '发现一个空行
                If StartXaxis > 0 Then
                    '如果 StartYaxis 不为0，则认为先前都是有内容的，需要把先前的内容给保留下来。
                
                    Dim objPic As clsPic
                    Set objPic = New clsPic
                    objPic.Top = 0
                    objPic.Left = StartXaxis
                    objPic.width = (x - 1) - StartXaxis + 1
                    objPic.Hight = UBound(arrImage, 2)
                    objPic.Content = getPartOfArray(arrImage, StartXaxis, x - 1)
                    dicChar.Add StartXaxis, objPic
                    StartXaxis = 0
                Else
                    '如果是0的话，那说明是连续的几个空行，不管它，继续扫描下去。
                    StartXaxis = 0
                End If
            End If
            
        Next

        Set verticalSplitImg = dicChar
        
        'debug输出信息
        If gbolNeedOutput Then
            Dim v As Variant
            Dim iFreeFile As Integer

            For Each v In dicChar.keys
            
                iFreeFile = VBA.FreeFile
                Open App.Path & "\tmp_" & v & ".txt" For Append As #iFreeFile
                Dim arrtmp As Variant
                arrtmp = dicChar.Item(v).Content

                For y = 1 To UBound(arrtmp, 2)
                    Dim Tmp As String
                    Tmp = ""

                    For x = 1 To UBound(arrtmp, 1)
                        Tmp = Tmp & arrtmp(x, y)
                    Next

                    Print #iFreeFile, Tmp
                Next

                Close #iFreeFile
            Next

        End If

End Function

'
'Private Function SplitWordByWidth(ByRef arr() As Integer, ByVal iEachWordWidth As Integer) As Scripting.Dictionary
'        '<EhHeader>
'        On Error GoTo SplitWord_Err
'        '</EhHeader>
'
'        Dim dicResult As Scripting.Dictionary
'100     Set dicResult = New Scripting.Dictionary
'
'        Dim iWidth As Long
'102     iWidth = UBound(arr, 1)
'        Dim iHeight As Long
'104     iHeight = UBound(arr, 2)
'
'        Dim x As Long
'        Dim y As Long
'
'        '先Y轴方向，后X轴方向，从左向右推进，把字拆分出来！
'        '比较简单，简单版本，不考虑笔画重叠和杂线穿越的情况
'        '默认情况下，每一个字与字之间，都是有足够的空白来区隔的！
'
'        '默认像素位置就用X,Y来纪录，因为先前序列化的时候，就已经按照像素来做了。
'        Dim iFirst As Long  '纪录当前被找到的字的第一个像素X位置
'
'        '先做横向切割，这里不考虑上下的空白，等后面再处理！（其实不处理也不影响什么。。。但是我是处女座的，比较作！还是要处理一下，否则心里会不舒服的！）
'106     iFirst = 0
'
'        Dim bolFindaWord As Boolean
'108     bolFindaWord = False
'
'110     For x = 1 To iWidth
'            '        Dim bolHavePix As Boolean
'            '        bolHavePix = False
'
'112         For y = 1 To mintHeight
'
'114             If Not bolFindaWord Then
'
'116                 If arr(x, y) <> 0 Then  '此处发现了某一个字的起始位置
'
'                        '只要有一个点有内容，就认为开始了！
'118                     iFirst = x
'120                     bolFindaWord = True
'                        Exit For '这一列就可以不扫描了，节约时间！
'                    End If
'
'                Else
'
'                    '                '起始位置已经有了，现在要找这个字的结束位置
'                    '                If arr(x, y) <> 0 Then
'                    '                    '必须这一列所有的点，都为空，则认为前一列为当前字的最后一列像素
'                    '                    bolHavePix = True
'                    '                    Exit For '发现字的右边缘失败，直接扫描下一列，这一列就可以不扫描了，节约时间！
'                    '                End If
'                End If
'
'            Next
'
'122         If bolFindaWord Then  'And Not bolHavePix Then
'
'                '找到了当前字的结尾，由此这个字的x组成就是iFirst -> x -1
'                '            If dicResult.Exists(iFirst) Then
'                '                dicResult.Item(iFirst) = GetWordData(arr, iFirst, x - 1)
'                '                bolFindaWord = False
'                '            Else
'124             dicResult.Add iFirst, GetWordData(arr, iFirst, iEachWordWidth)
'
'126             bolFindaWord = False
'128             x = x + mintWidth
'                '            End If
'
'            End If
'
'        Next
'
'130     Set SplitWord = dicResult
'
'        '<EhFooter>
'        Exit Function
'
'SplitWord_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in MyOCR.Form1.SplitWord " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Function

