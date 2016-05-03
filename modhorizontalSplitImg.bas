Attribute VB_Name = "modhorizontalSplitImg"
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
                
                    Dim objPicLine As clsLine
                    Set objPicLine = New clsLine
                    objPicLine.Top = StartYaxis
                    objPicLine.Left = 0
                    objPicLine.Width = UBound(arrImage, 1)
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
        Dim v As Variant
        Dim iFreeFile As Integer
        For Each v In dicLine.Keys
            
            
            iFreeFile = VBA.FreeFile
            Open App.Path & "\tmp_" & v & ".txt" For Append As #iFreeFile
            
            For x = 0 To UBound(dicLine.Item(v).Content, 1)
            
                For y = 0 To UBound(dicLine.Item(v).Content, 2)
                    Print #iFreeFile, dicLine.Item(v).Content(x, y)
                Next
            Next
            Close #iFreeFile
        Next
        
End Function
