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

        '�Ȱ�������ɫ���ֵĴ����ó���������һ��ͳ��
118
        Dim bolFindPix As Boolean
        bolFindPix = False
        Dim StartYaxis As Long
        StartYaxis = 0
        
120     For y = 1 To iHeight
            '����Y�ᣬһ��һ��ɨ�����飨ͼ��,������ĳһ��ȫ��Ϊ0��û����ɫ��������Ϊ��һ���ǿɷָ�ġ�
            'ֱ��ɨ�赽ĳһ�г�������1�������ˣ�����Ϊ���������ݡ�
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
                '����һ������
                If StartYaxis > 0 Then
                '��� StartYaxis ��Ϊ0������Ϊ��ǰ���������ݵģ���Ҫ����ǰ�����ݸ�����������
                
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
                '�����0�Ļ�����˵���������ļ������У�������������ɨ����ȥ��
                    StartYaxis = 0
                End If
            End If
            
        Next
        

        Set horizontalSplitImg = dicLine
        
        'debug�����Ϣ
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
