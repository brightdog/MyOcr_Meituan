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
                    '�����0�Ļ�����˵���������ļ������У�������������ɨ����ȥ��
                    StartYaxis = 0
                End If
            End If
            
        Next

        Set horizontalSplitImg = dicLine
        
        'debug�����Ϣ
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

        '�Ȱ�������ɫ���ֵĴ����ó���������һ��ͳ��
118
        Dim bolFindPix As Boolean
        bolFindPix = False
        Dim StartXaxis As Long
        StartXaxis = 0
        
120     For x = 1 To iWidth
            '����x�ᣬһ��һ��ɨ�����飨ͼ��,������ĳһ��ȫ��Ϊ0��û����ɫ��������Ϊ��һ���ǿɷָ�ġ�
            'ֱ��ɨ�赽ĳһ�г�������1�������ˣ�����Ϊ���������ݡ�
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

                '����һ������
                If StartXaxis > 0 Then
                    '��� StartYaxis ��Ϊ0������Ϊ��ǰ���������ݵģ���Ҫ����ǰ�����ݸ�����������
                
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
                    '�����0�Ļ�����˵���������ļ������У�������������ɨ����ȥ��
                    StartXaxis = 0
                End If
            End If
            
        Next

        Set verticalSplitImg = dicChar
        
        'debug�����Ϣ
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
'        '��Y�᷽�򣬺�X�᷽�򣬴��������ƽ������ֲ�ֳ�����
'        '�Ƚϼ򵥣��򵥰汾�������Ǳʻ��ص������ߴ�Խ�����
'        'Ĭ������£�ÿһ��������֮�䣬�������㹻�Ŀհ��������ģ�
'
'        'Ĭ������λ�þ���X,Y����¼����Ϊ��ǰ���л���ʱ�򣬾��Ѿ��������������ˡ�
'        Dim iFirst As Long  '��¼��ǰ���ҵ����ֵĵ�һ������Xλ��
'
'        '���������и���ﲻ�������µĿհף��Ⱥ����ٴ�������ʵ������Ҳ��Ӱ��ʲô�������������Ǵ�Ů���ģ��Ƚ���������Ҫ����һ�£���������᲻����ģ���
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
'116                 If arr(x, y) <> 0 Then  '�˴�������ĳһ���ֵ���ʼλ��
'
'                        'ֻҪ��һ���������ݣ�����Ϊ��ʼ�ˣ�
'118                     iFirst = x
'120                     bolFindaWord = True
'                        Exit For '��һ�оͿ��Բ�ɨ���ˣ���Լʱ�䣡
'                    End If
'
'                Else
'
'                    '                '��ʼλ���Ѿ����ˣ�����Ҫ������ֵĽ���λ��
'                    '                If arr(x, y) <> 0 Then
'                    '                    '������һ�����еĵ㣬��Ϊ�գ�����Ϊǰһ��Ϊ��ǰ�ֵ����һ������
'                    '                    bolHavePix = True
'                    '                    Exit For '�����ֵ��ұ�Եʧ�ܣ�ֱ��ɨ����һ�У���һ�оͿ��Բ�ɨ���ˣ���Լʱ�䣡
'                    '                End If
'                End If
'
'            Next
'
'122         If bolFindaWord Then  'And Not bolHavePix Then
'
'                '�ҵ��˵�ǰ�ֵĽ�β���ɴ�����ֵ�x��ɾ���iFirst -> x -1
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

