Attribute VB_Name = "modFilterColor"
Option Explicit

Public Function FilterBackGround(ByRef PicBox As VB.PictureBox) As Variant
        'ͨ�ô����򵥰ѵ�ɫ��ȥ����������ǰͼƬ�г��ִ�����2����ɫ
        Dim iWidth, iHeight, x, y As Long
        
        iWidth = PicBox.ScaleWidth
108     iHeight = PicBox.ScaleHeight
    
        Dim dicPicture As Scripting.Dictionary
110     Set dicPicture = New Scripting.Dictionary
        Dim arrImageColor() As Long
    
112     ReDim arrImageColor(1 To iWidth, 1 To iHeight) As Long
    
        Dim arrImageColorConvert() As Integer
    
114     ReDim arrImageColorConvert(1 To iWidth, 1 To iHeight) As Integer
    
        Dim iHDC As Long
116     iHDC = PicBox.hdc
    
        Dim iRGB As Long

        '�Ȱ�������ɫ���ֵĴ����ó���������һ��ͳ��
118     For x = 1 To iWidth
    
120         For y = 1 To iHeight
122             iRGB = GetPixel(iHDC, x, y)
124             arrImageColor(x, y) = iRGB '.Add x & "|" & y, Red(iRGB) & "|" & Green(iRGB) & "|" & Blue(iRGB)

                '                If x >= 19 And y >= 49 Then
                '                    Debug.Print iRGB
                '                End If

126             If dicPicture.Exists(iRGB) Then
            
128                 dicPicture.Item(iRGB) = dicPicture.Item(iRGB) + 1
                Else
            
130                 dicPicture.Add iRGB, 1
            
                End If
            
            Next
    
        Next
    
        Dim v As Variant
        Dim iMax As Long
132     iMax = 0

        '�����ִ���������ɫ����ȡ��������Ϊ��ɫ����Ҫ�ɵ��ġ�
134     For Each v In dicPicture.keys
    
136         If v > iMax Then
        
138             iMax = v
            End If
    
        Next
    
        'Me.Picture2.BackColor = iMax
        '����ɫ�ɵ���
        'ֻ������ǰͼƬ�г��ִ��������϶�����ɫ
        '����ȥ��һЩ��ɫֵΪ��-1 �� 0 �Ķ�����
140     For x = 1 To iWidth
    
142         For y = 1 To iHeight

144             If arrImageColor(x, y) = iMax Then
146                 'Call SetPixelV(iHDC, x, y, 16777215)
148                 arrImageColorConvert(x, y) = 0
                Else

                    If arrImageColor(x, y) > 0 Then
150                     'Call SetPixelV(iHDC, x, y, 0)
152                     arrImageColorConvert(x, y) = 1
                    Else
                        'Call SetPixelV(iHDC, x, y, 16777215)
                        arrImageColorConvert(x, y) = 0
                    End If
                End If
            
            Next
    
        Next

        FilterBackGround = arrImageColorConvert
        
        'debug ����Ϣ
        Dim strResult As String
        
        '��ȥ������հ׵���Ч�����ó�����
        
        If gbolNeedOutput Then

154         For y = 1 To iHeight
    
156             For x = 1 To iWidth
    
158                 strResult = strResult & arrImageColorConvert(x, y)
    
                Next
    
160             strResult = strResult & vbCrLf
    
            Next
    
162         Open App.Path & "\xxx.txt" For Output As #1
164         Print #1, strResult
166         Close #1
        End If

168     'PicBox.Refresh

End Function
