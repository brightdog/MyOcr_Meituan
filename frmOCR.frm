VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOCR 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   1320
   ClientTop       =   1395
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   12885
   Begin VB.ComboBox cboOCRES 
      Height          =   300
      ItemData        =   "frmOCR.frx":0000
      Left            =   2040
      List            =   "frmOCR.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CheckBox chkNeedOutput 
      Caption         =   "Need Output"
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtOutput 
      Height          =   4935
      Left            =   4860
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   900
      Width           =   7875
   End
   Begin VB.PictureBox picOutput 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtRAW 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   6.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmOCR.frx":0043
      Top             =   2340
      Width           =   1395
   End
   Begin VB.CommandButton cmdOCR 
      Caption         =   "OCR"
      Enabled         =   0   'False
      Height          =   915
      Left            =   1800
      TabIndex        =   2
      Top             =   4800
      Width           =   2115
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1800
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   1
      Top             =   60
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3300
      Top             =   3180
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOpenPicture 
      Caption         =   "open picture"
      Height          =   732
      Left            =   1800
      TabIndex        =   0
      Top             =   3600
      Width           =   2115
   End
   Begin VB.Label lblTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   2940
      Width           =   2475
   End
End
Attribute VB_Name = "frmOcr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pngClass As New LoadPNG

Dim mstrOCRFile As String

Dim mdblStartTime As Double
Dim mintHeight As Integer
Dim mintWidth As Integer
Dim mintTopOffset As Integer

Dim CurrentOCRES As String


Private Sub chkNeedOutput_Click()
gbolNeedOutput = Me.chkNeedOutput.value
End Sub

Private Sub cmdOpenPicture_Click()
    Dim filename As String
    CommonDialog1.ShowOpen
    filename = CommonDialog1.filename

    If filename <> "" Then

        Call LoadPic(filename)
        
        Me.cmdOCR.Enabled = True
    End If

End Sub

Private Sub LoadPic(ByVal picName As String)

    If LCase(Right(picName, 3)) = "png" Then
    
        pngClass.PicBox = picSource 'or Picturebox
        'pngClass.SetToBkgrnd True, 0, 0 'set to Background (True or false), x and y
        'pngClass.BackgroundPicture = Form1 'same Backgroundpicture
        'pngClass.SetAlpha = True 'when Alpha then alpha
        'pngClass.SetTrans = True 'when transparent Color then transparent Color
        pngClass.OpenPNG picName 'Open and display Picture
        
    Else
        picSource.Picture = LoadPicture(picName)
           
    End If

    picOutput.width = picSource.width
End Sub

'width5 = PicBox.ScaleWidth
'heigh5 = PicBox.ScaleHeight
'hdc5 = PicBox.hDC
'
'For i = 0 To width5
'    For j = 0 To heigh5
'        rgb5 = GetPixel(hdc5, i, j)
'
'        '一旦发现红黄蓝中,任意一个颜色的值小于阀值,就定义为白色
'        If Blue(rgb5) < lngEdge Or Red(rgb5) < lngEdge Or Green(rgb5) < lngEdge Then
'            y = 0
'        Else
'            y = 255
'        End If
'
'        '将灰度转换为RGB
'        rgb5 = RGB(y, y, y)
'        SetPixelV hdc5, i, j, rgb5     '将黑白图片回传给PICTUREBOX
'    Next
'Next

Private Sub cmdOCR_Click()
    Dim T As clsMSTimer
    Set T = New clsMSTimer
    T.StartTimer
    CurrentOCRES = "Meituan_Hotel_" & Split(Me.cboOCRES.Text, "_", 3)(2)

    If Me.cboOCRES.Text = "" Then
    
        Me.txtOutput.Text = StartOCR(CurrentOCRES)
    Else
        Me.txtOutput.Text = StartOCR(CurrentOCRES)
    End If

    'Me.txtOutput.Text = StartOCR("Meituan_Hotel_Price.OCRES")
    Me.lblTime.Caption = T.ShowMS / 1000 & "ms"
    Set T = Nothing
End Sub

Public Function StartOCR(ByVal OCRESName As String) As String
100     Me.lblTime.Caption = ""
102     Debug.Print Now
104     mdblStartTime = VBA.Timer()
        Dim x As Long
        Dim y As Long
    
        Dim iWidth As Long
        Dim iHeight As Long
        Dim arrImage As Variant
        
        arrImage = FilterBackGround(Me.picSource)
        Dim dicLine As Scripting.Dictionary
        Set dicLine = horizontalSplitImg(arrImage)
        '按行拆分已经好了。现在把每一行循环一下，把每个字都拆出来。
        Dim v As Variant
        Dim dicWords As Scripting.Dictionary
        Dim dicResult As Scripting.Dictionary
        Set dicResult = New Scripting.Dictionary
        
        For Each v In dicLine.keys
        
            Set dicWords = verticalSplitImg(dicLine.Item(v).Content)
            
            Dim k As Variant
            
            For Each k In dicWords.keys
            
                dicResult.Add k & "_" & dicLine.Item(v).Top, dicWords.Item(k).Content
            
            Next
            
        Next
        
        Dim dicOCRES As Scripting.Dictionary
        Set dicOCRES = LoadOCRES(OCRESName)
        
        For Each v In dicResult.keys

            If Not IsEmpty(v) Then
                Dim dic As Scripting.Dictionary
                Set dic = PreOcr(dicResult.Item(v))
                Dim strResult As String
ReOCR:
                strResult = getOCRResult(dic, dicOCRES, 10)
            
                If strResult = "" Then
                    Me.txtRAW.Text = VBA.Replace(dic.Item("RAW"), "&", vbCrLf)
                    Dim strNewWord As String

160                 strNewWord = InputBox("这个是什么字？", "输入当前字符")

                    If Len(strNewWord) = 1 Then
164                     Call WriteLocalDic(strNewWord, dic, dicOCRES, OCRESName)
                    
166                     Set dicOCRES = LoadOCRES(OCRESName)
                
                    End If

                    GoTo ReOCR
                
                End If
            
                dicResult.Item(v) = strResult
            End If

        Next
        
        strResult = ""

        For Each v In dicResult.keys
        
            strResult = strResult & v & ":" & dicResult.Item(v) & vbCrLf
        
        Next

        StartOCR = parseResultToJson(strResult)

184     Debug.Print Now

End Function
'
'Private Function OCR(ByVal dicWordData As Scripting.Dictionary, ByVal OCRFile As String, Optional ByVal MathValue As Integer = 10) As String
'ReStart:
'        Dim Fso As Scripting.FileSystemObject
'100     Set Fso = New Scripting.FileSystemObject
'        Dim strLocalData As String
'
'        Dim bolNeedStart As Boolean
'102     bolNeedStart = False
'
'104     If Fso.FileExists(App.Path & "\OCRES\" & OCRFile) Then
'            Dim tsDic As Scripting.TextStream
'106         Set tsDic = Fso.OpenTextFile(App.Path & "\OCRES\" & OCRFile)
'
'108         If Not tsDic.AtEndOfStream Then
'110             strLocalData = Fso.OpenTextFile(App.Path & "\OCRES\" & OCRFile).ReadAll
'            Else
'112             strLocalData = ""
'            End If
'
'        Else
'114         strLocalData = ""
'116         Call Fso.CreateTextFile(App.Path & "\OCRES\" & OCRFile)
'            '文件不存在，需要额外处理！
'        End If
'
'118     Set Fso = Nothing
'
'        Dim dicResult As Scripting.Dictionary
'120     Set dicResult = New Scripting.Dictionary
'
'        Dim arrLine() As String
'
'122     arrLine = Split(strLocalData, vbCrLf, -1, vbBinaryCompare)
'
'        Dim iLines As Integer
'
'124     iLines = UBound(arrLine)
'
'        'If iLines >= 0 Then
'        '这里开始循环那个字典，把所有匹配出来的数据，到本地字库里去滚一圈，把字都匹配出来！
'        Dim v As Variant
'        Dim strRaw As String
'
'126     For Each v In dicWordData.keys
'
'128         strRaw = JoinRawData(dicWordData.Item(v))
'            Dim bolFindWord As Boolean
'130         bolFindWord = False
'
'132         If strLocalData <> "" Then
'
'                Dim i As Integer
'
'134             For i = 0 To iLines
'
'                    Dim strDic As String
'                    Dim strWord As String
'                    Dim arrtmp() As String
'136                 arrtmp = Split(arrLine(i), "|", 2, vbBinaryCompare)
'
'138                 If UBound(arrtmp) > 0 Then
'
'140                     strWord = arrtmp(0)
'142                     strDic = arrtmp(1)
'
'144                     If strRaw = strDic Then
'146                         dicResult.Add v, strWord
'148                         bolFindWord = True
'                            Exit For
'                        Else
'
'150                         If modLD.LD(CompressSerial(strRaw, True), CompressSerial(strDic, True)) < 10 Then
'
'152                             dicResult.Add v, strWord
'154                             bolFindWord = True
'                                Exit For
'                            End If
'                        End If
'
'                    Else
'
'                        '本地字典中的数据异常，直接丢弃，不能影响程序的继续执行
'                    End If
'
'                Next
'
'            End If
'
'156         If Not bolFindWord Then
'
'158             Me.txtRAW.Text = MakeSymbol(strRaw, mintWidth)
'                Dim strNewWord As String
'
'                'Do
'160             strNewWord = InputBox("这个是什么字？", "输入当前字符")
'
'162             'Loop While Len(strNewWord) <> 1
'                If Len(strNewWord) = 1 Then
'164                 Call WriteLocalDic(strNewWord, strRaw, OCRFile)
'                    'dicResult.Add v, strWord
'166                 bolNeedStart = True
'
'                    Exit For
'                End If
'            End If
'
'        Next
'
'168     If bolNeedStart Then
'170         GoTo ReStart
'        End If
'
'        'Else
'
'        '文件里面是空的，需要额外处理一下
'
'        'End If
'        Dim strResult As String
'
'172     picOutput.Cls
'
'174     For Each v In dicResult.keys
'
'176         If Me.chkNeedOutput.value = 1 Then
'178             picOutput.CurrentX = v
'180             picOutput.CurrentY = 0
'182             picOutput.Print dicResult.Item(v)
'            End If
'
'184         strResult = strResult & """" & v & """:""" & dicResult.Item(v) & ""","
'
'        Next
'
'186     strResult = Left(strResult, Len(strResult) - 1)
'
'188     OCR = "{" & strResult & "}"
'
'End Function

Private Function GetWordData(ByRef arr() As Integer, ByVal iFirst As Long, ByVal iWidth As Long) As Integer()
        '<EhHeader>
        On Error GoTo GetWordData_Err
        '</EhHeader>

        Dim arrResult() As Integer
    
        Dim iHeight As Integer
100     iHeight = UBound(arr, 2)
    
102     ReDim arrResult(1 To iWidth, mintTopOffset To iHeight) As Integer
        'Dim UboundARR As Long
        'UboundARR = UBound(arr, 1)
        Dim x As Long
        Dim y As Long
    
104     For x = 1 To iWidth
    
106         For y = mintTopOffset To iHeight
                'If iFirst + x - 1 <= UboundARR Then
108             arrResult(x, y) = arr(x + iFirst - 1, y)
                'End If
            Next
        Next
    
110     GetWordData = arrResult
    
        '<EhFooter>
        Exit Function

GetWordData_Err:
        MsgBox Erl
        End
        '</EhFooter>
End Function

Private Function Blue(ByVal mlColor As Long) As Long
    ''从RGB值中获得蓝色值
    Blue = (mlColor \ &H10000) And &HFF
End Function

Private Function Green(ByVal mlColor As Long) As Long
    '从RGB值中获得绿色值
    Green = (mlColor \ &H100) And &HFF
End Function
    
Private Function Red(ByVal mlColor As Long) As Long
    '从RGB值中获得红色值
    Red = mlColor And &HFF
End Function

Private Sub Form_Load()
        Dim strCommand As String
100     strCommand = VBA.Command
    '        '带参数执行时，自动识别+返回
102     If strCommand <> "" Then

                Dim strFileName As String
                
104             strFileName = Split(strCommand, "\")(UBound(Split(strCommand, "\")))

106             Call LoadPic(strCommand)
                Dim strResult As String
                Dim arrName() As String
                
                arrName = Split(strFileName, "_", 4)
108             strResult = StartOCR(arrName(0) & "_" & arrName(1) & "_" & arrName(2) & ".OCRES")

                Dim iFile As Integer

110             iFile = VBA.FreeFile()
112             Open strCommand & ".json" For Output As #iFile
114             Print #iFile, strResult
116             Close #iFile

118             End

        End If
End Sub
