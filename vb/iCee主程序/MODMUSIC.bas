Attribute VB_Name = "���ģ��"
Option Explicit
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, BITMAP As Long) As GpStatus
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Public Const SPI_GETWORKAREA = 48

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Const WU_LOGPIXELSX = 88
Public Const WU_LOGPIXELSY = 90

Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal name As Long, ByVal fontCollection As Long, fontFamily As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As StringAlignment) As GpStatus
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipAddPathStringI Lib "gdiplus" (ByVal Path As Long, ByVal str As Long, ByVal Length As Long, ByVal family As Long, ByVal style As Long, ByVal emSize As Single, layoutRect As RECTL, ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipFillPath Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As GpUnit, pen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipCreateLineBrush Lib "gdiplus" (Point1 As PointF, Point2 As PointF, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Declare Function GdipSetClipRectI _
               Lib "gdiplus" (ByVal graphics As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal Width As Long, _
                              ByVal Height As Long, _
                              ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipResetClip Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Enum CombineMode
    CombineModeReplace = 0
    CombineModeIntersect
    CombineModeUnion
    CombineModeXor
    CombineModeExclude
    CombineModeComplement
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum QualityMode
   QualityModeInvalid = -1
   QualityModeDefault = 0
   QualityModeLow = 1
   QualityModeHigh = 2
End Enum

Public Enum SmoothingMode
   SmoothingModeInvalid = QualityModeInvalid
   SmoothingModeDefault = QualityModeDefault
   SmoothingModeHighSpeed = QualityModeLow
   SmoothingModeHighQuality = QualityModeHigh
   SmoothingModeNone
   SmoothingModeAntiAlias
End Enum

Public Enum FillMode
   FillModeAlternate
   FillModeWinding
End Enum

Public Enum GpUnit
   UnitWorld
   UnitDisplay
   UnitPixel
   UnitPoint
   UnitInch
   UnitDocument
   UnitMillimeter
End Enum

Public Type RECTF
    Top     As Single
    Left    As Single
    Width   As Single
    Height  As Single
End Type

Public Enum FontStyle
   FontStyleRegular = 0
   FontStyleBold = 1
   FontStyleItalic = 2
   FontStyleBoldItalic = 3
   FontStyleUnderline = 4
   FontStyleStrikeout = 8
End Enum

Public Enum StringAlignment
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Public Type PointF
    X As Long
    Y As Long
End Type

Public Enum WrapMode
   WrapModeTile         ' 0
   WrapModeTileFlipX    ' 1
   WrapModeTileFlipY    ' 2
   WrapModeTileFlipXY   ' 3
   WrapModeClamp        ' 4
End Enum

Public Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (RECT As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal Mode As LinearGradientMode, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
Public Enum LinearGradientMode
   LinearGradientModeHorizontal          ' 0
   LinearGradientModeVertical            ' 1
   LinearGradientModeForwardDiagonal     ' 2
   LinearGradientModeBackwardDiagonal    ' 3
End Enum

Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
'��Ӱ��ˢ����
Public Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal style As HatchStyle, ByVal forecolr As Long, ByVal backcolr As Long, brush As Long) As GpStatus
Public Declare Function GdipGetHatchStyle Lib "gdiplus" (ByVal brush As Long, style As HatchStyle) As GpStatus
Public Declare Function GdipGetHatchForegroundColor Lib "gdiplus" (ByVal brush As Long, forecolr As Long) As GpStatus
Public Declare Function GdipGetHatchBackgroundColor Lib "gdiplus" (ByVal brush As Long, backcolr As Long) As GpStatus
Public Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
'----------------
Public Enum HatchStyle
   HatchStyleHorizontal                   ' 0
   HatchStyleVertical                     ' 1
   HatchStyleForwardDiagonal              ' 2
   HatchStyleBackwardDiagonal             ' 3
   HatchStyleCross                        ' 4
   HatchStyleDiagonalCross                ' 5
   HatchStyle05Percent                    ' 6
   HatchStyle10Percent                    ' 7
   HatchStyle20Percent                    ' 8
   HatchStyle25Percent                    ' 9
   HatchStyle30Percent                    ' 10
   HatchStyle40Percent                    ' 11
   HatchStyle50Percent                    ' 12
   HatchStyle60Percent                    ' 13
   HatchStyle70Percent                    ' 14
   HatchStyle75Percent                    ' 15
   HatchStyle80Percent                    ' 16
   HatchStyle90Percent                    ' 17
   HatchStyleLightDownwardDiagonal        ' 18
   HatchStyleLightUpwardDiagonal          ' 19
   HatchStyleDarkDownwardDiagonal         ' 20
   HatchStyleDarkUpwardDiagonal           ' 21
   HatchStyleWideDownwardDiagonal         ' 22
   HatchStyleWideUpwardDiagonal           ' 23
   HatchStyleLightVertical                ' 24
   HatchStyleLightHorizontal              ' 25
   HatchStyleNarrowVertical               ' 26
   HatchStyleNarrowHorizontal             ' 27
   HatchStyleDarkVertical                 ' 28
   HatchStyleDarkHorizontal               ' 29
   HatchStyleDashedDownwardDiagonal       ' 30
   HatchStyleDashedUpwardDiagonal         ' 31
   HatchStyleDashedHorizontal             ' 32
   HatchStyleDashedVertical               ' 33
   HatchStyleSmallConfetti                ' 34
   HatchStyleLargeConfetti                ' 35
   HatchStyleZigZag                       ' 36
   HatchStyleWave                         ' 37
   HatchStyleDiagonalBrick                ' 38
   HatchStyleHorizontalBrick              ' 39
   HatchStyleWeave                        ' 40
   HatchStylePlaid                        ' 41
   HatchStyleDivot                        ' 42
   HatchStyleDottedGrid                   ' 43
   HatchStyleDottedDiamond                ' 44
   HatchStyleShingle                      ' 45
   HatchStyleTrellis                      ' 46
   HatchStyleSphere                       ' 47
   HatchStyleSmallGrid                    ' 48
   HatchStyleSmallCheckerBoard            ' 49
   HatchStyleLargeCheckerBoard            ' 50
   HatchStyleOutlinedDiamond              ' 51
   HatchStyleSolidDiamond                 ' 52

   HatchStyleTotal
   HatchStyleLargeGrid = HatchStyleCross  ' 4

   HatchStyleMin = HatchStyleHorizontal
   HatchStyleMax = HatchStyleTotal - 1
End Enum
    
Private m_lngToken As Long
Public Type LRCROWINFO
    lrcString       As String       '������
    lrcTime         As Single       '����ʱ��
End Type

Public myLrc()         As LRCROWINFO
Public iLrcRows        As Integer      '��ʺ�ʱ������
Public iCurPlay        As Integer      '��ǰ���ŵ�����һ����

Public strTmp          As String       '������ʱ����
Public m_LastTime      As Double       '�ϴβ���ʱ�ĸ��ʱ��/�����ж��Ƿ�ع��˲���ʱ��
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Sub ClearLrc()
    Erase myLrc()               '�����ǰ�ĸ����Ϣ
    iLrcRows = 0                '��λ
    iCurPlay = 0
    m_LastTime = 0
End Sub

Public Sub StopLrc()
    iCurPlay = 0                '��λ
    m_LastTime = 0
End Sub

'�õ���ʺ�ʱ��
Private Sub SplitLrc(TempStr As String)
    Dim temp()  As String
    Dim i       As Integer, j   As Integer
    
    If TempStr <> "" Then
        TempStr = Replace(TempStr, vbTab, "")
        TempStr = Replace(TempStr, Chr(0), "")
        TempStr = Replace(TempStr, vbCr, "")
        TempStr = Replace(TempStr, vbLf, "")
        
        temp = Split(TempStr, "]")          '�ָ��ʺ�ʱ��
        j = UBound(temp)
        If j >= 1 Then                      '������ڸ�ʺ�ʱ��
            If InStr(temp(0), ":") Then
                For i = 0 To j - 1 '�ӵ�һ��ʱ�䵽���һ��ʱ��
                    ReDim Preserve myLrc(iLrcRows)
                    
                    strTmp = Replace(temp(i), "[", "")
                    myLrc(iLrcRows).lrcTime = GetTimeSec(strTmp)              'ת��������
                    myLrc(iLrcRows).lrcString = temp(j)
                    
                    If myLrc(iLrcRows).lrcTime Or Len(temp(j)) Then
                        iLrcRows = iLrcRows + 1
                    End If
                Next i
            Else
                ReDim Preserve myLrc(iLrcRows)
                
                myLrc(iLrcRows).lrcTime = 0
                myLrc(iLrcRows).lrcString = temp(j)
            End If
        End If
    End If
End Sub

'������
Private Sub SortLrc()
    Dim i As Integer, j As Integer
    Dim tmpLrc      As LRCROWINFO
    
    For i = 0 To iLrcRows - 2
        For j = iLrcRows - 1 To i + 1 Step -1
            If myLrc(j).lrcTime < myLrc(j - 1).lrcTime Then         '�����һ����ʱ��С�ڱ������򻥻�
                tmpLrc = myLrc(j - 1)
                myLrc(j - 1) = myLrc(j)
                myLrc(j) = tmpLrc
            End If
        Next j
    Next i
    
    '������һ�䲻�ǿ�����Ӹ�����
    If Len(myLrc(iLrcRows - 1).lrcString) Then
        ReDim Preserve myLrc(iLrcRows)
        myLrc(iLrcRows).lrcTime = myLrc(iLrcRows - 1).lrcTime * 2 - myLrc(iLrcRows - 2).lrcTime
    Else
        iLrcRows = iLrcRows - 1
    End If
End Sub


'�õ���ǰ��ʱ��.��ֵ��         '�� "01:04:13.55" ��ʽ���ִ�ת��Ϊ���� 3853.55
Private Function GetTimeSec(TempStr As String) As Single
    On Error Resume Next
    
    Dim sTime       As Single
    Dim temp()      As String
    Dim Value       As Single
    Dim i           As Integer
    temp = Split(TempStr, ":")
    sTime = 0
    For i = 0 To UBound(temp)
        Value = UBound(temp) - i
        sTime = sTime + temp(i) * (60 ^ Value)
    Next i
    GetTimeSec = sTime
End Function

'��ȡ�ļ�
Public Function ReadFile(LrcFileName As String)
    Dim iFree       As Integer
    Dim strTmp      As String
    
    Call ClearLrc
    
    If PathFileExists(LrcFileName) Then
        iFree = FreeFile
        Open LrcFileName For Input As #iFree
             While Not EOF(iFree)
                Line Input #iFree, strTmp
                Call SplitLrc(Trim(strTmp))
            Wend
        Close #iFree
        
        If iLrcRows > 0 Then
            Call SortLrc            '������ڸ����������
        End If
    End If
End Function

Public Function SeekLrc(sTime As Double, Optional bChange As Boolean = False) As Boolean
    On Error Resume Next
    
    Dim i As Integer, j As Integer
    
    If iLrcRows = 0 Then Exit Function
    
    If sTime < m_LastTime Then bChange = True                   'С���ϴβ���ʱ�䣬Ӧ���ǻص��Ĳ��Ž���
    m_LastTime = sTime                                          '��������Ĳ���ʱ��
    
    If sTime > myLrc(iLrcRows).lrcTime Then                     '����Ѿ���ʾ���
        iCurPlay = iLrcRows                                     '��λ�����һ����
    ElseIf sTime <= myLrc(0).lrcTime Then                       'С�ڵ�һ����ʱ�䣬���൱�ڻ�δ��ʼ
        iCurPlay = -1
        SeekLrc = True
        Exit Function
    Else
        '�жϵ�ǰ����Ƿ��� ���� �����˲���λ��
        If sTime > myLrc(iCurPlay + 1).lrcTime Or bChange Then
ReFind:
            j = IIf(bChange, 0, iCurPlay + 1)                   '����һ���ʿ�ʼ����(bChangeʱ��ǿ�ƴ�ͷ��ʼ����Ϊ�п���λ�õ������˵�ǰλ��֮ǰ)
            For i = j To iLrcRows
                If myLrc(i).lrcTime <= sTime And _
                            myLrc(i + 1).lrcTime > sTime Then
                    iCurPlay = i                                '�ҵ����
                    Exit For
                End If
            Next
            
            If i > iLrcRows Then                                        '��δ�ҵ���ǿ����ͷ����һ��
                If bChange Then                                         '����ͷ��δ�ҵ�����ǵ����һ��
                    iCurPlay = iLrcRows
                Else
                    bChange = True
                    j = 0
                    GoTo ReFind
                End If
            End If
        End If
    End If
    SeekLrc = True
End Function
Public Function GDIPlusInitialize() As Boolean
    Dim GpInput As GdiplusStartupInput
    Dim lToken As Long
    
    GpInput.GdiplusVersion = 1
    If GdiplusStartup(lToken, GpInput) = OK Then
       m_lngToken = lToken
       GDIPlusInitialize = True
    End If
End Function
Public Sub GDIPlusTerminate()
   If m_lngToken <> 0 Then
      Call GdiplusShutdown(m_lngToken)
      m_lngToken = 0
   End If
End Sub

'Twips to Pixels � ת��Ϊ ����
Function ConvertTwipsToPixels(lngTwips As Long, Optional lngDirection As Long = 0) As Long
    'lngDirection  0 ��ʾˮƽ�����㴹ֱ��
   'Handle to device
   Dim lngDC As Long
   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)
   
   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch
End Function

Function ConvertPointsToPixels(sPoints As Single, Optional lngDirection As Long = 0) As Long
   Dim lngDC As Long
   Dim lngPixelsPerInch As Long
   Const nPointsPerInch = 72        'ÿӢ�� 72 ��
   lngDC = GetDC(0)
   
   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
   End If
   lngDC = ReleaseDC(0, lngDC)
   ConvertPointsToPixels = (sPoints / nPointsPerInch) * lngPixelsPerInch
End Function


Public Function URLEncoding(ByVal vstrIn As String) As String
Dim strReturn As String, innerCode, Hight8, Low8
    strReturn = ""
    Dim i
    Dim thisChr
    
    For i = 1 To Len(vstrIn)
        
        thisChr = Mid(vstrIn, i, 1)
        
        If Abs(Asc(thisChr)) < &HFF Then
            If thisChr = " " Then
                strReturn = strReturn & "+"
            ElseIf InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_.", thisChr) > 0 Then
                strReturn = strReturn & thisChr
            Else
                strReturn = strReturn & "%" & IIf(Asc(thisChr) > 16, "", "0") & Hex(Asc(thisChr))
            End If
        Else
            innerCode = Asc(thisChr)
            If innerCode < 0 Then
                innerCode = innerCode + &H10000
            End If
            Hight8 = (innerCode And &HFF00) \ &HFF
            Low8 = innerCode And &HFF
            strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
        End If
    Next
    
    URLEncoding = strReturn
    
End Function
