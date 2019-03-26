VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'Psd读写
Option Explicit
Private Const BI_RGB = 0&
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Private Const DIB_RGB_COLORS = 0

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(255) As RGBQUAD
End Type

Private Type PsdHeaderInfo '定义Psd的文件头结构
        PsSignature(0 To 3) As Byte
        PsVersion As Integer
        PsReserved(0 To 5) As Byte
        PsChannels As Integer
        PsHeight As Long
        PsWidth As Long
        PsChannels2 As Integer
        PsColorType As Integer
        PsPalletteLength As Long
        PsCompressionType As Integer
End Type

Private PsdBitInfo As BITMAPINFO
Private PsdFileHeaderInfo As PsdHeaderInfo
Private PsdWidth As Long, PsdHeight As Long
Private PsdPtr As Long
Private PsdDC As Long
Private PsdStream() As Byte
Private TempStream() As Byte
Private BmpStream() As Byte
Private PsdMap As Long
Private hOldMap As Long
Private BitCounts As Long, BitChannels As Long
Private PsdPallette(768) As Byte
Private RGBColor(255) As RGBQUAD

Public Property Get Width() As Long
Width = PsdWidth
End Property

Public Property Get Height() As Long
Height = PsdHeight
End Property

Public Property Get Scanline() As Long
Scanline = PsdWidth * 4
End Property

Public Property Get SizeImage() As Long
SizeImage = PsdWidth * PsdHeight * 4
End Property

Public Function SetToArray(Bits() As Byte) As Boolean '导出文件流到一个BGRABGRA格式的数组中
Dim j As Long, i As Long, fi As Long, RunLengths As Long
Dim Scanline As Long, UStream As Long
Dim TempByte As Long, TheByte As Long
Dim SizePallette As Long, DbSizePallette As Long, ReservedColor As Long
Dim SizePlane As Long, DbSizePlane As Long, TSizePlane As Long

ReDim Bits(SizeImage) As Byte

Select Case BitCounts
Case 1 '读取灰度图像
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    Bits(i + 2) = PsdStream(fi)
    Bits(i + 1) = PsdStream(fi)
    Bits(i) = PsdStream(fi)
    fi = fi + BitChannels
  Next i
Case 2 '读取索引图像
  SizePallette = PsdFileHeaderInfo.PsPalletteLength \ 3
  DbSizePallette = SizePallette * 2
  fi = 0
  For i = 0 To SizePallette - 1
    With RGBColor(i)
        .rgbBlue = PsdPallette(i)
        .rgbGreen = PsdPallette(i + SizePallette)
        .rgbRed = PsdPallette(i + DbSizePallette)
    End With
  Next i
  For i = 0 To SizeImage - 3 Step 4
    With RGBColor(PsdStream(fi))
      Bits(i + 2) = .rgbBlue
      Bits(i + 1) = .rgbGreen
      Bits(i) = .rgbRed
      fi = fi + BitChannels
    End With
  Next i
Case 3 '读取RGB-24位图象
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    Bits(i + 2) = PsdStream(fi)
    Bits(i + 1) = PsdStream(fi + SizePlane)
    Bits(i) = PsdStream(fi + DbSizePlane)
    fi = fi + BitChannels
  Next i
Case 4 '读取RGB-32位图像
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  TSizePlane = 3 * SizePlane
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    Bits(i) = PsdStream(fi + DbSizePlane)
    Bits(i + 1) = PsdStream(fi + SizePlane)
    Bits(i + 2) = PsdStream(fi)
    Bits(i + 3) = &HFF - PsdStream(fi + TSizePlane)
    fi = fi + BitChannels
  Next i
Case 5 '读取CMYK图像
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  TSizePlane = 3 * SizePlane
  fi = 0
'   下面是ACDSee之类的图像浏览软件对CMYK-->RGB的算法
  For i = 0 To SizeImage - 3 Step 4
    ReservedColor = PsdStream(fi + TSizePlane) - &HFF
    TempByte = PsdStream(fi + DbSizePlane) + ReservedColor
    If TempByte > &HFF Then
      Bits(i) = &HFF
    ElseIf TempByte < &H0 Then
        Bits(i) = &H0
      Else
        Bits(i) = TempByte
    End If
    TempByte = PsdStream(fi + SizePlane) + ReservedColor
    If TempByte > &HFF Then
      Bits(i + 1) = &HFF
    ElseIf TempByte < &H0 Then
        Bits(i + 1) = &H0
      Else
        Bits(i + 1) = TempByte
    End If
    TempByte = PsdStream(fi) + ReservedColor
    If TempByte > &HFF Then
      Bits(i + 2) = &HFF
    ElseIf TempByte < &H0 Then
        Bits(i + 2) = &H0
      Else
        Bits(i + 2) = TempByte
    End If
    fi = fi + BitChannels
  Next i
Case 6 '读取Lab图像
  Dim LColor As Double, aColor As Double, bColor As Double
  Dim xColor As Double, YYn3 As Double, yColor As Double, zColor As Double
  Dim dbR As Double, dbG As Double, dbB As Double
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    LColor = PsdStream(fi)
    aColor = PsdStream(fi + SizePlane) '- &H80
    bColor = PsdStream(fi + DbSizePlane) '- &H80
'L /= 2.55;
    LColor = LColor / 2.55
'a -= 128;
    aColor = aColor - 128
'b -= 128;
    bColor = bColor - 128
'
'YYn3 = (L + 16) / 116;
    YYn3 = (LColor + 16) / 116
'X = pow(YYn3 + a / 500, 3);
    xColor = (YYn3 + aColor / 500) ^ 3
'Y = pow(YYn3, 3);
    yColor = YYn3 ^ 3
'Z = pow(YYn3 - b / 200, 3);
    zColor = (YYn3 - bColor / 200) ^ 3
'
'R = 3.079978 * X - 1.537150 * Y - 0.542555 * Z;
    dbR = 3.079978 * xColor - 1.53715 * yColor - 0.542555 * zColor
'G = -0.921248 * X + 1.875992 * Y + 0.045225 * Z;
    dbG = -0.921248 * xColor + 1.875992 * yColor + 0.045225 * zColor
'B = 0.052891 * X - 0.204043 * Y + 1.150671 * Z;
    dbB = 0.052891 * xColor - 0.204043 * yColor + 1.150671 * zColor
'
'if ( R > 0.00304 )
    If dbR > 0.00304 Then dbR = 1.055 * dbR ^ 0.41667 - 0.055 Else dbR = 12.92 * dbR
'R = 1.055 * pow(R, 0.41667) - 0.055;
'Else
'R = 12.92 * R;
'if ( G > 0.00304 )
    If dbG > 0.00304 Then dbG = 1.055 * dbG ^ 0.41667 - 0.055 Else dbG = 12.92 * dbG
'G= 1.055 * pow(G, 0.41667) - 0.055;
'Else
'G = 12.92 * G;
'if ( B > 0.00304 )
    If dbB > 0.00304 Then dbB = 1.055 * dbB ^ 0.41667 - 0.055 Else dbB = 12.92 * dbB
'B = 1.055 * pow(B, 0.41667) - 0.055;
'Else
'B = 12.92 * B;
'
'R *= 255;
    TempByte = dbR * &HFF
    If TempByte > &HFF Then
      Bits(i + 2) = &HFF
    ElseIf TempByte < &H0 Then
        Bits(i + 2) = &H0
      Else
        Bits(i + 2) = TempByte
    End If
'G *= 255;
    TempByte = dbG * &HFF
    If TempByte > &HFF Then
      Bits(i + 1) = &HFF
    ElseIf TempByte < &H0 Then
        Bits(i + 1) = &H0
      Else
        Bits(i + 1) = TempByte
    End If
'B *= 255;
    TempByte = dbB * &HFF
    If TempByte > &HFF Then
      Bits(i) = &HFF
    ElseIf TempByte < &H0 Then
        Bits(i) = &H0
      Else
        Bits(i) = TempByte
    End If
    fi = fi + BitChannels
  Next i
End Select
SetToArray = True
End Function

Private Function CreatePsdDIB() As Boolean
'创建一个DIB，在Fantasia Photo使用此模块的时候删除了这个函数
With PsdBitInfo.bmiHeader
  .biSize = Len(PsdBitInfo.bmiHeader)
  .biBitCount = 32
  .biCompression = BI_RGB
  .biPlanes = 1
  .biWidth = PsdWidth
  .biHeight = PsdHeight
  .biSizeImage = Scanline * .biHeight
End With

PsdDC = CreateCompatibleDC(0)
PsdMap = CreateDIBSection(PsdDC, PsdBitInfo, DIB_RGB_COLORS, PsdPtr, 0, 0)

If PsdMap Then
  hOldMap = SelectObject(PsdDC, PsdMap)
  CreatePsdDIB = True
Else
  DeleteObject PsdDC
  CreatePsdDIB = False
  Exit Function
End If
End Function

Public Function LoadPsdFile(filename As String) As Boolean
On Error GoTo ERROR
Dim FileNum As Integer, FileLengths As Long
Dim j As Long, i As Long, fi As Long, fj As Long, RunLengths As Long
Dim Scanline As Long, UStream As Long
Dim TempByte As Long, TheByte As Long
Dim FileStream() As Byte, FilePtr As Long, FileOffSet As Long
Dim SizePallette As Long, DbSizePallette As Long, ReservedColor As Long
Dim SizePlane As Long, DbSizePlane As Long, TSizePlane As Long
Const Sqr16And6 As Long = &H1000000
Const Sqr16And4 As Long = &H10000
Const Sqr16And2 As Long = &H100
FileNum = FreeFile
FileLengths = FileLen(filename)
With PsdFileHeaderInfo
  Open filename For Binary As #FileNum
    ReDim FileStream(FileLengths)
    Get #FileNum, , FileStream
  Close #FileNum
  For i = 0 To 3
    .PsSignature(i) = FileStream(FilePtr)
    FilePtr = FilePtr + 1
  Next i
  .PsVersion = FileStream(FilePtr) * Sqr16And2 + FileStream(FilePtr + 1)
  FilePtr = FilePtr + 8
  .PsChannels = FileStream(FilePtr) * Sqr16And2 + FileStream(FilePtr + 1)
  FilePtr = FilePtr + 2
  .PsHeight = FileStream(FilePtr) * Sqr16And6 + FileStream(FilePtr + 1) * Sqr16And4 + FileStream(FilePtr + 2) * Sqr16And2 + FileStream(FilePtr + 3)
  PsdHeight = PsdFileHeaderInfo.PsHeight
  FilePtr = FilePtr + 4
  .PsWidth = FileStream(FilePtr) * Sqr16And6 + FileStream(FilePtr + 1) * Sqr16And4 + FileStream(FilePtr + 2) * Sqr16And2 + FileStream(FilePtr + 3)
  PsdWidth = PsdFileHeaderInfo.PsWidth
  FilePtr = FilePtr + 4
  .PsChannels2 = FileStream(FilePtr) * Sqr16And2 + FileStream(FilePtr + 1)
  FilePtr = FilePtr + 2
  .PsColorType = FileStream(FilePtr) * Sqr16And2 + FileStream(FilePtr + 1)
  FilePtr = FilePtr + 2
  FileOffSet = FileStream(FilePtr) * Sqr16And6 + FileStream(FilePtr + 1) * Sqr16And4 + FileStream(FilePtr + 2) * Sqr16And2 + FileStream(FilePtr + 3)
  .PsPalletteLength = FileOffSet
  FilePtr = FilePtr + 4
  If .PsColorType = 2 Then CopyMemory ByVal VarPtr(PsdPallette(0)), ByVal VarPtr(FileStream(FilePtr)), .PsPalletteLength
  FilePtr = FilePtr + FileOffSet
  FileOffSet = FileStream(FilePtr) * Sqr16And6 + FileStream(FilePtr + 1) * Sqr16And4 + FileStream(FilePtr + 2) * Sqr16And2 + FileStream(FilePtr + 3)
  FilePtr = FilePtr + 4 + FileOffSet
  FileOffSet = FileStream(FilePtr) * Sqr16And6 + FileStream(FilePtr + 1) * Sqr16And4 + FileStream(FilePtr + 2) * Sqr16And2 + FileStream(FilePtr + 3)
  FilePtr = FilePtr + 4 + FileOffSet
  .PsCompressionType = FileStream(FilePtr) * Sqr16And2 + FileStream(FilePtr + 1)
  FilePtr = FilePtr + 2
  Select Case .PsColorType
  Case 1 '灰度图像
    BitCounts = 1
  Case 2 '索引图像
    BitCounts = 2
  Case 3 'RGB图像
    Select Case .PsChannels
    Case 4
      BitCounts = 4
    Case Else
      BitCounts = 3
    End Select
  Case 4 'CMYK图像
    BitCounts = 5
  Case 9 'Lab图像
    BitCounts = 6
  End Select
  BitChannels = .PsChannels2 \ 8
End With
If PsdFileHeaderInfo.PsCompressionType = 1 Then FilePtr = FilePtr + PsdFileHeaderInfo.PsHeight * PsdFileHeaderInfo.PsChannels * 2
UStream = FileLengths - FilePtr
If PsdFileHeaderInfo.PsCompressionType = 0 Then
  ReDim PsdStream(UStream)
  CopyMemory ByVal VarPtr(PsdStream(0)), ByVal VarPtr(FileStream(FilePtr)), UStream
Else
  ReDim TempStream(UStream)
  CopyMemory ByVal VarPtr(TempStream(0)), ByVal VarPtr(FileStream(FilePtr)), UStream
End If
ReDim BmpStream(SizeImage) As Byte

If PsdFileHeaderInfo.PsCompressionType = 1 Then
  UStream = UBound(TempStream)
  ReDim PsdStream(0 To PsdWidth * 5 * PsdHeight)
  j = 0
  i = 0
  '解压的过程
  Do Until i >= UStream
    If (TempStream(i) And &H80) Then  '如果字节大于&H80
      RunLengths = &H100 - TempStream(i) '取得重复的个数
      i = i + 1
      TheByte = TempStream(i)
      For fi = 0 To RunLengths '重复那么多次
        PsdStream(j) = TheByte
        j = j + 1
      Next fi
      i = i + 1
    Else '否则就直接读入
      RunLengths = TempStream(i)
      i = i + 1
      For fi = 0 To RunLengths
        PsdStream(j) = TempStream(i)
        i = i + 1
        j = j + 1
      Next fi
    End If
  Loop
End If

Select Case BitCounts
Case 1 '读取灰度图像
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    BmpStream(i + 2) = PsdStream(fi)
    BmpStream(i + 1) = PsdStream(fi)
    BmpStream(i) = PsdStream(fi)
    fi = fi + BitChannels
  Next i
Case 2 '读取索引图像
  SizePallette = PsdFileHeaderInfo.PsPalletteLength \ 3
  DbSizePallette = SizePallette * 2
  fi = 0
  For i = 0 To SizePallette - 1
    With RGBColor(i)
        .rgbBlue = PsdPallette(i)
        .rgbGreen = PsdPallette(i + SizePallette)
        .rgbRed = PsdPallette(i + DbSizePallette)
    End With
  Next i
  For i = 0 To SizeImage - 3 Step 4
    With RGBColor(PsdStream(fi))
      BmpStream(i + 2) = .rgbBlue
      BmpStream(i + 1) = .rgbGreen
      BmpStream(i) = .rgbRed
      fi = fi + BitChannels
    End With
  Next i
Case 3 '读取RGB-24位图象
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    BmpStream(i + 2) = PsdStream(fi)
    BmpStream(i + 1) = PsdStream(fi + SizePlane)
    BmpStream(i) = PsdStream(fi + DbSizePlane)
    fi = fi + BitChannels
  Next i
Case 4 '读取RGB-32位图像
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  TSizePlane = 3 * SizePlane
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    BmpStream(i) = PsdStream(fi + DbSizePlane)
    BmpStream(i + 1) = PsdStream(fi + SizePlane)
    BmpStream(i + 2) = PsdStream(fi)
    BmpStream(i + 3) = &HFF - PsdStream(fi + TSizePlane)
    fi = fi + BitChannels
  Next i
Case 5 '读取CMYK图像
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  TSizePlane = 3 * SizePlane
  fi = 0
'   下面是ACDSee之类的图像浏览软件对CMYK-->RGB的算法
  For i = 0 To SizeImage - 3 Step 4
    ReservedColor = PsdStream(fi + TSizePlane) - &HFF
    TempByte = PsdStream(fi + DbSizePlane) + ReservedColor
    If TempByte > &HFF Then
      BmpStream(i) = &HFF
    ElseIf TempByte < &H0 Then
        BmpStream(i) = &H0
      Else
        BmpStream(i) = TempByte
    End If
    TempByte = PsdStream(fi + SizePlane) + ReservedColor
    If TempByte > &HFF Then
      BmpStream(i + 1) = &HFF
    ElseIf TempByte < &H0 Then
        BmpStream(i + 1) = &H0
      Else
        BmpStream(i + 1) = TempByte
    End If
    TempByte = PsdStream(fi) + ReservedColor
    If TempByte > &HFF Then
      BmpStream(i + 2) = &HFF
    ElseIf TempByte < &H0 Then
        BmpStream(i + 2) = &H0
      Else
        BmpStream(i + 2) = TempByte
    End If
    fi = fi + BitChannels
  Next i
'   下面是我找到的PhotoShop对CMYK-->RGB的算法
'    For i = 0 To SizeImage - 3 Step 4
'      ReservedColor = psdstream(fi + TSizePlane) - &HFF
'      BmpStream(i) = psdstream(fi + DbSizePlane) + ReservedColor
'      BmpStream(i + 1) = psdstream(fi + SizePlane) + ReservedColor
'      BmpStream(i + 2) = psdstream(fi) + ReservedColor
'      fi = fi + 1
'    Next i
'    RGB -> CMYK                         | CMYK -> RGB
'    Black=minimum(1-Red,1-Green,1-Blue) | Red=1-minimum(1,Cyan*(1-Black)+Black)
'    Cyan=(1-Red-Black)/(1-Black)        | Green=1-minimum(1,Magenta*(1-Black)+Black)
'    Magenta=(1-Green-Black)/(1-Black)   | Blue=1-minimum(1,Yellow*(1-Black)+Black)
'    Yellow=(1-Blue-Black)/(1-Black)
Case 6 '读取Lab图像
  Dim LColor As Double, aColor As Double, bColor As Double
  Dim xColor As Double, YYn3 As Double, yColor As Double, zColor As Double
  Dim dbR As Double, dbG As Double, dbB As Double
  SizePlane = PsdWidth * PsdHeight * BitChannels
  DbSizePlane = 2 * SizePlane
  fi = 0
  For i = 0 To SizeImage - 3 Step 4
    LColor = PsdStream(fi)
    aColor = PsdStream(fi + SizePlane) '- &H80
    bColor = PsdStream(fi + DbSizePlane) '- &H80
'L /= 2.55;
    LColor = LColor / 2.55
'a -= 128;
    aColor = aColor - 128
'b -= 128;
    bColor = bColor - 128
'
'YYn3 = (L + 16) / 116;
    YYn3 = (LColor + 16) / 116
'X = pow(YYn3 + a / 500, 3);
    xColor = (YYn3 + aColor / 500) ^ 3
'Y = pow(YYn3, 3);
    yColor = YYn3 ^ 3
'Z = pow(YYn3 - b / 200, 3);
    zColor = (YYn3 - bColor / 200) ^ 3
'
'R = 3.079978 * X - 1.537150 * Y - 0.542555 * Z;
    dbR = 3.079978 * xColor - 1.53715 * yColor - 0.542555 * zColor
'G = -0.921248 * X + 1.875992 * Y + 0.045225 * Z;
    dbG = -0.921248 * xColor + 1.875992 * yColor + 0.045225 * zColor
'B = 0.052891 * X - 0.204043 * Y + 1.150671 * Z;
    dbB = 0.052891 * xColor - 0.204043 * yColor + 1.150671 * zColor
'
'if ( R > 0.00304 )
    If dbR > 0.00304 Then dbR = 1.055 * dbR ^ 0.41667 - 0.055 Else dbR = 12.92 * dbR
'R = 1.055 * pow(R, 0.41667) - 0.055;
'Else
'R = 12.92 * R;
'if ( G > 0.00304 )
    If dbG > 0.00304 Then dbG = 1.055 * dbG ^ 0.41667 - 0.055 Else dbG = 12.92 * dbG
'G= 1.055 * pow(G, 0.41667) - 0.055;
'Else
'G = 12.92 * G;
'if ( B > 0.00304 )
    If dbB > 0.00304 Then dbB = 1.055 * dbB ^ 0.41667 - 0.055 Else dbB = 12.92 * dbB
'B = 1.055 * pow(B, 0.41667) - 0.055;
'Else
'B = 12.92 * B;
'
'R *= 255;
    TempByte = dbR * &HFF
    If TempByte > &HFF Then
      BmpStream(i + 2) = &HFF
    ElseIf TempByte < &H0 Then
        BmpStream(i + 2) = &H0
      Else
        BmpStream(i + 2) = TempByte
    End If
'G *= 255;
    TempByte = dbG * &HFF
    If TempByte > &HFF Then
      BmpStream(i + 1) = &HFF
    ElseIf TempByte < &H0 Then
        BmpStream(i + 1) = &H0
      Else
        BmpStream(i + 1) = TempByte
    End If
'B *= 255;
    TempByte = dbB * &HFF
    If TempByte > &HFF Then
      BmpStream(i) = &HFF
    ElseIf TempByte < &H0 Then
        BmpStream(i) = &H0
      Else
        BmpStream(i) = TempByte
    End If
    fi = fi + BitChannels
  Next i
End Select
CreatePsdDIB
SetBitmapBits PsdMap, SizeImage, BmpStream(0)
LoadPsdFile = True
ERROR:
End Function

'关键的注释写完了，懒得写其它的了，应该看得懂
Private Function ChkFileWrite(filename As String) As Boolean
Dim FileNum As Integer
FileNum = FreeFile
On Error Resume Next
Open filename For Output As #FileNum
If ERR.Number Then
Else
  Close #FileNum
  ChkFileWrite = True
End If
On Error GoTo 0
End Function

Public Function DrawToDC(DestDC As Long, X As Long, Y As Long) As Long
DrawToDC = BitBlt(DestDC, X, Y, Width, Height, PsdDC, 0, 0, vbSrcCopy)
End Function

Public Function FreePsd() As Boolean
On Error Resume Next
If hOldMap Then DeleteObject SelectObject(PsdDC, hOldMap)
DeleteObject PsdDC
With PsdBitInfo.bmiHeader
.biBitCount = 0
.biHeight = 0
.biSizeImage = 0
.biWidth = 0
End With
FreePsd = True
End Function

Private Sub Class_Terminate()
FreePsd
End Sub

Public Function ksCheckBounds(ByVal Color As Long) As Byte
If Color > 255 Then
    ksCheckBounds = 255
ElseIf Color >= 0 Then
    ksCheckBounds = Color
End If
End Function


