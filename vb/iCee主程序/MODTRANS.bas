Attribute VB_Name = "桌面传输"
Option Explicit

Public Const CW_USEDEFAULT As Long = &H80000000

Public Const S_OK = 0

Public Const DefNum16 As Integer = &H8000
Public Const DefNum32 As Long = &H80000000
Public Const DefNum = DefNum32

Public Const ErrIdx As Long = -1

Private m_Inited As Boolean

Public BitPosMask(0 To 31) As Long
Public BitMapMask(0 To 31) As Long
Public BitsMask(0 To 32) As Long

Public Type Size
    cx As Long
    cy As Long
End Type


Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type POINTS
    X As Integer
    Y As Integer
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function CopyRect Lib "user32.dll" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function EqualRect Lib "user32.dll" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function IntersectRect Lib "user32.dll" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function IsRectEmpty Lib "user32.dll" (lpRect As RECT) As Long
Public Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PtInRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetRectEmpty Lib "user32.dll" (lpRect As RECT) As Long
Public Declare Function SubtractRect Lib "user32.dll" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Public Declare Function UnionRect Lib "user32.dll" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Public Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

Public Type RGNDATAHEADER
    dwSize As Long
    ITYPE As Long
    nCount As Long
    nRgnSize As Long
    rcBound As RECT
End Type
Public Const RDH_RECTANGLES As Long = 1
Public Type RGNDATA
    rdh As RGNDATAHEADER
    Buffer(0 To 0) As Byte
End Type


Public Const RGN_AND As Long = 1
Public Const RGN_OR As Long = 2
Public Const RGN_XOR As Long = 3
Public Const RGN_DIFF As Long = 4
Public Const RGN_COPY As Long = 5
Public Const RGN_MAX As Long = RGN_COPY
Public Const RGN_MIN As Long = RGN_AND

Public Const RGN_ERROR As Long = 0
Public Const NULLREGION As Long = 1
Public Const SIMPLEREGION As Long = 2
Public Const COMPLEXREGION As Long = 3

Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32.dll" (lpRect As RECT) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32.dll" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePolyPolygonRgn Lib "gdi32.dll" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (lpRect As RECT) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal y3 As Long) As Long
Public Declare Function EqualRgn Lib "gdi32.dll" (ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long) As Long
Public Declare Function ExtCreateRegion Lib "gdi32.dll" (lpXform As XFORM, ByVal nCount As Long, lpRgnData As Any) As Long
Public Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetPolyFillMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Public Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function InvertRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function PaintRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function RectInRegion Lib "gdi32.dll" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function SetPolyFillMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetRectRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long


'-- PolyFillMode
Public Const ALTERNATE As Long = 1
Public Const WINDING As Long = 2

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)


Public Declare Function IsBadCodePtr Lib "kernel32.dll" (ByVal lpfn As Long) As Long
Public Declare Function IsBadReadPtr Lib "kernel32.dll" (LP As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadWritePtr Lib "kernel32.dll" (LP As Any, ByVal ucb As Long) As Long

Public Declare Function IsBadStringPtr Lib "kernel32.dll" Alias "IsBadStringPtrA" (ByVal lpsz As String, ByVal ucchMax As Long) As Long
Public Declare Function IsBadStringPtrA Lib "kernel32.dll" (lpsz As Any, ByVal ucchMax As Long) As Long
Public Declare Function IsBadStringPtrW Lib "kernel32.dll" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long

Public Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcmp Lib "kernel32.dll" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function lstrcatA Lib "kernel32.dll" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrcmpA Lib "kernel32.dll" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrcmpiA Lib "kernel32.dll" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrcpyA Lib "kernel32.dll" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function lstrcpynA Lib "kernel32.dll" (lpString1 As Any, lpString2 As Any, ByVal iMaxLength As Long) As Long
Public Declare Function lstrlenA Lib "kernel32.dll" (lpString As Any) As Long
Public Declare Function lstrcatW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Public Declare Function lstrcmpW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Public Declare Function lstrcmpiW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Public Declare Function lstrcpyW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Public Declare Function lstrcpynW Lib "kernel32.dll" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Public Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long


Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

Public Type SAFEARRAY
    cDims As Integer         '这个数组有几维.
    fFeatures As Integer     '这个数组有什么特性.
    cbElements As Long       '数组的每个元素有多大.
    cLocks As Long           '这个数组被锁定过几次.
    pvData As Long           '这个数组里的数据放在什么地方.
End Type
Public Type SAFEARRAYBOUND
    cElements As Long      '这一维有多少个元素.
    lLbound As Long        '它的索引从几开始.
End Type
Public Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 0) As SAFEARRAYBOUND
End Type
Public Const FADF_AUTO         As Long = &H1
Public Const FADF_STATIC       As Long = &H2
Public Const FADF_EMBEDDED     As Long = &H4
Public Const FADF_FIXEDSIZE   As Long = &H10
Public Const FADF_RECORD      As Long = &H20
Public Const FADF_HAVEIID     As Long = &H40
Public Const FADF_HAVEVARTYPE As Long = &H80
Public Const FADF_BSTR       As Long = &H100
Public Const FADF_UNKNOWN    As Long = &H200
Public Const FADF_DISPATCH   As Long = &H400
Public Const FADF_VARIANT    As Long = &H800
Public Const FADF_RESERVED  As Long = &HF008

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Public Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(32) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(32) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type


'The following functions are used with system time.
Public Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Public Declare Function GetSystemTimeAdjustment Lib "kernel32.dll" (lpTimeAdjustment As Long, lpTimeIncrement As Long, lpTimeAdjustmentDisabled As Boolean) As Long
Public Declare Function LocalFileTimeToFileTime Lib "kernel32.dll" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Public Declare Function SetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SetSystemTimeAdjustment Lib "kernel32.dll" (ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Boolean) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long


'The following functions are used with local time.
Public Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Sub GetLocalTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Public Declare Function GetTimeZoneInformation Lib "kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Declare Function SetLocalTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SetTimeZoneInformation Lib "kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long


'The following functions are used with file time.
Public Declare Function CompareFileTime Lib "kernel32.dll" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Sub GetSystemTimeAsFileTime Lib "kernel32.dll" (ByRef lpSystemTimeAsFileTime As FILETIME)
Public Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long


'The following functions are used with MS-DOS date and time.
Public Declare Function DosDateTimeToFileTime Lib "kernel32.dll" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FILETIME) As Long
Public Declare Function FileTimeToDosDateTime Lib "kernel32.dll" (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long


'The following functions are used with Windows time.
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


'Timer(计时器)
Public Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Const VBLine_StepStart = &H1 '起始坐标是相对坐标
Public Const VBLine_UseColor = &H2 '使用Color参数
Public Const VBLine_UseStart = &H4 '使用起始坐标
Public Const VBLine_StepEnd = &H8 '结束坐标是相对坐标
Public Const VBLine_B = &H10 '线框
Public Const VBLine_BF = &H20 '填充举行

Public Const VBLine_FillRectAll = VBLine_UseColor Or VBLine_UseStart Or VBLine_BF

Private m_Frequency As Currency

Private m_IsBusy As Long
Private m_OldTime As Currency
Private m_BusyTime As Currency

Public Function MyAddressOf(ByVal FunPtr As Long)
    MyAddressOf = FunPtr
End Function

Public Function VbColor2RGB(ByVal Color As OLE_COLOR, Optional hPalette As Long) As Long
    Dim t As Long
    If OleTranslateColor(Color, hPalette, t) = 0 Then
        VbColor2RGB = t
    Else
        VbColor2RGB = ErrIdx
    End If
End Function

'单位:毫秒
Public Function GetCurTime() As Currency
    If m_Frequency = 0 Then '未初始化
        If QueryPerformanceFrequency(m_Frequency) = 0 Then
            m_Frequency = ErrIdx '无高精度计数器
        End If
    End If
    
    If m_Frequency <> ErrIdx Then
        Dim CurCount As Currency
        Call QueryPerformanceCounter(CurCount)
        GetCurTime = CurCount * 1000@ / m_Frequency
    Else
        GetCurTime = GetTickCount()
    End If
    
End Function

Public Function StartBusy() As Currency
    If m_IsBusy = 0 Then
        m_OldTime = GetCurTime()
        m_BusyTime = 0
        Screen.MousePointer = vbHourglass
    End If
    m_IsBusy = m_IsBusy + 1
    StartBusy = m_OldTime
End Function

Public Function EndBusy() As Currency
    If m_IsBusy > 0 Then
        m_IsBusy = m_IsBusy - 1
        If m_IsBusy = 0 Then
            m_BusyTime = GetCurTime() - m_OldTime
            Screen.MousePointer = vbDefault
        End If
    End If
    EndBusy = m_BusyTime
End Function

Public Property Get BusyTime() As Currency
    BusyTime = m_BusyTime
End Property
Public Function Hex2(ByVal Value As Byte) As String
    Hex2 = Right("0" & Hex(Value), 2)
End Function

Public Function Hex4(ByVal Value As Integer) As String
    Hex4 = Right(String(3, "0") & Hex(Value), 4)
End Function

Public Function Hex8(ByVal Value As Long) As String
    Hex8 = Right(String(7, "0") & Hex(Value), 8)
End Function
Public Function MakePoint(ByVal pArray As Long, _
        ByRef SA As SAFEARRAY1D, ByVal ItemSize As Long, _
        Optional ByVal lLbound As Long = 0, _
        Optional ByVal cElements As Long = &H7FFFFFFF) As Boolean
    If pArray = 0 Then Exit Function
    
    With SA
        .cDims = 1
        .fFeatures = 0
        .cbElements = ItemSize
        .cLocks = 0
        .pvData = 0
        .Bounds(0).lLbound = lLbound
        .Bounds(0).cElements = cElements
    End With
    CopyMemory ByVal pArray, VarPtr(SA), 4
    
    MakePoint = True
    
End Function

Public Function FreePoint(ByVal pArray As Long) As Boolean
    If pArray = 0 Then Exit Function
    
    CopyMemory ByVal pArray, 0&, 4
    
    FreePoint = True
    
End Function

Public Property Get Ptr(ByRef SA As SAFEARRAY1D) As Long
    Ptr = SA.pvData - SA.Bounds(0).lLbound * SA.cbElements
End Property
Public Property Let Ptr(ByRef SA As SAFEARRAY1D, ByVal RHS As Long)
    SA.pvData = RHS + SA.Bounds(0).lLbound * SA.cbElements
End Property

Public Function ChkFileRead(filename As String) As Boolean
    Dim hF As Integer
    hF = FreeFile(1)
    On Error Resume Next
    Open filename For Input Access Read Lock Write As hF
    If ERR.Number Then
        ChkFileRead = False
    Else
        Close hF
        ChkFileRead = True
    End If
End Function

Public Function ChkFileWrite(filename As String) As Boolean
    Dim hF As Integer
    hF = FreeFile(1)
    On Error Resume Next
    Open filename For Output As hF
    If ERR.Number Then
        ChkFileWrite = False
    Else
        Close hF
        ChkFileWrite = True
    End If
End Function


Public Property Get Inited() As Boolean
    Inited = m_Inited
End Property

Public Sub Init()
    If m_Inited Then Exit Sub
    m_Inited = True
    
    Dim i As Long
    
    For i = 0 To 30
        BitPosMask(i) = 2& ^ i
    Next i
    BitPosMask(31) = &H80000000
    
    For i = 0 To 7
        BitMapMask(i) = BitPosMask(7 - i)
    Next i
    For i = 8 To &HF
        BitMapMask(i) = BitPosMask(&HF - i + 8)
    Next i
    For i = &H10 To &H17
        BitMapMask(i) = BitPosMask(&H17 - i + &H10)
    Next i
    For i = &H18 To &H1F
        BitMapMask(i) = BitPosMask(&H1F - i + &H18)
    Next i
    
    For i = 0 To 30
        BitsMask(i) = 2& ^ i - 1
    Next i
    BitsMask(31) = &H7FFFFFFF
    BitsMask(32) = -1 '&HFFFFFFFF
    
End Sub
Public Property Get LoBit4(ByRef Data As Byte) As Byte
    LoBit4 = Data And &HF
End Property

Public Property Let LoBit4(ByRef Data As Byte, ByVal RHS As Byte)
    Data = (Data And &HF0) Or (RHS And &HF)
End Property

Public Property Get HiBit4(ByRef Data As Byte) As Byte
    HiBit4 = (Data And &HF0) \ &H10
End Property

Public Property Let HiBit4(ByRef Data As Byte, ByVal RHS As Byte)
    Data = (Data And &HF) Or ((RHS And &HF) * &H10)
End Property

Public Function MakeByte(ByVal hi As Long, ByVal Lo As Long) As Byte
    MakeByte = ((hi And &HF) * &H10) Or (Lo And &HF)
End Function
Public Property Get LoByte(ByRef Word As Integer) As Byte
    LoByte = Word And &HFF
End Property

Public Property Get HiByte(ByRef Word As Integer) As Byte
    HiByte = ((Word And &H7F00) \ &H100) Or (((Word And &H8000) <> 0) And &H80)
End Property

Public Property Let LoByte(ByRef Word As Integer, ByVal vData As Byte)
    Word = (Word And &HFF00) Or vData
End Property

Public Property Let HiByte(ByRef Word As Integer, ByVal vData As Byte)
    Word = (Word And &HFF) Or ((vData And &H7F) * &H100) Or (((vData And &H80) <> 0) And &H8000)
End Property

Public Function MakeWord(ByVal HiByte As Byte, ByVal LoByte As Byte) As Integer
    MakeWord = ((HiByte And &H7F) * &H100 Or (((HiByte And &H80) <> 0) And &H8000)) Or LoByte
End Function

'------------------------------------------------

Public Property Get ULoWord(ByRef DWord As Long) As Long
    ULoWord = DWord And &HFFFF&
End Property

Public Property Get UHiWord(ByRef DWord As Long) As Long
    UHiWord = ((DWord And &H7FFF0000) \ &H10000) Or (((DWord And &H80000000) <> 0) And &H8000&)
End Property

Public Property Let ULoWord(ByRef DWord As Long, ByVal vData As Long)
    DWord = (DWord And &HFFFF0000) Or (vData And &HFFFF)
End Property

Public Property Let UHiWord(ByRef DWord As Long, ByVal vData As Long)
    DWord = (DWord And &HFFFF&) Or ((vData And &H7FFF) * &H10000) Or (((vData And &H8000&) <> 0) And &H80000000)
End Property

Public Function UMakeDWord(ByVal HIWORD As Long, ByVal LOWORD As Long) As Long
    UMakeDWord = ((HIWORD And &H7FFF) * &H10000 Or (((HIWORD And &H8000&) <> 0) And &H80000000)) _
            Or (LOWORD And &HFFFF)
End Function

'------------------------------------------------

Public Property Get LOWORD(ByRef DWord As Long) As Integer
    LOWORD = (DWord And &H7FFF&) Or (((DWord And &H8000&) <> 0) And &H8000)
End Property

Public Property Get HIWORD(ByRef DWord As Long) As Integer
    HIWORD = ((DWord And &H7FFF0000) \ &H10000) Or (((DWord And &H80000000) <> 0) And &H8000)
End Property

Public Property Let LOWORD(ByRef DWord As Long, ByVal vData As Integer)
    DWord = (DWord And &HFFFF0000) Or (vData And &H7FFF) Or (((vData And &H8000) <> 0) And &H8000&)
End Property

Public Property Let HIWORD(ByRef DWord As Long, ByVal vData As Integer)
    DWord = (DWord And &HFFFF&) Or ((vData And &H7FFF) * &H10000) Or (((vData And &H8000) <> 0) And &H80000000)
End Property

Public Function MakeDWord(ByVal HIWORD As Integer, ByVal LOWORD As Integer) As Long
    MakeDWord = ((HIWORD And &H7FFF) * &H10000 Or (((HIWORD And &H8000) <> 0) And &H80000000)) _
            Or ((LOWORD And &H7FFF) Or (((LOWORD And &H8000) <> 0) And &H8000&))
End Function

Public Function MAKELPARAM(ByVal L As Integer, ByVal H As Integer) As Long
    MAKELPARAM = MakeDWord(H, L)
End Function


Public Function MAKELONG(ByVal wLow As Integer, ByVal wHigh As Integer) As Long
    MAKELONG = MakeDWord(wHigh, wLow)
End Function

'------------------------------------------------

Public Property Get ColorR(ByRef Color As Long) As Byte
    ColorR = Color And &HFF
End Property

Public Property Get ColorG(ByRef Color As Long) As Byte
    ColorG = (Color And &HFF00&) \ &H100&
End Property

Public Property Get ColorB(ByRef Color As Long) As Byte
    ColorB = (Color And &HFF0000) \ &H10000
End Property

Public Property Get ColorA(ByRef Color As Long) As Byte
    ColorA = ((Color And &H7F000000) \ &H1000000) Or (((Color And &H80000000) <> 0) And &H80)
End Property

Public Property Let ColorR(ByRef Color As Long, ByVal vData As Byte)
    Color = (Color And &HFFFFFF00) Or vData
End Property

Public Property Let ColorG(ByRef Color As Long, ByVal vData As Byte)
    Color = (Color And &HFFFF00FF) Or (vData * &H100&)
End Property

Public Property Let ColorB(ByRef Color As Long, ByVal vData As Byte)
    Color = (Color And &HFF00FFFF) Or (vData * &H10000)
End Property

Public Property Let ColorA(ByRef Color As Long, ByVal vData As Byte)
    Color = (Color And &HFFFFFF) Or ((vData And &H7F) * &H1000000) Or (((vData And &H80) <> 0) And &H80000000)
End Property

Public Function RGBA(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, ByVal Alpha As Byte) As Long
    RGBA = Red Or Green * &H100& Or Blue * &H10000 Or ((Alpha And &H7F) * &H1000000 Or (((Alpha And &H80) <> 0) And &H80000000))
End Function

'------------------------------------------------

'将大端方式的Word转为小端方式
Public Property Get WordBig(ByRef BigData As Integer) As Long
    WordBig = ((BigData And &H7F) * &H100&) _
            Or (((BigData And &H80) <> 0) And &H8000&) _
            Or ((BigData And &H7F00) \ &H100&) _
            Or (((BigData And &H8000) <> 0) And &H80&)
End Property
Public Property Let WordBig(ByRef BigData As Integer, ByVal RHS As Long)
    BigData = ((RHS And &H7F&) * &H100) _
            Or (((RHS And &H80&) <> 0) And &H8000) _
            Or ((RHS And &H7F00&) \ &H100) _
            Or (((RHS And &H8000&) <> 0) And &H80)
End Property



'------------------------------------------------

Public Function Bin(ByVal Data As Long, Optional ByVal Size As Long = -1) As String
    Dim Sign As Boolean
    Dim TempStr As String
    
    Sign = Data < 0
    Data = Data And &H7FFFFFFF
    While Data
        TempStr = (Data And 1) & TempStr
        Data = Data \ 2
    Wend
    If Len(TempStr) = 0 Then TempStr = "0"
    If Sign Then
        TempStr = "1" & String$(32 - Len(TempStr) - 1, "0") & TempStr
    End If
    
    If Size > Len(TempStr) Then TempStr = String$(Size - Len(TempStr), "0") & TempStr
    'Debug.Print TempStr
    
    Bin = TempStr
    
End Function


'------------------------------------------------

'检查数字占多少位
Public Function ChkNumBits(ByVal Value As Long) As Long
    If Value = &H80000000 Then ChkNumBits = 32: Exit Function
    If Value < 0 Then Value = Abs(Value)
    Dim i As Long
    For i = 0 To 31
        If Value <= BitsMask(i) Then Exit For
    Next i
    ChkNumBits = i
End Function

'检查数字占多少位，并根据正负翻转位（JPEG系数的规定）
Public Function ChkNumBitsAuto(ByRef Value As Long) As Long
    If Value = &H80000000 Then ChkNumBitsAuto = 32: Exit Function
    Dim Sign As Long '为了速度，Long比Boolean快
    Dim i As Long
    Sign = Value And &H80000000
    If Sign Then Value = Abs(Value)
    For i = 0 To 31
        If Value <= BitsMask(i) Then Exit For
    Next i
    If Sign Then Value = Value Xor BitsMask(i)
    ChkNumBitsAuto = i
End Function





