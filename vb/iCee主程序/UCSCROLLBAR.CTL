VERSION 5.00
Begin VB.UserControl ucScrollbar 
   Alignable       =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H8000000F&
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
End
Attribute VB_Name = "ucScrollbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                                           'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                                          'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE                          'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                         'Subclass data type
    hwnd                   As Long                                            'Handle of the window being subclassed
    nAddrSub               As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig              As Long                                            'The address of the pre-existing WndProc
    nMsgCntA               As Long                                            'Msg after table entry count
    nMsgCntB               As Long                                            'Msg before table entry count
    aMsgTblA()             As Long                                            'Msg after table array
    aMsgTblB()             As Long                                            'Msg Before table array
End Type

Private sc_aSubData()      As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED   As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC  As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04     As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05     As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08     As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09     As Long = 137                                      'Table A (after) entry count patch offset

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const SM_CXVSCROLL  As Long = 2
Private Const SM_CYHSCROLL  As Long = 3
Private Const SM_CYVSCROLL  As Long = 20
Private Const SM_CXHSCROLL  As Long = 21
Private Const SM_SWAPBUTTON As Long = 23

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SPI_GETKEYBOARDDELAY As Long = 22
Private Const SPI_GETKEYBOARDPREF  As Long = 68

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const DFC_SCROLL          As Long = 3
Private Const DFCS_SCROLLUP       As Long = &H0
Private Const DFCS_SCROLLDOWN     As Long = &H1
Private Const DFCS_SCROLLLEFT     As Long = &H2
Private Const DFCS_SCROLLRIGHT    As Long = &H3
Private Const DFCS_INACTIVE       As Long = &H100
Private Const DFCS_PUSHED         As Long = &H200
Private Const DFCS_FLAT           As Long = &H4000
Private Const DFCS_MONO           As Long = &H8000

Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const BDR_RAISED As Long = &H5
Private Const BF_RECT    As Long = &HF

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const COLOR_BTNFACE     As Long = 15
Private Const COLOR_3DSHADOW    As Long = 16
Private Const COLOR_BTNTEXT     As Long = 18
Private Const COLOR_3DHIGHLIGHT As Long = 20
Private Const COLOR_3DDKSHADOW  As Long = 21

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Const WHITE_BRUSH As Long = 0
Private Const BLACK_BRUSH As Long = 4

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
    
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
 
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Type PAINTSTRUCT
    hdc             As Long
    fErase          As Long
    rcPaint         As RECT
    fRestore        As Long
    fIncUpdate      As Long
    rgbReserved(32) As Byte
End Type
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long



Private Const WM_SIZE           As Long = &H5
Private Const WM_PAINT          As Long = &HF
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_CANCELMODE     As Long = &H1F
Private Const WM_TIMER          As Long = &H113
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_LBUTTONDBLCLK  As Long = &H203
Private Const WM_THEMECHANGED   As Long = &H31A

Private Const MK_LBUTTON        As Long = &H1


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "uxtheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long


Private Const SB_THEME As String = "Scrollbar"

Private Const SBP_ARROWBTN = 1
Private Const SBP_THUMBBTNHORZ = 2
Private Const SBP_THUMBBTNVERT = 3
Private Const SBP_LOWERTRACKHORZ = 4
Private Const SBP_UPPERTRACKHORZ = 5
Private Const SBP_LOWERTRACKVERT = 6
Private Const SBP_UPPERTRACKVERT = 7
Private Const SBP_GRIPPERHORZ = 8
Private Const SBP_GRIPPERVERT = 9
Private Const SBP_SIZEBOX = 10

Private Const ABS_UPNORMAL = 1
Private Const ABS_UPHOT = 2
Private Const ABS_UPPRESSED = 3
Private Const ABS_UPDISABLED = 4
Private Const ABS_DOWNNORMAL = 5
Private Const ABS_DOWNHOT = 6
Private Const ABS_DOWNPRESSED = 7
Private Const ABS_DOWNDISABLED = 8
Private Const ABS_LEFTNORMAL = 9
Private Const ABS_LEFTHOT = 10
Private Const ABS_LEFTPRESSED = 11
Private Const ABS_LEFTDISABLED = 12
Private Const ABS_RIGHTNORMAL = 13
Private Const ABS_RIGHTHOT = 14
Private Const ABS_RIGHTPRESSED = 15
Private Const ABS_RIGHTDISABLED = 16

Private Const HSS_NORMAL = 1
Private Const HSS_HOT = 2
Private Const HSS_PUSHED = 3
Private Const HSS_DISABLED = 4

Private Const HTS_NORMAL = 1
Private Const HTS_HOT = 2
Private Const HTS_PUSHED = 3
Private Const HTS_DISABLED = 4

Private Const VSS_NORMAL = 1
Private Const VSS_HOT = 2
Private Const VSS_PUSHED = 3
Private Const VSS_DISABLED = 4

Private Const VTS_NORMAL = 1
Private Const VTS_HOT = 2
Private Const VTS_PUSHED = 3
Private Const VTS_DISABLED = 4
Public Enum sbOrientationCts
    [oVertical] = 0
    [oHorizontal] = 1
End Enum

Public Enum sbStyleCts
    [sClassic] = 0
    [sFlat] = 1
    [sThemed] = 2
    [sCustomDraw] = 3
End Enum

Public Enum sbOnPaintPartCts
    [ppTLButton] = 0
    [ppBRButton] = 1
    [ppTLTrack] = 2
    [ppBRTrack] = 3
    [ppNullTrack] = 4
    [ppThumb] = 5
End Enum

Public Enum sbOnPaintPartStateCts
    [ppsNormal] = 0
    [ppsPressed] = 1
    [ppsHot] = 2
    [ppsDisabled] = 3
End Enum

Private Enum eFlatButtonStateCts
    [fbsNormal] = 0
    [fbsSelected] = 1
    [fbsHot] = 2
End Enum

Private Const SBack = 0
Private Const SThumb = 1
Private Const STGrip = 2
Private Const SArrow = 3
Private Const SArrowTL = 4
Private Const SArrowBR = 5

Private Const HT_NOTHING          As Long = 0
Private Const HT_TLBUTTON         As Long = 1
Private Const HT_BRBUTTON         As Long = 2
Private Const HT_TLTRACK          As Long = 3
Private Const HT_BRTRACK          As Long = 4
Private Const HT_THUMB            As Long = 5

Private Const TIMERID_CHANGE1     As Long = 1
Private Const TIMERID_CHANGE2     As Long = 2
Private Const TIMERID_HOT         As Long = 3
Private Const TIMERID_HOT2        As Long = 4

Private Const CHANGEDELAY_MIN     As Long = 0
Private Const CHANGEFREQUENCY_MIN As Long = 25
Private Const TIMERDT_HOT         As Long = 25

Private Const THUMBSIZE_MIN       As Long = 8
Private Const GRIPPERSIZE_MIN     As Long = 16

Private BarPics(5)                As New pcMemDC
Private sDrawing                  As New cScrollbarDraw

Private m_bHasTrack               As Boolean
Private m_bHasNullTrack           As Boolean
Private m_uRctNullTrack           As RECT

Private m_uRctTLButton            As RECT
Private m_uRctBRButton            As RECT
Private m_uRctTLTrack             As RECT
Private m_uRctBRTrack             As RECT
Private m_uRctThumb               As RECT
Private m_lThumbOffset            As Long
Private m_uRctDrag                As RECT

Private m_bTLButtonPressed        As Boolean
Private m_bBRButtonPressed        As Boolean
Private m_bTLTrackPressed         As Boolean
Private m_bBRTrackPressed         As Boolean
Private m_bThumbPressed           As Boolean

Private m_bTLButtonHot            As Boolean
Private m_bBRButtonHot            As Boolean
Private m_bThumbHot               As Boolean

Private m_mOverUC                 As Boolean

Private m_lAbsRange               As Long
Private m_lThumbPos               As Long
Private m_lThumbSize              As Long
Private m_eHitTest                As Long
Private m_eHitTestHot             As Long
Private m_X                       As Long
Private m_Y                       As Long
Private m_lValueStartDrag         As Long

Private m_hPatternBrush           As Long

Private m_lChangeDelay            As Long
Private m_lChangeFrequency        As Long
Private m_lMax                    As Long
Private m_lMin                    As Long
Private m_lValue                  As Long
Private m_lSmallChange            As Long
Private m_lLargeChange            As Long
Private m_eOrientation            As sbOrientationCts
Private m_bShowButtons            As Boolean

Private Const ENABLED_DEF         As Boolean = True
Private Const MIN_DEF             As Long = 0
Private Const MAX_DEF             As Long = 100
Private Const VALUE_DEF           As Long = MIN_DEF
Private Const SMALLCHANGE_DEF     As Long = 1
Private Const LARGECHANGE_DEF     As Long = 10
Private Const CHANGEDELAY_DEF     As Long = 500
Private Const CHANGEFREQUENCY_DEF As Long = 50
Private Const ORIENTATION_DEF     As Long = [oVertical]
Private Const SHOWBUTTONS_DEF     As Boolean = True

Public Event Change()
Attribute Change.VB_MemberFlags = "200"
Public Event Scroll()
Public Event Show()
Public Event Hide()

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub


Private Sub UserControl_Initialize()
    Call pvCreatePatternBrush
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

Private Sub UserControl_Terminate()
    
    On Error GoTo Catch

    Call Subclass_StopAll
    
Catch:
    On Error GoTo 0
  
    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
    
    '-- Clean up
    Call DeleteObject(m_hPatternBrush)
End Sub

Private Sub UserControl_Resize()
    If (Ambient.UserMode = False) Then
        Call pvOnSize
    End If
End Sub

Private Sub UserControl_Paint()
    If (Ambient.UserMode = False) Then
        Call pvOnPaint(UserControl.hdc)
    End If
End Sub

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, _
                          ByRef bHandled As Boolean, _
                          ByRef lReturn As Long, _
                          ByRef lhwnd As Long, _
                          ByRef uMsg As Long, _
                          ByRef wParam As Long, _
                          ByRef lParam As Long _
                          )
Attribute zSubclass_Proc.VB_MemberFlags = "40"
                          
  Dim uPS As PAINTSTRUCT
  
    Select Case lhwnd
        
        Case UserControl.hwnd
        
            Select Case uMsg
            
                Case WM_PAINT
                    Call BeginPaint(lhwnd, uPS)
                    Call pvOnPaint(uPS.hdc)
                    Call EndPaint(lhwnd, uPS)
                    bHandled = True: lReturn = 0
                    
                Case WM_SIZE
                    Call pvOnSize
                    bHandled = True: lReturn = 0
                
                Case WM_LBUTTONDOWN
                    Call pvOnMouseDown(wParam, lParam)
                    
                Case WM_MOUSEMOVE
                    Call pvOnMouseMove(wParam, lParam)
                
                Case WM_LBUTTONUP, WM_CANCELMODE
                    Call pvOnMouseUp
                   
                Case WM_LBUTTONDBLCLK
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    
                Case WM_TIMER
                    Call pvOnTimer(wParam)
            End Select
    End Select
End Sub

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub


Private Sub pvOnSize()
 
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

Private Sub pvOnPaint( _
            ByVal lhDC As Long _
            )
            
  Dim lfHorz As Long, Idx As Integer
  
    lfHorz = -CLng(m_eOrientation = [oHorizontal])

    If (UserControl.Enabled) Then
    
  
        If (m_bTLButtonHot) Then
            Idx = 2
          Else
            If (m_bTLButtonPressed) Then
                Idx = 3
              Else
                Idx = Abs(m_mOverUC)
            End If
        End If
        sDrawing.DrawTopLeftButton lhDC, Val(m_eOrientation), Idx, m_uRctTLButton
        
        If (m_bBRButtonHot) Then
            Idx = 2
          Else
            If (m_bBRButtonPressed) Then
                Idx = 3
              Else
                Idx = Abs(m_mOverUC)
            End If
        End If
        sDrawing.DrawBottomRightButton lhDC, Val(m_eOrientation), Idx, m_uRctBRButton
        

        If (m_bHasTrack) Then

            Idx = Abs(m_bTLTrackPressed)
            sDrawing.DrawBack lhDC, Val(m_eOrientation), Idx, m_uRctTLTrack

            Idx = Abs(m_bBRTrackPressed)
            sDrawing.DrawBack lhDC, Val(m_eOrientation), Idx, m_uRctBRTrack

            If (m_bThumbHot) Then
                Idx = 1
              Else
                If (m_bThumbPressed) Then
                    Idx = 2
                  Else
                    Idx = 0
                End If
            End If
            sDrawing.DrawBar lhDC, Val(m_eOrientation), Idx, m_uRctThumb
        End If
        If (m_bHasNullTrack) Then
            Call FillRect(lhDC, m_uRctNullTrack, m_hPatternBrush)
        End If
        
      
      Else

        sDrawing.DrawTopLeftButton lhDC, Val(m_eOrientation), 4, m_uRctTLButton
        sDrawing.DrawBottomRightButton lhDC, Val(m_eOrientation), 4, m_uRctBRButton
        If (m_bHasTrack) Then
            sDrawing.DrawBack lhDC, Val(m_eOrientation), 0, m_uRctTLTrack
            sDrawing.DrawBack lhDC, Val(m_eOrientation), 0, m_uRctBRTrack
            sDrawing.DrawBar lhDC, Val(m_eOrientation), 3, m_uRctThumb
        End If
        If (m_bHasNullTrack) Then
            Call FillRect(lhDC, m_uRctNullTrack, m_hPatternBrush)
        End If
    End If
End Sub

Private Sub pvOnMouseDown( _
            ByVal wParam As Long, _
            ByVal lParam As Long _
            )
  
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        
        Call pvMakePoints(lParam, m_X, m_Y)
        m_eHitTest = pvHitTest(m_X, m_Y)
        
        Select Case m_eHitTest
        
            Case HT_THUMB
                Select Case m_eOrientation
                    Case [oVertical]
                        m_lThumbOffset = m_uRctThumb.Top - m_Y
                    Case [oHorizontal]
                        m_lThumbOffset = m_uRctThumb.Left - m_X
                End Select
                m_bThumbPressed = True
                m_bThumbHot = False
                m_lValueStartDrag = m_lValue
                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                
            Case HT_TLBUTTON
                m_bTLButtonPressed = True
                m_bTLButtonHot = False
                Call pvScrollPosDec(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            
            Case HT_BRBUTTON
                m_bBRButtonPressed = True
                m_bBRButtonHot = False
                Call pvScrollPosInc(m_lSmallChange, True)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            
            Case HT_TLTRACK
                m_bTLTrackPressed = True
                Call pvScrollPosDec(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
            
            Case HT_BRTRACK
                m_bBRTrackPressed = True
                Call pvScrollPosInc(m_lLargeChange)
                Call pvKillTimer(TIMERID_CHANGE1)
                Call pvSetTimer(TIMERID_CHANGE1, m_lChangeDelay)
        End Select
    End If
End Sub

Private Sub pvOnMouseMove( _
            ByVal wParam As Long, _
            ByVal lParam As Long _
            )
  
  Dim lValuePrev As Long
  Dim lThumbPosPrev As Long
  Dim bPressed As Boolean
  Dim bHot As Boolean
        
    Call pvMakePoints(lParam, m_X, m_Y)
                    
    If (wParam And MK_LBUTTON = MK_LBUTTON) Then
        
        Select Case m_eHitTest
        
            Case HT_THUMB
            
                lValuePrev = m_lValue
                lThumbPosPrev = m_lThumbPos
                
                If (PtInRect(m_uRctDrag, m_X, m_Y)) Then
                    
                    Select Case m_eOrientation
                        
                        Case [oVertical]
                        
                            m_lThumbPos = m_Y + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.Bottom) Then
                                m_lThumbPos = m_uRctTLButton.Bottom
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.Top) Then
                                m_lThumbPos = m_uRctBRButton.Top - m_lThumbSize
                            End If
                        
                        Case [oHorizontal]
                        
                            m_lThumbPos = m_X + m_lThumbOffset
                            If (m_lThumbPos < m_uRctTLButton.Right) Then
                                m_lThumbPos = m_uRctTLButton.Right
                            End If
                            If (m_lThumbPos + m_lThumbSize > m_uRctBRButton.Left) Then
                                m_lThumbPos = m_uRctBRButton.Left - m_lThumbSize
                            End If
                    End Select
                    m_lValue = pvGetScrollPos()
                  
                  Else
                    
                    m_lValue = m_lValueStartDrag
                    m_lThumbPos = pvGetThumbPos()
                End If
                
                If (m_lThumbPos <> lThumbPosPrev) Then
                    Call pvSizeTrack
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    If (m_lValue <> lValuePrev) Then
                        RaiseEvent Scroll
                    End If
                End If
            
            Case HT_TLBUTTON
                
                bPressed = (PtInRect(m_uRctTLButton, m_X, m_Y) <> 0)
                If (bPressed Xor m_bTLButtonPressed) Then
                    m_bTLButtonPressed = bPressed
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                End If
                
            Case HT_BRBUTTON
                
                bPressed = (PtInRect(m_uRctBRButton, m_X, m_Y) <> 0)
                If (bPressed Xor m_bBRButtonPressed) Then
                    m_bBRButtonPressed = bPressed
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                End If
        End Select
    
      Else
        
        m_eHitTestHot = pvHitTest(m_X, m_Y)
        
        If Not m_mOverUC Then
            m_mOverUC = True
            Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
            Call pvKillTimer(TIMERID_HOT2)
            Call pvSetTimer(TIMERID_HOT2, TIMERDT_HOT)
        End If
        
        Select Case m_eHitTestHot
            
            Case HT_TLBUTTON
                bHot = (PtInRect(m_uRctTLButton, m_X, m_Y) <> 0)
                If (m_bTLButtonHot Xor bHot) Then
                    m_bTLButtonHot = True
                    m_bBRButtonHot = False
                    m_bThumbHot = False
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            
            Case HT_BRBUTTON
                bHot = (PtInRect(m_uRctBRButton, m_X, m_Y) <> 0)
                If (m_bBRButtonHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = True
                    m_bThumbHot = False
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
            
            Case HT_THUMB
                
                bHot = (PtInRect(m_uRctThumb, m_X, m_Y) <> 0)
                If (m_bThumbHot Xor bHot) Then
                    m_bTLButtonHot = False
                    m_bBRButtonHot = False
                    m_bThumbHot = True
                    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    Call pvKillTimer(TIMERID_HOT)
                    Call pvSetTimer(TIMERID_HOT, TIMERDT_HOT)
                End If
        End Select
    End If
End Sub

Private Sub pvOnMouseUp()

    Call pvKillTimer(TIMERID_HOT)
    Call pvKillTimer(TIMERID_CHANGE1)
    Call pvKillTimer(TIMERID_CHANGE2)
    
    If (m_eHitTest = HT_THUMB) Then
        If (m_lValue <> m_lValueStartDrag) Then
            RaiseEvent Change
        End If
    End If
    m_eHitTest = HT_NOTHING
    
    m_bTLButtonPressed = False
    m_bBRButtonPressed = False
    m_bThumbPressed = False
    m_bTLTrackPressed = False
    m_bBRTrackPressed = False
    
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Sub

Private Sub pvOnTimer(ByVal wParam As Long)
  
  Dim uPt As POINTAPI
  
    Select Case wParam
    
        Case TIMERID_CHANGE1
        
            Call pvKillTimer(TIMERID_CHANGE1)
            Call pvSetTimer(TIMERID_CHANGE2, m_lChangeFrequency)
       
        Case TIMERID_CHANGE2
        
            Select Case m_eHitTest
                
                Case HT_TLBUTTON
                    If (PtInRect(m_uRctTLButton, m_X, m_Y)) Then
                        If (pvScrollPosDec(m_lSmallChange) = False) Then
                            Call pvKillTimer(TIMERID_CHANGE2)
                        End If
                    End If
                
                Case HT_BRBUTTON
                    If (PtInRect(m_uRctBRButton, m_X, m_Y)) Then
                        If (pvScrollPosInc(m_lSmallChange) = False) Then
                            Call pvKillTimer(TIMERID_CHANGE2)
                        End If
                    End If
                    
                Case HT_TLTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos > m_Y) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                              Else
                                m_bTLTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos > m_X) Then
                                m_bTLTrackPressed = True
                                Call pvScrollPosDec(m_lLargeChange)
                              Else
                                m_bTLTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                    End Select
                
                Case HT_BRTRACK
                    Select Case m_eOrientation
                        Case [oVertical]
                            If (m_lThumbPos + m_lThumbSize < m_Y) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                              Else
                                m_bBRTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                        Case [oHorizontal]
                            If (m_lThumbPos + m_lThumbSize < m_X) Then
                                m_bBRTrackPressed = True
                                Call pvScrollPosInc(m_lLargeChange)
                              Else
                                m_bBRTrackPressed = False
                                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                            End If
                    End Select
           End Select
      
        Case TIMERID_HOT
            
            Call GetCursorPos(uPt)
            Call ScreenToClient(hwnd, uPt)
            
            Select Case True
                
                Case m_bTLButtonHot
                    If (PtInRect(m_uRctTLButton, uPt.X, uPt.Y) = 0) Then
                        m_bTLButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    End If
               
                Case m_bBRButtonHot
                    If (PtInRect(m_uRctBRButton, uPt.X, uPt.Y) = 0) Then
                        m_bBRButtonHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    End If
               
                Case m_bThumbHot
                    If (PtInRect(m_uRctThumb, uPt.X, uPt.Y) = 0) Then
                        m_bThumbHot = False
                        Call pvKillTimer(TIMERID_HOT)
                        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
                    End If
            End Select
            
        Case TIMERID_HOT2
            
            Call GetCursorPos(uPt)
            If WindowFromPoint(uPt.X, uPt.Y) <> UserControl.hwnd And m_mOverUC = True And m_eHitTest = 0 Then
                m_mOverUC = False
                Call pvKillTimer(TIMERID_HOT2)
                Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
            End If
        
    End Select
End Sub

'========================================================================================
' Private
'========================================================================================

'----------------------------------------------------------------------------------------
' Sizing
'----------------------------------------------------------------------------------------

Private Sub pvSizeButtons()
 
 Dim uRct        As RECT
 Dim lButtonSize As Long
    
    Call GetClientRect(hwnd, uRct)
    m_bHasTrack = False
    m_bHasNullTrack = False
    
    Select Case m_eOrientation
        
        Case [oVertical]
        
            '-- Size buttons
            lButtonSize = 18 * -CLng(m_bShowButtons) 'GetSystemMetrics(SM_CYVSCROLL) * -CLng(m_bShowButtons)
            With uRct
                If (2 * lButtonSize + THUMBSIZE_MIN > .Bottom) Then
                    If (2 * lButtonSize < .Bottom) Then
                        Call SetRect(m_uRctTLButton, 0, 0, .Right, lButtonSize)
                        Call SetRect(m_uRctBRButton, 0, .Bottom - lButtonSize, .Right, .Bottom)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, 0, lButtonSize, .Right, .Bottom - lButtonSize)
                      Else
                        Call SetRect(m_uRctTLButton, 0, 0, .Right, .Bottom \ 2)
                        Call SetRect(m_uRctBRButton, 0, .Bottom \ 2 + (.Bottom Mod 2), .Right, .Bottom)
                        m_bHasNullTrack = CBool(.Bottom Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, 0, .Bottom \ 2, .Right, .Bottom \ 2 + 1)
                        End If
                    End If
                  Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, .Right, lButtonSize)
                    Call SetRect(m_uRctBRButton, 0, .Bottom - lButtonSize, .Right, .Bottom)
                End If
            End With
            
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 250, 25)
            
        Case [oHorizontal]
            
            '-- Size buttons
            lButtonSize = 18 * -CLng(m_bShowButtons) 'GetSystemMetrics(SM_CXHSCROLL) * -CLng(m_bShowButtons)
            With uRct
                If (2 * lButtonSize + THUMBSIZE_MIN > .Right) Then
                    If (2 * lButtonSize < .Right) Then
                        Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Bottom)
                        Call SetRect(m_uRctBRButton, .Right - lButtonSize, 0, .Right, .Bottom)
                        m_bHasNullTrack = True
                        Call SetRect(m_uRctNullTrack, lButtonSize, 0, .Right - lButtonSize, .Bottom)
                      Else
                        Call SetRect(m_uRctTLButton, 0, 0, .Right \ 2, .Bottom)
                        Call SetRect(m_uRctBRButton, .Right \ 2 + (.Right Mod 2), 0, .Right, .Bottom)
                        m_bHasNullTrack = CBool(.Right Mod 2)
                        If (m_bHasNullTrack) Then
                            Call SetRect(m_uRctNullTrack, .Right \ 2, 0, .Right \ 2 + 1, .Bottom)
                        End If
                    End If
                  Else
                    m_bHasTrack = True
                    Call SetRect(m_uRctTLButton, 0, 0, lButtonSize, .Bottom)
                    Call SetRect(m_uRctBRButton, .Right - lButtonSize, 0, .Right, .Bottom)
                End If
            End With
            
            '-- Get max. drag area
            Call CopyRect(m_uRctDrag, uRct)
            Call InflateRect(m_uRctDrag, 25, 250)
    End Select
    
    '-- No track: avoid pvSizeTrack() calcs.
    If (m_bHasTrack = False) Then
        Call SetRectEmpty(m_uRctTLTrack)
        Call SetRectEmpty(m_uRctBRTrack)
        Call SetRectEmpty(m_uRctThumb)
    End If
End Sub

Private Sub pvSizeTrack()
 
    If (m_bHasTrack) Then
    
        '-- Tracks and thumbs exist
        Select Case m_eOrientation
            
            Case [oVertical]
                
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, 0, m_uRctTLButton.Bottom, m_uRctTLButton.Right, m_lThumbPos)
                Call SetRect(m_uRctBRTrack, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.Right, m_uRctBRButton.Top)
                Call SetRect(m_uRctThumb, 0, m_lThumbPos, m_uRctBRButton.Right, m_lThumbPos + m_lThumbSize)
                
            Case [oHorizontal]
            
                '-- Size both track parts and thumb
                Call SetRect(m_uRctTLTrack, m_uRctTLButton.Right, 0, m_lThumbPos, m_uRctTLButton.Bottom)
                Call SetRect(m_uRctBRTrack, m_lThumbPos + m_lThumbSize, 0, m_uRctBRButton.Left, m_uRctBRButton.Bottom)
                Call SetRect(m_uRctThumb, m_lThumbPos, 0, m_lThumbPos + m_lThumbSize, m_uRctBRButton.Bottom)
        End Select
    End If
End Sub

Private Function pvGetThumbSize() As Long
    
    On Error Resume Next
    
    Select Case m_eOrientation
        
        Case [oVertical]
        
            pvGetThumbSize = (m_uRctBRButton.Top - m_uRctTLButton.Bottom) \ (m_lAbsRange \ m_lLargeChange + 1)
            If (pvGetThumbSize < THUMBSIZE_MIN) Then
                pvGetThumbSize = THUMBSIZE_MIN
            End If
            
        Case [oHorizontal]
        
            pvGetThumbSize = (m_uRctBRButton.Left - m_uRctTLButton.Right) \ (m_lAbsRange \ m_lLargeChange + 1)
            If (pvGetThumbSize < THUMBSIZE_MIN) Then
                pvGetThumbSize = THUMBSIZE_MIN
            End If
    End Select
    
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Controling value
'----------------------------------------------------------------------------------------

Private Function pvScrollPosDec( _
                 ByVal lSteps As Long, _
                 Optional ByVal bForceRepaint As Boolean = False _
                 ) As Boolean
    
  Dim bChange    As Boolean
  Dim lValuePrev As Long
        
    lValuePrev = m_lValue
    
    m_lValue = m_lValue - lSteps
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
    End If
    
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    
    pvScrollPosDec = bChange
End Function

Private Function pvScrollPosInc( _
                 ByVal lSteps As Long, _
                 Optional ByVal bForceRepaint As Boolean = False _
                 ) As Boolean
    
  Dim bChange    As Boolean
  Dim lValuePrev As Long
        
    lValuePrev = m_lValue
    
    m_lValue = m_lValue + lSteps
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
    End If
    
    If (m_lValue <> lValuePrev) Then
        m_lThumbPos = pvGetThumbPos()
        Call pvSizeTrack
        bChange = True
    End If
    If (bChange Or bForceRepaint) Then
        Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
        If (bChange) Then
            RaiseEvent Change
        End If
    End If
    
    pvScrollPosInc = bChange
End Function

'----------------------------------------------------------------------------------------
' Positioning thumb and getting value from thumb position
'----------------------------------------------------------------------------------------

Private Function pvGetThumbPos() As Long

    On Error Resume Next
    
    Select Case m_eOrientation
        Case [oVertical]
            pvGetThumbPos = m_uRctTLButton.Bottom
            pvGetThumbPos = pvGetThumbPos + (m_uRctBRButton.Top - m_uRctTLButton.Bottom - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
        Case [oHorizontal]
            pvGetThumbPos = m_uRctTLButton.Right
            pvGetThumbPos = pvGetThumbPos + (m_uRctBRButton.Left - m_uRctTLButton.Right - m_lThumbSize) / m_lAbsRange * (m_lValue - m_lMin)
    End Select
    
    On Error GoTo 0
End Function

Private Function pvGetScrollPos() As Long
    
    On Error Resume Next
    
    Select Case m_eOrientation
        Case [oVertical]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.Bottom) / (m_uRctBRButton.Top - m_uRctTLButton.Bottom - m_lThumbSize) * m_lAbsRange
        Case [oHorizontal]
            pvGetScrollPos = m_lMin + (m_lThumbPos - m_uRctTLButton.Right) / (m_uRctBRButton.Left - m_uRctTLButton.Right - m_lThumbSize) * m_lAbsRange
    End Select
    
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Hit-Test
'----------------------------------------------------------------------------------------

Private Function pvHitTest(ByVal X As Long, ByVal Y As Long) As Long
    
    Select Case True
        Case PtInRect(m_uRctTLButton, X, Y)
            pvHitTest = HT_TLBUTTON
        Case PtInRect(m_uRctBRButton, X, Y)
            pvHitTest = HT_BRBUTTON
        Case PtInRect(m_uRctTLTrack, X, Y)
            pvHitTest = HT_TLTRACK
        Case PtInRect(m_uRctBRTrack, X, Y)
            pvHitTest = HT_BRTRACK
        Case PtInRect(m_uRctThumb, X, Y)
            pvHitTest = HT_THUMB
    End Select
End Function

Private Sub pvMakePoints( _
            ByVal lPoint As Long, _
            X As Long, _
            Y As Long _
            )
            
    If (lPoint And &H8000&) Then
        X = &H8000 Or (lPoint And &H7FFF&)
      Else
        X = lPoint And &HFFFF&
    End If
    If (lPoint And &H80000000) Then
        Y = (lPoint \ &H10000) - 1
      Else
        Y = lPoint \ &H10000
    End If
End Sub

'----------------------------------------------------------------------------------------
' Timing
'----------------------------------------------------------------------------------------

Private Sub pvSetTimer( _
            ByVal lTimerID As Long, _
            ByVal ldT As Long _
            )
    
    Call SetTimer(UserControl.hwnd, lTimerID, ldT, 0)
End Sub

Private Sub pvKillTimer( _
            ByVal lTimerID As Long _
            )
            
    Call KillTimer(UserControl.hwnd, lTimerID)
    m_eHitTestHot = HT_NOTHING
End Sub

'----------------------------------------------------------------------------------------
' Painting
'----------------------------------------------------------------------------------------

Private Sub pvInitPictures()
  Select Case m_eOrientation
  Case [oVertical]
    BarPics(SBack).CreateFromPicture LoadResPicture("SCRBACKV", vbResBitmap)
    BarPics(SThumb).CreateFromPicture LoadResPicture("SCRTHUMBV", vbResBitmap)
    BarPics(STGrip).CreateFromPicture LoadResPicture("SCRGRIPV", vbResBitmap)
    BarPics(SArrow).CreateFromPicture LoadResPicture("SCRARROWV", vbResBitmap)
    BarPics(SArrowTL).CreateFromPicture LoadResPicture("SCRARROWVT", vbResBitmap)
    BarPics(SArrowBR).CreateFromPicture LoadResPicture("SCRARROWVB", vbResBitmap)
  Case [oHorizontal]
    BarPics(SBack).CreateFromPicture LoadResPicture("SCRBACKH", vbResBitmap)
    BarPics(SThumb).CreateFromPicture LoadResPicture("SCRTHUMBH", vbResBitmap)
    BarPics(STGrip).CreateFromPicture LoadResPicture("SCRGRIPH", vbResBitmap)
    BarPics(SArrow).CreateFromPicture LoadResPicture("SCRARROWH", vbResBitmap)
    BarPics(SArrowTL).CreateFromPicture LoadResPicture("SCRARROWHL", vbResBitmap)
    BarPics(SArrowBR).CreateFromPicture LoadResPicture("SCRARROWHR", vbResBitmap)
  End Select
End Sub

Private Sub pvDrawFlatButton( _
            ByVal hdc As Long, _
            uRct As RECT, _
            ByVal lfArrowDirection As Long, _
            ByVal eState As eFlatButtonStateCts _
            )

  Dim uRctMem    As RECT
  
  Dim hDCMem1    As Long
  Dim hDCMem2    As Long
  Dim hBmp1      As Long
  Dim hBmp2      As Long
  Dim hBmpOld1   As Long
  Dim hBmpOld2   As Long
  
  Dim clrBkOld   As Long
  Dim clrTextOld As Long
        
    With uRct
    
        '-- Monochrome bitmap to convert the arrow to black/white mask
        hDCMem1 = CreateCompatibleDC(hdc)
        hBmp1 = CreateBitmap(.Right - .Left, .Bottom - .Top, 1, 1, ByVal 0)
        hBmpOld1 = SelectObject(hDCMem1, hBmp1)
        
        '-- Normal bitmap to draw the arrow into
        hDCMem2 = CreateCompatibleDC(hdc)
        hBmp2 = CreateCompatibleBitmap(hdc, .Right - .Left, .Bottom - .Top)
        hBmpOld2 = SelectObject(hDCMem2, hBmp2)
        
        '-- Draw frame normaly
        Call CopyRect(uRctMem, uRct)
        Call OffsetRect(uRctMem, -.Left, -.Top)
        Call DrawFrameControl(hDCMem2, uRctMem, DFC_SCROLL, DFCS_FLAT Or lfArrowDirection)
        
        Select Case eState
        
            Case [fbsNormal]
                
                '-- Nothing to do
                Call BitBlt(hdc, .Left, .Top, .Right - .Left, .Bottom - .Top, hDCMem2, 0, 0, vbSrcCopy)
            
            Case [fbsSelected]
                
                '-- Invert
                Call InvertRect(hDCMem2, uRctMem)
                Call BitBlt(hdc, .Left, .Top, .Right - .Left, .Bottom - .Top, hDCMem2, 0, 0, vbSrcCopy)
            
            Case [fbsHot]
            
                '-- Mask glyph
                Call SetBkColor(hDCMem2, GetSysColor(COLOR_BTNTEXT))
                Call BitBlt(hDCMem1, 0, 0, .Right - .Left, .Bottom - .Top, hDCMem2, 0, 0, vbSrcCopy)
                clrBkOld = SetBkColor(hdc, GetSysColor(COLOR_3DHIGHLIGHT))
                clrTextOld = SetTextColor(hdc, GetSysColor(COLOR_3DSHADOW))
                Call BitBlt(hdc, .Left, .Top, .Right - .Left, .Bottom - .Top, hDCMem1, 0, 0, vbSrcCopy)
                Call SetBkColor(hdc, clrBkOld)
                Call SetTextColor(hdc, clrTextOld)
        End Select
    End With
        
    '-- Clean up
    Call DeleteObject(SelectObject(hDCMem1, hBmpOld1))
    Call DeleteObject(SelectObject(hDCMem2, hBmpOld2))
    Call DeleteDC(hDCMem1)
    Call DeleteDC(hDCMem2)
End Sub

Private Function pvDrawThemePart( _
                 ByVal lhDC As Long, _
                 ByVal sClass As String, _
                 ByVal lPart As Long, _
                 ByVal lState As Long, _
                 lpRect As RECT _
                 ) As Boolean
  
  Dim hTheme As Long
    
    On Error GoTo Catch
    
    hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
    If (hTheme <> 0) Then
        pvDrawThemePart = (DrawThemeBackground(hTheme, lhDC, lPart, lState, lpRect, lpRect) = 0)
    End If
    
Catch:
    On Error GoTo 0
End Function

'----------------------------------------------------------------------------------------
' Misc.
'----------------------------------------------------------------------------------------

'-- Creating a pattern bitmap (track)

Private Sub pvCreatePatternBrush()

  Dim hBitmap          As Long
  Dim nPattern(1 To 8) As Integer
    
    '-- Brush pattern (8x8)
    nPattern(1) = &HAA
    nPattern(2) = &H55
    nPattern(3) = &HAA
    nPattern(4) = &H55
    nPattern(5) = &HAA
    nPattern(6) = &H55
    nPattern(7) = &HAA
    nPattern(8) = &H55
    
    '-- Create brush from bitmap
    hBitmap = CreateBitmap(8, 8, 1, 1, nPattern(1))
    m_hPatternBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
End Sub

'========================================================================================
' UserControl persistent properties
'========================================================================================

Private Sub UserControl_InitProperties()
    
    '-- Initialization default values
    Let m_lChangeDelay = CHANGEDELAY_DEF
    Let m_lChangeFrequency = CHANGEFREQUENCY_DEF
    Let m_lMin = MIN_DEF
    Let m_lMax = MAX_DEF
    Let m_lValue = VALUE_DEF
    Let m_lSmallChange = SMALLCHANGE_DEF
    Let m_lLargeChange = LARGECHANGE_DEF
    Let m_eOrientation = ORIENTATION_DEF
    Let m_bShowButtons = SHOWBUTTONS_DEF
    
    '-- Initialize rectangles
    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    '-- Bag properties
    With PropBag
        
        '-- Read inherently-stored properties
        Let UserControl.Enabled = .ReadProperty("Enabled", ENABLED_DEF)
        
        '-- Read 'memory' properties
        Let m_lMin = .ReadProperty("Min", MIN_DEF)
        Let m_lMax = .ReadProperty("Max", MAX_DEF)
        Let m_lValue = .ReadProperty("Value", VALUE_DEF)
        Let m_lSmallChange = .ReadProperty("SmallChange", SMALLCHANGE_DEF)
        Let m_lLargeChange = .ReadProperty("LargeChange", LARGECHANGE_DEF)
        Let m_lChangeDelay = .ReadProperty("ChangeDelay", CHANGEDELAY_DEF)
        Let m_lChangeFrequency = .ReadProperty("ChangeFrequency", CHANGEFREQUENCY_DEF)
        Let m_eOrientation = .ReadProperty("Orientation", ORIENTATION_DEF)
        Let m_bShowButtons = .ReadProperty("ShowButtons", SHOWBUTTONS_DEF)
    End With
    
    Call pvInitPictures
    
    '-- Initialize rectangles
    Let m_lAbsRange = m_lMax - m_lMin
    Call pvSizeButtons
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    
    '-- Run-time.
    If (Ambient.UserMode) Then
        '-- Subclass UC window and process following messages
        Call Subclass_Start(UserControl.hwnd)
        Call Subclass_AddMsg(UserControl.hwnd, WM_PAINT, [MSG_BEFORE])
        Call Subclass_AddMsg(UserControl.hwnd, WM_SIZE, [MSG_BEFORE])
        Call Subclass_AddMsg(UserControl.hwnd, WM_CANCELMODE)
        Call Subclass_AddMsg(UserControl.hwnd, WM_MOUSEMOVE)
        Call Subclass_AddMsg(UserControl.hwnd, WM_LBUTTONDOWN)
        Call Subclass_AddMsg(UserControl.hwnd, WM_LBUTTONUP)
        Call Subclass_AddMsg(UserControl.hwnd, WM_LBUTTONDBLCLK)
        Call Subclass_AddMsg(UserControl.hwnd, WM_TIMER)
        Call Subclass_AddMsg(UserControl.hwnd, WM_SYSCOLORCHANGE)
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, ENABLED_DEF)
        Call .WriteProperty("Min", m_lMin, MIN_DEF)
        Call .WriteProperty("Max", m_lMax, MAX_DEF)
        Call .WriteProperty("Value", m_lValue, VALUE_DEF)
        Call .WriteProperty("SmallChange", m_lSmallChange, SMALLCHANGE_DEF)
        Call .WriteProperty("LargeChange", m_lLargeChange, LARGECHANGE_DEF)
        Call .WriteProperty("ChangeDelay", m_lChangeDelay, CHANGEDELAY_DEF)
        Call .WriteProperty("ChangeFrequency", m_lChangeFrequency, CHANGEFREQUENCY_DEF)
        Call .WriteProperty("Orientation", m_eOrientation, ORIENTATION_DEF)
        Call .WriteProperty("ShowButtons", m_bShowButtons, SHOWBUTTONS_DEF)
    End With
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enable As Boolean)
    UserControl.Enabled = New_Enable
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Misc"
    Max = m_lMax
End Property

Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < m_lMin) Then
        New_Max = m_lMin
    End If
    m_lMax = New_Max
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue > m_lMax) Then
        m_lValue = m_lMax
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Misc"
    Min = m_lMin
End Property

Public Property Let Min(ByVal New_Min As Long)
    If (New_Min > m_lMax) Then
        New_Min = m_lMax
    End If
    m_lMin = New_Min
    m_lAbsRange = m_lMax - m_lMin
    If (m_lValue < m_lMin) Then
        m_lValue = m_lMin
    End If
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the value of an object."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Value.VB_UserMemId = 0
    Value = m_lValue
End Property

Public Property Let Value(ByVal New_Value As Long)

  Dim lValuePrev As Long

    If (New_Value < m_lMin) Then
        New_Value = m_lMin
    ElseIf (New_Value > m_lMax) Then
        New_Value = m_lMax
    End If
    lValuePrev = m_lValue
    m_lValue = New_Value
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
    
    If (m_lValue <> lValuePrev) Then
        RaiseEvent Change
    End If
End Property

Public Property Get SmallChange() As Long
Attribute SmallChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks a scroll arrow."
Attribute SmallChange.VB_ProcData.VB_Invoke_Property = ";Misc"
    SmallChange = m_lSmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Long)
    If (New_SmallChange < 1) Then
        New_SmallChange = 1
    End If
    m_lSmallChange = New_SmallChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get LargeChange() As Long
Attribute LargeChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks the scroll bar area."
    LargeChange = m_lLargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
    If (New_LargeChange < 1) Then
        New_LargeChange = 1
    End If
    m_lLargeChange = New_LargeChange
    m_lThumbSize = pvGetThumbSize()
    m_lThumbPos = pvGetThumbPos()
    Call pvSizeTrack
    Call InvalidateRect(UserControl.hwnd, ByVal 0, 0)
End Property

Public Property Get ChangeDelay() As Long
    ChangeDelay = m_lChangeDelay
End Property

Public Property Let ChangeDelay(ByVal New_ChangeDelay As Long)
    If (New_ChangeDelay < CHANGEDELAY_MIN) Then
        New_ChangeDelay = CHANGEDELAY_MIN
    End If
    m_lChangeDelay = New_ChangeDelay
End Property

Public Property Get ChangeFrequency() As Long
    ChangeFrequency = m_lChangeFrequency
End Property

Public Property Let ChangeFrequency(ByVal New_ChangeFrequency As Long)
    If (New_ChangeFrequency < CHANGEFREQUENCY_MIN) Then
        New_ChangeFrequency = CHANGEFREQUENCY_MIN
    End If
    m_lChangeFrequency = New_ChangeFrequency
End Property

Public Property Get Orientation() As sbOrientationCts
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Misc"
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As sbOrientationCts)
    If (New_Orientation < [oVertical]) Then
        New_Orientation = [oVertical]
    ElseIf (New_Orientation > [oHorizontal]) Then
        New_Orientation = [oHorizontal]
    End If
    m_eOrientation = New_Orientation
    Call pvInitPictures
    Call pvOnSize
End Property

Public Property Get ShowButtons() As Boolean
Attribute ShowButtons.VB_ProcData.VB_Invoke_Property = ";Misc"
    ShowButtons = m_bShowButtons
End Property

Public Property Let ShowButtons(ByVal New_ShowButtons As Boolean)
    m_bShowButtons = New_ShowButtons
    Call pvOnSize
End Property
Public Property Get hwnd() As Long
Attribute hwnd.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute hwnd.VB_UserMemId = -515
Attribute hwnd.VB_MemberFlags = "400"
    hwnd = UserControl.hwnd
End Property
Private Sub Subclass_AddMsg(ByVal lhwnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  
    With sc_aSubData(zIdx(lhwnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lhwnd As Long) As Long

  Const CODE_LEN              As Long = 202
  Const FUNC_CWP              As String = "CallWindowProcA"
  Const FUNC_EBM              As String = "EbMode"
  Const FUNC_SWL              As String = "SetWindowLongA"
  Const MOD_USER              As String = "user32"
  Const MOD_VBA5              As String = "vba5"
  Const MOD_VBA6              As String = "vba6"
  Const PATCH_01              As Long = 18
  Const PATCH_02              As Long = 68
  Const PATCH_03              As Long = 78
  Const PATCH_06              As Long = 116
  Const PATCH_07              As Long = 121
  Const PATCH_0A              As Long = 186
  Static aBuf(1 To CODE_LEN)  As Byte
  Static pCWP                 As Long
  Static pEbMode              As Long
  Static pSWL                 As Long
  Dim i                       As Long
  Dim j                       As Long
  Dim nSubIdx                 As Long
  Dim sHex                    As String
  
    If (aBuf(1) = 0) Then
  
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
    
        If (Subclass_InIDE) Then
            aBuf(16) = &H90
            aBuf(17) = &H90
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (pEbMode = 0) Then
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If
    
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
      Else
        nSubIdx = zIdx(lhwnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lhwnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
    End With
End Function

Private Sub Subclass_Stop(ByVal lhwnd As Long)
  
    With sc_aSubData(zIdx(lhwnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hwnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB()
        Erase .aMsgTblA()
    End With
End Sub

Private Sub Subclass_StopAll()
  
  Dim i As Long
  
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If (.hwnd <> 0) Then
                Call Subclass_Stop(.hwnd)
            End If
        End With
        i = i - 1
    Loop
End Sub

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  
  Dim nEntry  As Long
  Dim nOff1   As Long
  Dim nOff2   As Long
  
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGES
      Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop

        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
      Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Function zIdx(ByVal lhwnd As Long, Optional ByVal bAdd As Boolean = False) As Long

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hwnd = lhwnd) Then
                If (Not bAdd) Then
                    Exit Function
                End If
            ElseIf (.hwnd = 0) Then
                If (bAdd) Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False
    End If
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
