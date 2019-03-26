VERSION 5.00
Begin VB.UserControl AquaButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   DefaultCancel   =   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "AquaButton.ctx":0000
End
Attribute VB_Name = "AquaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Option Explicit

'***************************************************************************
'*  Title:      JC button
'*  Function:   An ownerdrawn Aqua Style button
'*  Author:     Juned Chhipa
'*  Created:    November 2008
'*  Contact me: juned.chhipa@yahoo.com
'*
'*  Copyright ?2008-2009 Juned Chhipa. All rights reserved.
'***************************************************************************
'* This control can be used as an alternative to Command Button. It is
'* a lightweight button control which will emulate new button styles.
'*
'* This control uses self-subclassing routines of Paul Caton.
'* Feel free to use this control. Please read Documentation.chm
'* Please send comments/suggestions/bug reports to juned.chhipa@yahoo.com
'****************************************************************************
'*
'* - CREDITS:
'* - Dana Seman  :-  Worked much for this control (Thanks a million)
'* - Paul Caton  :-  Self-Subclass Routines
'* - Noel Dacara :-  Inspiration for DropDown menu support
'* - Tuan Hai    :-  Numerous Suggestions and appreciating me ;)
'* - Fred.CPP    :-  For the amazing Aqua Style and for flexible tooltips
'* - Gonkuchi    :-  For his sub TransBlt to make grayscale pictures
'* - Carles P.V. :-  For fastest gradient routines
'*
'* I have tested this control painstakingly and tried my best to make
'* it work as a real command button. But still, if any bugs found,
'* please report to the email address provided above ;)

'****************************************************************************
'* This software is provided "as-is" without any express/implied warranty.  *
'* In no event shall the author be held liable for any damages arising      *
'* from the use of this software.                                           *
'* If you do not agree with these terms, do not install "JCButton". Use     *
'* of the program implicitly means you have agreed to these terms.          *        *
'                                                                           *
'* Permission is granted to anyone to use this software for any purpose,    *
'* including commercial use, and to alter and redistribute it, provided     *
'* that the following conditions are met:                                   *
'*                                                                          *
'* 1.All redistributions of source code files must retain all copyright     *
'*   notices that are currently in place, and this list of conditions       *
'*   without any modification.                                              *
'*                                                                          *
'* 2.All redistributions in binary form must retain all occurrences of      *
'*   above copyright notice and web site addresses that are currently in    *
'*   place (for example, in the About boxes).                               *
'*                                                                          *
'* 3.Modified versions in source or binary form must be plainly marked as   *
'*   such, and must not be misrepresented as being the original software.   *
'****************************************************************************

'* N'joy ;)

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As tLOGFONT) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'User32 Declares
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TransparentBlt Lib "MSIMG32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' --for tooltips
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long

Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'==========================================================================================================================================================================================================================================================================================
' Subclassing Declares
Private Enum MsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

'Windows Messages
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_THEMECHANGED           As Long = &H31A
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_MOVING                 As Long = &H216
Private Const WM_NCACTIVATE             As Long = &H86
Private Const WM_ACTIVATE               As Long = &H6

Private Const ALL_MESSAGES              As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED                As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC               As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04                  As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05                  As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08                  As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09                  As Long = 137                                      'Table A (after) entry count patch offset

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                   As Long
    dwFlags                  As TRACKMOUSEEVENT_FLAGS
    hwndTrack                As Long
    dwHoverTime              As Long
End Type

'for subclass
Private Type SubClassDatatype
    hWnd                         As Long
    nAddrSclass                  As Long
    nAddrOrig                    As Long
    nMsgCountA                   As Long
    nMsgCountB                   As Long
    aMsgTabelA()                 As Long
    aMsgTabelB()                 As Long
End Type

'for subclass
Private SubclassData()  As SubClassDatatype                                       'Subclass data array
Private TrackUser32     As Boolean

'Kernel32 declares used by the Subclasser
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'  End of Subclassing Declares
'==========================================================================================================================================================================================================================================================================================================

Public Enum enumAquaButtonModes
    [ebmCommandButton]
    [ebmCheckBox]
    [ebmOptionButton]
End Enum

#If False Then
Private ebmCommandButton, ebmCheckBox, ebmOptionButton
#End If

Public Enum enumAquaButtonStates
    [eStateNormal]              'Normal State
    [eStateOver]                'Hover State
    [eStateDown]                'Down State
End Enum

#If False Then
'A trick to preserve casing when typing in IDE
Private eStateNormal, eStateOver, eStateDown, eStateFocused
#End If

Public Enum enumAquaDisabledPicMode
    [edpBlended]
    [edpGrayed]
End Enum

#If False Then
'A trick to preserve casing when typing in IDE
Private edpBlended, edpGrayed
#End If

Public Enum enumAquaCaptionAlign
    [ecLeftAlign]
    [ecCenterAlign]
    [ecRightAlign]
End Enum

#If False Then
'A trick to preserve casing when typing in IDE
Private ecLeftAlign, ecCenterAlign, ecRightAlign
#End If

Public Enum enumAquaPictureAlign
    [epLeftEdge]
    [epLeftOfCaption]
    [epRightEdge]
    [epRightOfCaption]
    [epBehindCaption]
    [epTopEdge]
    [epTopOfCaption]
    [epBottomEdge]
    [epBottomOfCaption]
End Enum

#If False Then
Private epLeftEdge, epRightEdge, epRightOfCaption, epLeftOfCaption, epBehindCaption
Private epTopEdge, epTopOfCaption, epBottomEdge, epBottomOfCaption
#End If

' --Tooltip Icons
Public Enum enumAquaIconType
    TTNoIcon
    TTIconInfo
    TTIconWarning
    TTIconError
End Enum

#If False Then
Private TTNoIcon, TTIconInfo, TTIconWarning, TTIconError
#End If

' --Tooltip [ Balloon / Standard ]
Public Enum enumAquaTooltipStyle
    TooltipStandard
    TooltipBalloon
End Enum

#If False Then
Private TooltipStandard, TooltipBalloon
#End If

' --Caption effects
Public Enum enumAquaCaptionEffects
    [eseNone]
    [eseEmbossed]
    [eseEngraved]
    [eseShadowed]
    [eseOutline]
    [eseCover]
End Enum

#If False Then
Private eseNone, eseEmbossed, eseEngraved, eseShadowed, eseOutline, eseCover
#End If

Public Enum enumAquaPicEffect
    [epeNone]
    [epeLighter]
    [epeDarker]
End Enum

#If False Then
Private epeNone, epeLighter, epeDarker
#End If

' --For dropdown symbols
Public Enum enumAquaSymbol
    ebsNone
    ebsArrowUp = 5
    ebsArrowDown = 6
    ebsArrowRight = 4
End Enum

#If False Then
Private ebsArrowUp, ebsArrowDown, ebsNone, ebsArrowRight
#End If

Public Enum enumAquaXPThemeColors
    [ecsBlue]
    [ecsOliveGreen]
    [ecsSilver]
    [ecsCustom]
End Enum

' --A trick to preserve casing of enums while typing in IDE
#If False Then
Private ecsBlue, ecsOliveGreen, ecsSilver, ecsCustom
#End If

' --For gradient subs
Public Enum AquaGradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

' --A trick to preserve casing of enums when typing in IDE
#If False Then
Private gdHorizontal, gdVertical, gdDownwardDiagonal, gdUpwardDiagonal
#End If

Public Enum enumAquaMenuAlign
    [edaBottom] = 0
    [edaTop] = 1
    [edaLeft] = 2
    [edaRight] = 3
    [edaTopLeft] = 4
    [edaBottomLeft] = 5
    [edaTopRight] = 6
    [edaBottomRight] = 7
End Enum

#If False Then
Private edaBottom, edaTop, edaTopLeft, edaBottomLeft, edaTopRight, edaBottomRight
#End If

'  used for Button colors
Private Type tButtonColors
    tBackColor      As Long
    tDisabledColor  As Long
    tForeColor      As Long
    tForeColorOver  As Long
    tGreyText       As Long
End Type

'  used to define various graphics areas
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

''Tooltip Window Types
Private Type TOOLINFO
    lSize           As Long
    lFlags          As Long
    lHwnd           As Long
    lId             As Long
    lpRect          As RECT
    hInstance       As Long
    lpStr           As String
    lParam          As Long
End Type

''Tooltip Window Types [for UNICODE support]
Private Type TOOLINFOW
    lSize                As Long
    lFlags               As Long
    lHwnd                As Long
    lId                  As Long
    lpRect               As RECT
    hInstance            As Long
    lpStrW               As Long
    lParam               As Long
End Type

Private Type POINT
    x       As Long
    Y       As Long
End Type

' --Used for creating a drop down symbol
' --I m using Marlett Font to create that symbol
Private Type tLOGFONT
    lfHeight                        As Long
    lfWidth                         As Long
    lfEscapement                    As Long
    lfOrientation                   As Long
    lfWeight                        As Long
    lfItalic                        As Byte
    lfUnderline                     As Byte
    lfStrikeOut                     As Byte
    lfCharSet                       As Byte
    lfOutPrecision                  As Byte
    lfClipPrecision                 As Byte
    lfQuality                       As Byte
    lfPitchAndFamily                As Byte
    lfFaceName                      As String * 32
End Type

'  RGB Colors structure
Private Type RGBColor
    r       As Single
    g       As Single
    B       As Single
End Type

Private Type BITMAP
    bmType               As Long
    bmWidth              As Long
    bmHeight             As Long
    bmWidthBytes         As Long
    bmPlanes             As Integer
    bmBitsPixel          As Integer
    bmBits               As Long
End Type

'  for gradient painting and bitmap tiling
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type ICONINFO
    fIcon       As Long
    xHotspot    As Long
    yHotspot    As Long
    hbmMask     As Long
    hbmColor    As Long
End Type

Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type RGBQUAD
    rgbBlue              As Byte
    rgbGreen             As Byte
    rgbRed               As Byte
    rgbAlpha             As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
End Type

' --constants for unicode support
Private Const VER_PLATFORM_WIN32_NT = 2

' --constants for  Flat Button
Private Const BDR_RAISEDINNER   As Long = &H4

' --constants for Win 98 style buttons
Private Const BDR_SUNKEN95 As Long = &HA
Private Const BDR_RAISED95 As Long = &H5

Private Const BF_LEFT       As Long = &H1
Private Const BF_TOP        As Long = &H2
Private Const BF_RIGHT      As Long = &H4
Private Const BF_BOTTOM     As Long = &H8
Private Const BF_RECT       As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

' --System Hand Pointer
Private Const IDC_HAND As Long = 32649

' --Color Constant
Private Const COLOR_BTNFACE      As Long = 15
Private Const COLOR_BTNHIGHLIGHT As Long = 20
Private Const COLOR_BTNSHADOW    As Long = 16
Private Const COLOR_HIGHLIGHT    As Long = 13
Private Const COLOR_GRAYTEXT     As Long = 17
Private Const CLR_INVALID        As Long = &HFFFF
Private Const DIB_RGB_COLORS     As Long = 0

' --Windows Messages
Private Const WM_USER                   As Long = &H400
Private Const GWL_STYLE                 As Long = -16
Private Const WS_CAPTION                As Long = &HC00000
Private Const WS_THICKFRAME             As Long = &H40000
Private Const WS_MINIMIZEBOX            As Long = &H20000
Private Const SWP_REFRESH               As Long = (&H1 Or &H2 Or &H4 Or &H20)
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_SHOWWINDOW            As Long = &H40
Private Const HWND_TOPMOST              As Long = -&H1
Private Const CW_USEDEFAULT             As Long = &H80000000

''Tooltip Window Constants
Private Const TTS_NOPREFIX As Long = &H2
Private Const TTF_TRANSPARENT As Long = &H100
Private Const TTF_IDISHWND As Long = &H1
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const TTF_CENTERTIP As Long = &H2
Private Const TTM_ADDTOOLA As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW As Long = (WM_USER + 50)
Private Const TTM_ACTIVATE As Long = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA As Long = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH As Long = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR As Long = (WM_USER + 20)
Private Const TTM_SETTITLE As Long = (WM_USER + 32)
Private Const TTM_SETTITLEW As Long = (WM_USER + 33)
Private Const TTS_BALLOON As Long = &H40
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTF_SUBCLASS As Long = &H10
Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"

' --Formatting Text Consts
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_RTLREADING = &H20000              ' Right to left
Private Const DT_DRAWFLAG As Long = DT_CENTER Or DT_WORDBREAK

' --for drawing Icon Constants
Private Const DI_NORMAL As Long = &H3

' --Property Variables:

Private m_Buttonstate       As enumAquaButtonStates     'Normal / Over / Down

Private m_bIsDown           As Boolean              'Is button is pressed?
Private m_bMouseInCtl       As Boolean              'Is Mouse in Control
Private m_bHasFocus         As Boolean              'Has focus?
Private m_bHandPointer      As Boolean              'Use Hand Pointer
Private m_lCursor           As Long
Private m_bDefault          As Boolean              'Is Default?
Private m_DropDownSymbol    As enumAquaSymbol
Private m_bDropDownSep      As Boolean
Private m_ButtonMode        As enumAquaButtonModes      'Command/Check/Option button
Private m_CaptionEffects    As enumAquaCaptionEffects
Private m_bValue            As Boolean              'Value (Checked/Unchekhed)
Private m_bShowFocus        As Boolean              'Bool to show focus
Private m_bParentActive     As Boolean              'Parent form Active or not
Private m_lParenthWnd       As Long                 'Is parent active?
Private m_WindowsNT         As Long                 'OS Supports Unicode?
Private m_bEnabled          As Boolean              'Enabled/Disabled
Private m_Caption           As String               'String to draw caption
Private m_CaptionAlign      As enumAquaCaptionAlign
Private m_bColors           As tButtonColors        'Button Colors
Private m_bUseMaskColor     As Boolean              'Transparent areas
Private m_lMaskColor        As Long                 'Set Transparent color
Private m_lButtonRgn        As Long                 'Button Region
Private m_bIsSpaceBarDown   As Boolean              'Space bar down boolean
Private m_ButtonRect        As RECT                 'Button Position
Private m_FocusRect         As RECT
Private WithEvents mFont    As StdFont
Attribute mFont.VB_VarHelpID = -1
Private m_lXPColor          As enumAquaXPThemeColors

Private m_lDownButton       As Integer              'For click/Dblclick events
Private m_lDShift           As Integer              'A flag for dblClick
Private m_lDX               As Single
Private m_lDY               As Single

' --Popup menu variables
Private m_bPopupEnabled     As Boolean              'Popus is enabled
Private m_bPopupShown       As Boolean              'Popupmenu is shown
Private m_bPopupInit        As Boolean              'Flag to prevent WM_MOUSLEAVE to redraw the button
Private DropDownMenu        As VB.Menu              'Popupmenu to be shown
Private MenuAlign           As enumAquaMenuAlign        'PopupMenu Alignments
Private MenuFlags           As Long                 'PopupMenu Flags
Private DefaultMenu         As VB.Menu              'Default menu in the popupmenu

' --Tooltip variables
Private m_sTooltipText      As String
Private m_sTooltiptitle     As String
Private m_lToolTipIcon      As enumAquaIconType
Private m_lTooltipType      As enumAquaTooltipStyle
Private m_lttBackColor      As Long
Private m_lttForeColor      As Long
Private m_lttCentered       As Boolean
Private m_lttHwnd           As Long
Private ttip                As TOOLINFO
Private m_bttRTL            As Boolean
Private m_hMode         As Long                 'Added this, as tooltips

' --Caption variables
Private CaptionW As Long                            'Width of Caption
Private CaptionH As Long                            'Height of Caption
Private CaptionX As Long                            'Left of Caption
Private CaptionY As Long                            'Top of Caption
Private lpSignRect As RECT                          'Drop down Symbol rect
Private m_bRTL          As Boolean
Private m_TextRect As RECT                          'Caption drawing area

' --Picture variables
Private m_Picture           As StdPicture
Private m_PictureHot        As StdPicture
Private m_PictureDown       As StdPicture
Private m_PicSemiTrans      As Boolean
Private m_PicDisabledMode   As enumAquaDisabledPicMode
Private m_PictureAlign      As enumAquaPictureAlign     'Picture Alignments
Private m_PicEffectonOver   As enumAquaPicEffect        'Blend effect
Private m_PicEffectonDown   As enumAquaPicEffect        'Blend effect
Private m_bPicPushOnHover   As Boolean
Private m_PictureShadow As Boolean
Private m_PictureOpacity As Byte
Private m_PicOpacityOnOver As Byte
Private PicH     As Long
Private PicW     As Long
Private aLighten(255)   As Byte                 'Light Picture
Private aDarken(255)    As Byte                 'Dark Picture

Private tmppic   As New StdPicture                  'Temp picture
Private PicX     As Long                            'X position of picture
Private PicY     As Long                            'Y Position of Picture
Private m_PicRect  As RECT                          'Picture drawing area

Private lh       As Long                            'ScaleHeight of button
Private lw       As Long                            'ScaleWidth of button

'  Events
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over the button."
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user clicks over the button twice."
Attribute DblClick.VB_UserMemId = -601
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occrus when the cursor moves around the button for the first time."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the cursor leaves/moves outside the button."
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the cursor moves over the button."
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the button has the focus."
Attribute MouseUp.VB_UserMemId = -607
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the button has the focus."
Attribute MouseDown.VB_UserMemId = -605
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while the button has the focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while the button has the focus."
Attribute KeyUp.VB_UserMemId = -604
Public Event KeyPress(KeyAcsii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603

Private Sub DrawLineApi(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

'****************************************************************************
'*  draw lines
'****************************************************************************

Dim pt      As POINT
Dim hPen    As Long
Dim hPenOld As Long

    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X1, Y1, pt
    LineTo hdc, X2, Y2
    SelectObject hdc, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld

End Sub

Private Function BlendColorEx(Color1 As Long, Color2 As Long, Optional Percent As Long) As Long

'   Combines two colors together by how many percent.
'   Inspired from dcbutton (honestly not copied!!) hehe

Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim r3 As Long, g3 As Long, b3 As Long

'If Percent <= 0 Then Percent = 0':(?-> replaced by:

    If Percent <= 0 Then
        Percent = 0
    End If
    'If Percent >= 100 Then Percent = 100':(?-> replaced by:
    If Percent >= 100 Then
        Percent = 100
    End If

    r1 = Color1 And 255
    g1 = (Color1 \ 256) And 255
    b1 = (Color1 \ 65536) And 255

    r2 = Color2 And 255
    g2 = (Color2 \ 256) And 255
    b2 = (Color2 \ 65536) And 255

    r3 = r1 + (r1 - r2) * Percent \ 100
    g3 = g1 + (g1 - g2) * Percent \ 100
    b3 = b1 + (b1 - b2) * Percent \ 100

    BlendColorEx = r3 + 256& * g3 + 65536 * b3

End Function

Private Function BlendColors(ByVal lBackColorFrom As Long, ByVal lBackColorTo As Long) As Long

'***************************************************************************
'*  Combines (mix) two colors                                              *
'*  This is another method in which you can't specify percentage
'***************************************************************************

    BlendColors = RGB(((lBackColorFrom And &HFF) + (lBackColorTo And &HFF)) / 2, (((lBackColorFrom \ &H100) And &HFF) + ((lBackColorTo \ &H100) And &HFF)) / 2, (((lBackColorFrom \ &H10000) And &HFF) + ((lBackColorTo \ &H10000) And &HFF)) / 2)

End Function

Private Sub DrawRectangle(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)

'****************************************************************************
'*  Draws a rectangle specified by coords and color of the rectangle        *
'****************************************************************************

Dim brect As RECT
Dim hBrush As Long
Dim ret As Long

    brect.Left = x
    brect.Top = Y
    brect.Right = x + Width
    brect.Bottom = Y + Height

    hBrush = CreateSolidBrush(Color)

    ret = FrameRect(hdc, brect, hBrush)

    ret = DeleteObject(hBrush)

End Sub

Private Sub DrawFocusRectangle(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)

'****************************************************************************
'*  Draws a Focus Rectangle inside button if m_bShowFocus property is True  *
'****************************************************************************

Dim brect As RECT
Dim RetVal As Long

    brect.Left = x
    brect.Top = Y
    brect.Right = x + Width
    brect.Bottom = Y + Height

    RetVal = DrawFocusRect(hdc, brect)

End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False)

   '****************************************************************************
   '* Routine : To make transparent and grayscale images
   '* Author  : Gonkuchi
   '
   '* Modified by Dana Seaman
   '****************************************************************************

   Dim B                As Long, h As Long, F As Long, i As Long, newW As Long
   Dim TmpDC            As Long, TmpBmp As Long, TmpObj As Long
   Dim Sr2DC            As Long, Sr2Bmp As Long, Sr2Obj As Long
   Dim DataDest()       As RGBTRIPLE, DataSrc() As RGBTRIPLE
   Dim Info             As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
   Dim hOldOb           As Long, PicEffect As enumAquaPicEffect
   Dim SrcDC            As Long, tObj As Long, ttt As Long
   Dim bDisOpacity      As Byte
   Dim OverOpacity      As Byte
   Dim a2               As Long
   Dim a1               As Long

   If DstW = 0 Or DstH = 0 Then Exit Sub
   If SrcPic Is Nothing Then Exit Sub

   If m_Buttonstate = eStateOver Then
      PicEffect = m_PicEffectonOver
   ElseIf m_Buttonstate = eStateDown Then
      PicEffect = m_PicEffectonDown
   End If
   
   If Not m_bEnabled Then
      Select Case m_PicDisabledMode
      Case edpBlended
         bDisOpacity = 52
      Case edpGrayed
         bDisOpacity = m_PictureOpacity * 0.75
         isGreyscale = True
      End Select
   End If
   
   If m_Buttonstate = eStateOver Then
      OverOpacity = m_PicOpacityOnOver
   End If

   SrcDC = CreateCompatibleDC(hdc)

   If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
   If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)

   If SrcPic.Type = vbPicTypeBitmap Then 'check if it's an icon or a bitmap
      tObj = SelectObject(SrcDC, SrcPic)
   Else
      Dim hBrush           As Long
      tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
      hBrush = CreateSolidBrush(TransColor)
      DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, DI_NORMAL
      DeleteObject hBrush
   End If

   TmpDC = CreateCompatibleDC(SrcDC)
   Sr2DC = CreateCompatibleDC(SrcDC)
   TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   TmpObj = SelectObject(TmpDC, TmpBmp)
   Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
   ReDim DataDest(DstW * DstH * 3 - 1)
   ReDim DataSrc(UBound(DataDest))
   With Info.bmiHeader
      .biSize = Len(Info.bmiHeader)
      .biWidth = DstW
      .biHeight = DstH
      .biPlanes = 1
      .biBitCount = 24
   End With

   BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
   BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
   GetDIBits TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0
   GetDIBits Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0

   If BrushColor > 0 Then
      BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
      BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
      BrushRGB.rgbRed = BrushColor And &HFF
   End If

   ' --No Maskcolor to use
   If Not m_bUseMaskColor Then TransColor = -1

   newW = DstW - 1

   For h = 0 To DstH - 1
      F = h * DstW
      For B = 0 To newW
         i = F + B
         If m_Buttonstate = eStateOver Then
            a1 = OverOpacity
         Else
            a1 = IIf(m_bEnabled, m_PictureOpacity, bDisOpacity)
         End If
         a2 = 255 - a1
         If GetNearestColor(hdc, CLng(DataSrc(i).rgbRed) + 256& * DataSrc(i).rgbGreen + 65536 * DataSrc(i).rgbBlue) <> TransColor Then
            With DataDest(i)
               If BrushColor > -1 Then
                  If MonoMask Then
                     If (CLng(DataSrc(i).rgbRed) + DataSrc(i).rgbGreen + DataSrc(i).rgbBlue) <= 384 Then DataDest(i) = BrushRGB
                  Else
                     If a1 = 255 Then
                        DataDest(i) = BrushRGB
                     ElseIf a1 > 0 Then
                        .rgbRed = (a2 * .rgbRed + a1 * BrushRGB.rgbRed) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * BrushRGB.rgbGreen) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * BrushRGB.rgbBlue) \ 256
                     End If
                  End If
               Else
                  If isGreyscale Then
                     gCol = CLng(DataSrc(i).rgbRed * 0.3) + DataSrc(i).rgbGreen * 0.59 + DataSrc(i).rgbBlue * 0.11
                     If a1 = 255 Then
                        .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                     ElseIf a1 > 0 Then
                        .rgbRed = (a2 * .rgbRed + a1 * gCol) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * gCol) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * gCol) \ 256
                     End If
                  Else
                     If a1 = 255 Then
                        If PicEffect = epeLighter Then
                           .rgbRed = aLighten(DataSrc(i).rgbRed)
                           .rgbGreen = aLighten(DataSrc(i).rgbGreen)
                           .rgbBlue = aLighten(DataSrc(i).rgbBlue)
                        ElseIf PicEffect = epeDarker Then
                           .rgbRed = aDarken(DataSrc(i).rgbRed)
                           .rgbGreen = aDarken(DataSrc(i).rgbGreen)
                           .rgbBlue = aDarken(DataSrc(i).rgbBlue)
                        Else
                           DataDest(i) = DataSrc(i)
                        End If
                     ElseIf a1 > 0 Then
                        If (PicEffect = epeLighter) Then
                           .rgbRed = (a2 * .rgbRed + a1 * aLighten(DataSrc(i).rgbRed)) \ 256
                           .rgbGreen = (a2 * .rgbGreen + a1 * aLighten(DataSrc(i).rgbGreen)) \ 256
                           .rgbBlue = (a2 * .rgbBlue + a1 * aLighten(DataSrc(i).rgbBlue)) \ 256
                        ElseIf PicEffect = epeDarker Then
                           .rgbRed = (a2 * .rgbRed + a1 * aDarken(DataSrc(i).rgbRed)) \ 256
                           .rgbGreen = (a2 * .rgbGreen + a1 * aDarken(DataSrc(i).rgbGreen)) \ 256
                           .rgbBlue = (a2 * .rgbBlue + a1 * aDarken(DataSrc(i).rgbBlue)) \ 256
                        Else
                           .rgbRed = (a2 * .rgbRed + a1 * DataSrc(i).rgbRed) \ 256
                           .rgbGreen = (a2 * .rgbGreen + a1 * DataSrc(i).rgbGreen) \ 256
                           .rgbBlue = (a2 * .rgbBlue + a1 * DataSrc(i).rgbBlue) \ 256
                        End If
                     End If
                  End If
               End If
            End With
         End If
      Next B
   Next h

   ' /--Paint it!
   SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0

   Erase DataDest, DataSrc
   DeleteObject SelectObject(TmpDC, TmpObj)
   DeleteObject SelectObject(Sr2DC, Sr2Obj)
   If SrcPic.Type = vbPicTypeIcon Then DeleteObject SelectObject(SrcDC, tObj)
   DeleteDC TmpDC
   DeleteDC Sr2DC
   DeleteObject tObj
   DeleteDC SrcDC

End Sub

Private Sub TransBlt32(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal BrushColor As Long = -1, Optional ByVal isGreyscale As Boolean = False)

   '****************************************************************************
   '* Routine : Renders 32 bit Bitmap                                          *
   '* Author  : Dana Seaman                                                    *
   '****************************************************************************

   Dim B                As Long, h As Long, F As Long, i As Long, newW As Long
   Dim TmpDC            As Long, TmpBmp As Long, TmpObj As Long
   Dim Sr2DC            As Long, Sr2Bmp As Long, Sr2Obj As Long
   Dim DataDest()       As RGBQUAD, DataSrc() As RGBQUAD
   Dim Info             As BITMAPINFO, BrushRGB As RGBQUAD, gCol As Long
   Dim hOldOb           As Long, PicEffect As enumAquaPicEffect
   Dim SrcDC            As Long, tObj As Long, ttt As Long
   Dim bDisOpacity      As Byte
   Dim OverOpacity      As Byte
   Dim a2               As Long
   Dim a1               As Long

   If DstW = 0 Or DstH = 0 Then Exit Sub
   If SrcPic Is Nothing Then Exit Sub

   If m_Buttonstate = eStateOver Then
      PicEffect = m_PicEffectonOver
   ElseIf m_Buttonstate = eStateDown Then
      PicEffect = m_PicEffectonDown
   End If
   
   If Not m_bEnabled Then
      Select Case m_PicDisabledMode
      Case edpBlended
         bDisOpacity = 52
      Case edpGrayed
         bDisOpacity = m_PictureOpacity * 0.75
         isGreyscale = True
      End Select
   End If
      
   If m_Buttonstate = eStateOver Then
      OverOpacity = m_PicOpacityOnOver
   End If
   
   SrcDC = CreateCompatibleDC(hdc)

   If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
   If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)

   tObj = SelectObject(SrcDC, SrcPic)

   TmpDC = CreateCompatibleDC(SrcDC)
   Sr2DC = CreateCompatibleDC(SrcDC)

   TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
   TmpObj = SelectObject(TmpDC, TmpBmp)
   Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)

   With Info.bmiHeader
      .biSize = Len(Info.bmiHeader)
      .biWidth = DstW
      .biHeight = DstH
      .biPlanes = 1
      .biBitCount = 32
      .biSizeImage = 4 * ((DstW * .biBitCount + 31) \ 32) * DstH
   End With
   ReDim DataDest(Info.bmiHeader.biSizeImage - 1)
   ReDim DataSrc(UBound(DataDest))

   BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
   BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
   GetDIBits TmpDC, TmpBmp, 0, DstH, DataDest(0), Info, 0
   GetDIBits Sr2DC, Sr2Bmp, 0, DstH, DataSrc(0), Info, 0

   If BrushColor <> -1 Then
      BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
      BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
      BrushRGB.rgbRed = BrushColor And &HFF
   End If

   newW = DstW - 1

   For h = 0 To DstH - 1
      F = h * DstW
      For B = 0 To newW
         i = F + B
         If m_bEnabled Then
             If m_Buttonstate = eStateOver Then
                a1 = (CLng(DataSrc(i).rgbAlpha) * OverOpacity) \ 255
             Else
                a1 = (CLng(DataSrc(i).rgbAlpha) * m_PictureOpacity) \ 255
             End If
         Else
            a1 = (CLng(DataSrc(i).rgbAlpha) * bDisOpacity) \ 255
         End If
         a2 = 255 - a1
         With DataDest(i)
            If BrushColor <> -1 Then
               If a1 = 255 Then
                  DataDest(i) = BrushRGB
               ElseIf a1 > 0 Then
                  .rgbRed = (a2 * .rgbRed + a1 * BrushRGB.rgbRed) \ 256
                  .rgbGreen = (a2 * .rgbGreen + a1 * BrushRGB.rgbGreen) \ 256
                  .rgbBlue = (a2 * .rgbBlue + a1 * BrushRGB.rgbBlue) \ 256
               End If
            Else
               If isGreyscale Then
                  gCol = CLng(DataSrc(i).rgbRed * 0.3) + DataSrc(i).rgbGreen * 0.59 + DataSrc(i).rgbBlue * 0.11
                  If a1 = 255 Then
                     .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                  ElseIf a1 > 0 Then
                     .rgbRed = (a2 * .rgbRed + a1 * gCol) \ 256
                     .rgbGreen = (a2 * .rgbGreen + a1 * gCol) \ 256
                     .rgbBlue = (a2 * .rgbBlue + a1 * gCol) \ 256
                  End If
               Else
                  If a1 = 255 Then
                     If (PicEffect = epeLighter) Then
                        .rgbRed = aLighten(DataSrc(i).rgbRed)
                        .rgbGreen = aLighten(DataSrc(i).rgbGreen)
                        .rgbBlue = aLighten(DataSrc(i).rgbBlue)
                     ElseIf PicEffect = epeDarker Then
                        .rgbRed = aDarken(DataSrc(i).rgbRed)
                        .rgbGreen = aDarken(DataSrc(i).rgbGreen)
                        .rgbBlue = aDarken(DataSrc(i).rgbBlue)
                     Else
                        DataDest(i) = DataSrc(i)
                     End If
                  ElseIf a1 > 0 Then
                     If (PicEffect = epeLighter) Then
                        .rgbRed = (a2 * .rgbRed + a1 * aLighten(DataSrc(i).rgbRed)) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * aLighten(DataSrc(i).rgbGreen)) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * aLighten(DataSrc(i).rgbBlue)) \ 256
                     ElseIf PicEffect = epeDarker Then
                        .rgbRed = (a2 * .rgbRed + a1 * aDarken(DataSrc(i).rgbRed)) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * aDarken(DataSrc(i).rgbGreen)) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * aDarken(DataSrc(i).rgbBlue)) \ 256
                     Else
                        .rgbRed = (a2 * .rgbRed + a1 * DataSrc(i).rgbRed) \ 256
                        .rgbGreen = (a2 * .rgbGreen + a1 * DataSrc(i).rgbGreen) \ 256
                        .rgbBlue = (a2 * .rgbBlue + a1 * DataSrc(i).rgbBlue) \ 256
                     End If
                  End If
               End If
            End If
         End With
      Next B
   Next h

   ' /--Paint it!
   SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, DataDest(0), Info, 0

   Erase DataDest, DataSrc
   DeleteObject SelectObject(TmpDC, TmpObj)
   DeleteObject SelectObject(Sr2DC, Sr2Obj)
   If SrcPic.Type = vbPicTypeIcon Then DeleteObject SelectObject(SrcDC, tObj)
   DeleteDC TmpDC
   DeleteDC Sr2DC
   DeleteObject tObj
   DeleteDC SrcDC

End Sub

' --By Dana Seaman

Private Function Lighten(ByVal Color As Byte) As Byte

Dim lColor           As Long

    lColor = Color * 1.15
    If lColor > 255 Then
        Lighten = 255
    Else
        Lighten = lColor
    End If

End Function

' --By Dana Seaman

Private Function Darken(ByVal Color As Byte) As Byte

    Darken = Color * 0.85

End Function

Private Sub DrawGradientEx(ByVal x As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal GradientDirection As AquaGradientDirectionCts)

'****************************************************************************
'* Draws very fast Gradient in four direction.                              *
'* Author: Carles P.V (Gradient Master)                                     *
'* This routine works as a heart for this control.                          *
'* Thank you so much Carles.                                                *
'****************************************************************************

Dim uBIH    As BITMAPINFOHEADER
Dim lBits() As Long
Dim lGrad() As Long

Dim r1      As Long
Dim g1      As Long
Dim b1      As Long
Dim r2      As Long
Dim g2      As Long
Dim b2      As Long
Dim dR      As Long
Dim dG      As Long
Dim dB      As Long

Dim Scan    As Long
Dim i       As Long
Dim iEnd    As Long
Dim iOffset As Long
Dim j       As Long
Dim jEnd    As Long
Dim iGrad   As Long

'-- A minor check

'If (Width < 1 Or Height < 1) Then Exit Sub

    If (Width < 1 Or Height < 1) Then
        Exit Sub
    End If

    '-- Decompose colors
    Color1 = Color1 And &HFFFFFF
    r1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    g1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    b1 = Color1 Mod &H100&
    Color2 = Color2 And &HFFFFFF
    r2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    g2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    b2 = Color2 Mod &H100&

    '-- Get color distances
    dR = r2 - r1
    dG = g2 - g1
    dB = b2 - b1

    '-- Size gradient-colors array
    Select Case GradientDirection
    Case [gdHorizontal]
        ReDim lGrad(0 To Width - 1)
    Case [gdVertical]
        ReDim lGrad(0 To Height - 1)
    Case Else
        ReDim lGrad(0 To Width + Height - 2)
    End Select

    '-- Calculate gradient-colors
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (b1 \ 2 + b2 \ 2) + 256 * (g1 \ 2 + g2 \ 2) + 65536 * (r1 \ 2 + r2 \ 2)
    Else
        For i = 0 To iEnd
            lGrad(i) = b1 + (dB * i) \ iEnd + 256 * (g1 + (dG * i) \ iEnd) + 65536 * (r1 + (dR * i) \ iEnd)
        Next i
    End If

    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width

    '-- Render gradient DIB
    Select Case GradientDirection

    Case [gdHorizontal]

        For j = 0 To jEnd
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(i - iOffset)
            Next i
            iOffset = iOffset + Scan
        Next j

    Case [gdVertical]

        For j = jEnd To 0 Step -1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(j)
            Next i
            iOffset = iOffset + Scan
        Next j

    Case [gdDownwardDiagonal]

        iOffset = jEnd * Scan
        For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset - Scan
            iGrad = j
        Next j

    Case [gdUpwardDiagonal]

        iOffset = 0
        For j = 1 To jEnd + 1
            For i = iOffset To iEnd + iOffset
                lBits(i) = lGrad(iGrad)
                iGrad = iGrad + 1
            Next i
            iOffset = iOffset + Scan
            iGrad = j
        Next j
    End Select

    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With

    '-- Paint it!
    StretchDIBits hdc, x, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy

End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional ByRef hPalette As Long = 0) As Long

'****************************************************************************
'*  System color code to long rgb                                           *
'****************************************************************************

    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function

Private Sub RedrawButton()

'****************************************************************************
'*  The main routine of this usercontrol. Everything is drawn here.         *
'****************************************************************************

    UserControl.Cls                                'Clears usercontrol
    lh = ScaleHeight
    lw = ScaleWidth

    SetRect m_ButtonRect, 0, 0, lw, lh             'Sets the button rectangle

    If m_bParentActive Then
        If m_bDefault Or m_bHasFocus Then           'MAC OS X draws
            If m_Buttonstate = eStateDown Then             'Hotstate
                m_Buttonstate = eStateDown                 'for focused buttons
            Else
                m_Buttonstate = eStateOver
            End If
        End If
    End If

    On Error GoTo h:
    UserControl.BackColor = Ambient.BackColor
    Select Case m_Buttonstate

    Case eStateNormal
        CreateRegion
        DrawAquaNormal
    Case eStateOver
        DrawAquaHot
    Case eStateDown
        DrawAquaDown
    End Select

    DrawPicwithCaption

h:

End Sub

Private Sub DrawAquaNormal()

Dim tmph As Long, tmpw As Long
Dim tmph1 As Long, tmpw1 As Long
Dim lpRect As RECT

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    SetRect lpRect, 4, 4, lw - 4, lh - 4
    PaintRect &HEAE7E8, lpRect

    SetPixel hdc, 6, 0, &HFEFEFE: SetPixel hdc, 7, 0, &HE6E6E6: SetPixel hdc, 8, 0, &HACACAC: SetPixel hdc, 9, 0, &H7A7A7A: SetPixel hdc, 10, 0, &H6C6C6C: SetPixel hdc, 11, 0, &H6B6B6B: SetPixel hdc, 12, 0, &H6F6F6F: SetPixel hdc, 13, 0, &H716F6F: SetPixel hdc, 14, 0, &H727070: SetPixel hdc, 15, 0, &H676866: SetPixel hdc, 16, 0, &H6C6D6B: SetPixel hdc, 17, 0, &H67696A: SetPixel hdc, 5, 1, &HEFEFEF: SetPixel hdc, 6, 1, &H939393: SetPixel hdc, 7, 1, &H676767: SetPixel hdc, 8, 1, &H797979: SetPixel hdc, 9, 1, &HB3B3B3: SetPixel hdc, 10, 1, &HDBDBDB: SetPixel hdc, 11, 1, &HEBEDEE: SetPixel hdc, 12, 1, &HF5F4F6: SetPixel hdc, 13, 1, &HF5F4F6: SetPixel hdc, 14, 1, &HF5F4F6: SetPixel hdc, 15, 1, &HF5F4F6: SetPixel hdc, 16, 1, &HF5F4F6: SetPixel hdc, 17, 1, &HF5F4F6
    SetPixel hdc, 3, 2, &HFEFEFE: SetPixel hdc, 4, 2, &HE5E5E5: SetPixel hdc, 5, 2, &H737373: SetPixel hdc, 6, 2, &H656565: SetPixel hdc, 7, 2, &H939393: SetPixel hdc, 8, 2, &HDCDCDC: SetPixel hdc, 9, 2, &HE9E9E9: SetPixel hdc, 10, 2, &HF2F1F3: SetPixel hdc, 11, 2, &HF3F2F4: SetPixel hdc, 12, 2, &HF2F1F3: SetPixel hdc, 13, 2, &HF3F2F4: SetPixel hdc, 14, 2, &HF2F1F3: SetPixel hdc, 15, 2, &HF3F2F4: SetPixel hdc, 16, 2, &HF2F1F3: SetPixel hdc, 17, 2, &HF3F2F4: SetPixel hdc, 3, 3, &HEEEEEE: SetPixel hdc, 4, 3, &H717171: SetPixel hdc, 5, 3, &H6C6C6C: SetPixel hdc, 6, 3, &H909090: SetPixel hdc, 7, 3, &HD2D2D2: SetPixel hdc, 8, 3, &HE3E3E3: SetPixel hdc, 9, 3, &HECECEC: SetPixel hdc, 10, 3, &HEDEDED: SetPixel hdc, 11, 3, &HEEEEEE: SetPixel hdc, 12, 3, &HEDEDED: SetPixel hdc, 13, 3, &HEEEEEE: SetPixel hdc, 14, 3, &HEDEDED: SetPixel hdc, 15, 3, &HEEEEEE: SetPixel hdc, 16, 3, &HEDEDED: SetPixel hdc, 17, 3, &HEEEEEE
    SetPixel hdc, 2, 4, &HFBFBFB: SetPixel hdc, 3, 4, &H858585: SetPixel hdc, 4, 4, &H686868: SetPixel hdc, 5, 4, &H959595: SetPixel hdc, 6, 4, &HB1B1B1: SetPixel hdc, 7, 4, &HDCDCDC: SetPixel hdc, 8, 4, &HE3E3E3: SetPixel hdc, 9, 4, &HE3E3E3: SetPixel hdc, 10, 4, &HEAEAEA: SetPixel hdc, 11, 4, &HEBEBEB: SetPixel hdc, 12, 4, &HEBEBEB: SetPixel hdc, 13, 4, &HEBEBEB: SetPixel hdc, 14, 4, &HEBEBEB: SetPixel hdc, 15, 4, &HEBEBEB: SetPixel hdc, 16, 4, &HEBEBEB: SetPixel hdc, 17, 4, &HEBEBEB:
    SetPixel hdc, 1, 5, &HFEFEFE: SetPixel hdc, 2, 5, &HCACACA: SetPixel hdc, 3, 5, &H696969: SetPixel hdc, 4, 5, &H949494: SetPixel hdc, 5, 5, &HA6A6A6: SetPixel hdc, 6, 5, &HC5C5C5: SetPixel hdc, 7, 5, &HD8D8D8: SetPixel hdc, 8, 5, &HE0E0E0: SetPixel hdc, 9, 5, &HE1E1E1: SetPixel hdc, 10, 5, &HEAE9EA: SetPixel hdc, 11, 5, &HE7E7E7: SetPixel hdc, 12, 5, &HE9E7E8: SetPixel hdc, 13, 5, &HEBE8EA: SetPixel hdc, 14, 5, &HEAE7E9: SetPixel hdc, 15, 5, &HEBE8EA: SetPixel hdc, 16, 5, &HEAE7E9: SetPixel hdc, 17, 5, &HEBE8EA
    SetPixel hdc, 1, 6, &HF9F9F9: SetPixel hdc, 2, 6, &H808080: SetPixel hdc, 3, 6, &H878787: SetPixel hdc, 4, 6, &HA8A8A8: SetPixel hdc, 5, 6, &HB3B3B3: SetPixel hdc, 6, 6, &HC6C6C6: SetPixel hdc, 7, 6, &HDEDEDE: SetPixel hdc, 8, 6, &HE0E0E0: SetPixel hdc, 9, 6, &HE2E2E2: SetPixel hdc, 10, 6, &HE3E2E2: SetPixel hdc, 11, 6, &HE9EAE9: SetPixel hdc, 12, 6, &HE9E8E9: SetPixel hdc, 13, 6, &HEBE8EA: SetPixel hdc, 14, 6, &HEBE8EA: SetPixel hdc, 15, 6, &HEBE8EA: SetPixel hdc, 16, 6, &HEBE8EA: SetPixel hdc, 17, 6, &HEBE8EA
    SetPixel hdc, 1, 7, &HE8E8E8: SetPixel hdc, 2, 7, &H777777: SetPixel hdc, 3, 7, &H9B9B9B: SetPixel hdc, 4, 7, &HB1B1B1: SetPixel hdc, 5, 7, &HB9B9B9: SetPixel hdc, 6, 7, &HC5C5C5: SetPixel hdc, 7, 7, &HD6D6D6: SetPixel hdc, 8, 7, &HE0E0E0: SetPixel hdc, 9, 7, &HE0E0E0: SetPixel hdc, 10, 7, &HE7E7E7: SetPixel hdc, 11, 7, &HE7E7E7: SetPixel hdc, 12, 7, &HE9E9E9: SetPixel hdc, 13, 7, &HEAEAEA: SetPixel hdc, 14, 7, &HEAEAEA: SetPixel hdc, 15, 7, &HEAEAEA: SetPixel hdc, 16, 7, &HEAEAEA: SetPixel hdc, 17, 7, &HEAEAEA
    SetPixel hdc, 0, 8, &HFDFDFD: SetPixel hdc, 1, 8, &HC6C6C6: SetPixel hdc, 2, 8, &H7E7E7E: SetPixel hdc, 3, 8, &HABABAB: SetPixel hdc, 4, 8, &HC1C1C1: SetPixel hdc, 5, 8, &HC1C1C1: SetPixel hdc, 6, 8, &HCBCBCB: SetPixel hdc, 7, 8, &HCECECE: SetPixel hdc, 8, 8, &HD5D5D5: SetPixel hdc, 9, 8, &HD8D8D8: SetPixel hdc, 10, 8, &HDADADA: SetPixel hdc, 11, 8, &HDDDDDD: SetPixel hdc, 12, 8, &HDEDEDE: SetPixel hdc, 13, 8, &HE1E1E1: SetPixel hdc, 14, 8, &HE0E0E0: SetPixel hdc, 15, 8, &HE1E1E1: SetPixel hdc, 16, 8, &HE0E0E0: SetPixel hdc, 17, 8, &HE1E1E1
    SetPixel hdc, 0, 9, &HFAFAFA: SetPixel hdc, 1, 9, &HAEAEAE: SetPixel hdc, 2, 9, &H919191: SetPixel hdc, 3, 9, &HB9B9B9: SetPixel hdc, 4, 9, &HC4C4C4: SetPixel hdc, 5, 9, &HCECECE: SetPixel hdc, 6, 9, &HD1D1D1: SetPixel hdc, 7, 9, &HDADADA: SetPixel hdc, 8, 9, &HDCDCDC: SetPixel hdc, 9, 9, &HDBDBDB: SetPixel hdc, 10, 9, &HDFDFDF: SetPixel hdc, 11, 9, &HE1E3E1: SetPixel hdc, 12, 9, &HE2E3E2: SetPixel hdc, 13, 9, &HE5E2E3: SetPixel hdc, 14, 9, &HE5E2E3: SetPixel hdc, 15, 9, &HE5E2E3: SetPixel hdc, 16, 9, &HE5E2E3: SetPixel hdc, 17, 9, &HE5E2E3
    SetPixel hdc, 0, 10, &HF7F7F7: SetPixel hdc, 1, 10, &HA0A0A0: SetPixel hdc, 2, 10, &H999999: SetPixel hdc, 3, 10, &HC3C3C3: SetPixel hdc, 4, 10, &HC9C9C9: SetPixel hdc, 5, 10, &HD5D5D5: SetPixel hdc, 6, 10, &HD7D7D7: SetPixel hdc, 7, 10, &HDFDFDF: SetPixel hdc, 8, 10, &HE0E0E0: SetPixel hdc, 9, 10, &HE0E0E0: SetPixel hdc, 10, 10, &HE4E4E4: SetPixel hdc, 11, 10, &HE6E8E6: SetPixel hdc, 12, 10, &HE8E7E7: SetPixel hdc, 13, 10, &HEAE7E8: SetPixel hdc, 14, 10, &HEAE7E8: SetPixel hdc, 15, 10, &HEAE7E8: SetPixel hdc, 16, 10, &HEAE7E8: SetPixel hdc, 17, 10, &HEAE7E8
    SetPixel hdc, 0, 11, &HF5F5F5: SetPixel hdc, 1, 11, &HA3A3A3: SetPixel hdc, 2, 11, &H9B9B9B: SetPixel hdc, 3, 11, &HC6C6C6: SetPixel hdc, 4, 11, &HD3D3D3: SetPixel hdc, 5, 11, &HD6D6D6: SetPixel hdc, 6, 11, &HDDDDDD: SetPixel hdc, 7, 11, &HE1E1E1: SetPixel hdc, 8, 11, &HE3E3E3: SetPixel hdc, 9, 11, &HE6E6E6: SetPixel hdc, 10, 11, &HE7E8E7: SetPixel hdc, 11, 11, &HE9EAE9: SetPixel hdc, 12, 11, &HE8EAE9: SetPixel hdc, 13, 11, &HE8EBE9: SetPixel hdc, 14, 11, &HE8EBE9: SetPixel hdc, 15, 11, &HE8EBE9: SetPixel hdc, 16, 11, &HE8EBE9: SetPixel hdc, 17, 11, &HE8EBE9
    SetPixel hdc, 0, 12, &HF5F5F5: SetPixel hdc, 1, 12, &HAAAAAA: SetPixel hdc, 2, 12, &H8E8E8E: SetPixel hdc, 3, 12, &HD0D0D0: SetPixel hdc, 4, 12, &HDADADA: SetPixel hdc, 5, 12, &HDFDFDF: SetPixel hdc, 6, 12, &HE4E4E4: SetPixel hdc, 7, 12, &HE6E6E6: SetPixel hdc, 8, 12, &HE8E8E8: SetPixel hdc, 9, 12, &HECECEC: SetPixel hdc, 10, 12, &HEEEFEE: SetPixel hdc, 11, 12, &HEEF0EF: SetPixel hdc, 12, 12, &HEEF0EF: SetPixel hdc, 13, 12, &HEEF1EF: SetPixel hdc, 14, 12, &HEEF1EF: SetPixel hdc, 15, 12, &HEEF1EF: SetPixel hdc, 16, 12, &HEEF1EF: SetPixel hdc, 17, 12, &HEEF1EF
    tmph = lh - 22
    SetPixel hdc, 0, tmph + 12, &HF5F5F5: SetPixel hdc, 1, tmph + 12, &HAAAAAA: SetPixel hdc, 2, tmph + 12, &H8E8E8E: SetPixel hdc, 3, tmph + 12, &HD0D0D0: SetPixel hdc, 4, tmph + 12, &HDADADA: SetPixel hdc, 5, tmph + 12, &HDFDFDF: SetPixel hdc, 6, tmph + 12, &HE4E4E4: SetPixel hdc, 7, tmph + 12, &HE6E6E6: SetPixel hdc, 8, tmph + 12, &HE8E8E8: SetPixel hdc, 9, tmph + 12, &HECECEC: SetPixel hdc, 10, tmph + 12, &HEEEFEE: SetPixel hdc, 11, tmph + 12, &HEEF0EF: SetPixel hdc, 12, tmph + 12, &HEEF0EF: SetPixel hdc, 13, tmph + 12, &HEEF1EF: SetPixel hdc, 14, tmph + 12, &HEEF1EF: SetPixel hdc, 15, tmph + 12, &HEEF1EF: SetPixel hdc, 16, tmph + 12, &HEEF1EF: SetPixel hdc, 17, tmph + 12, &HEEF1EF
    SetPixel hdc, 0, tmph + 13, &HF7F7F7: SetPixel hdc, 1, tmph + 13, &HC2C2C2: SetPixel hdc, 2, tmph + 13, &H838383: SetPixel hdc, 3, tmph + 13, &HCFCFCF: SetPixel hdc, 4, tmph + 13, &HDEDEDE: SetPixel hdc, 5, tmph + 13, &HE3E3E3: SetPixel hdc, 6, tmph + 13, &HE8E8E8: SetPixel hdc, 7, tmph + 13, &HEAEAEA: SetPixel hdc, 8, tmph + 13, &HEDEDED: SetPixel hdc, 9, tmph + 13, &HF1F1F1: SetPixel hdc, 10, tmph + 13, &HF2F2F2: SetPixel hdc, 11, tmph + 13, &HF2F2F2: SetPixel hdc, 12, tmph + 13, &HF2F2F2: SetPixel hdc, 13, tmph + 13, &HF2F2F2: SetPixel hdc, 14, tmph + 13, &HF2F2F2: SetPixel hdc, 15, tmph + 13, &HF2F2F2: SetPixel hdc, 16, tmph + 13, &HF2F2F2: SetPixel hdc, 17, tmph + 13, &HF2F2F2
    SetPixel hdc, 0, tmph + 14, &HFBFBFB: SetPixel hdc, 1, tmph + 14, &HE1E1E1: SetPixel hdc, 2, tmph + 14, &H818181: SetPixel hdc, 3, tmph + 14, &HABABAB: SetPixel hdc, 4, tmph + 14, &HDCDCDC: SetPixel hdc, 5, tmph + 14, &HE5E5E5: SetPixel hdc, 6, tmph + 14, &HEDEDED: SetPixel hdc, 7, tmph + 14, &HEFEFEF: SetPixel hdc, 8, tmph + 14, &HF1F1F1: SetPixel hdc, 9, tmph + 14, &HF4F4F4: SetPixel hdc, 10, tmph + 14, &HF5F5F5: SetPixel hdc, 11, tmph + 14, &HF5F5F5: SetPixel hdc, 12, tmph + 14, &HF5F5F5: SetPixel hdc, 13, tmph + 14, &HF5F5F5: SetPixel hdc, 14, tmph + 14, &HF5F5F5: SetPixel hdc, 15, tmph + 14, &HF5F5F5: SetPixel hdc, 16, tmph + 14, &HF5F5F5: SetPixel hdc, 17, tmph + 14, &HF5F5F5
    SetPixel hdc, 0, tmph + 15, &HFEFEFE: SetPixel hdc, 1, tmph + 15, &HEDEDED: SetPixel hdc, 2, tmph + 15, &HA0A0A0: SetPixel hdc, 3, tmph + 15, &H898989: SetPixel hdc, 4, tmph + 15, &HDEDEDE: SetPixel hdc, 5, tmph + 15, &HE9E9E9: SetPixel hdc, 6, tmph + 15, &HEEEEEE: SetPixel hdc, 7, tmph + 15, &HF4F4F4: SetPixel hdc, 8, tmph + 15, &HF5F5F5: SetPixel hdc, 9, tmph + 15, &HFAFAFA: SetPixel hdc, 10, tmph + 15, &HFFFDFD: SetPixel hdc, 11, tmph + 15, &HFFFEFE: SetPixel hdc, 12, tmph + 15, &HFFFDFD: SetPixel hdc, 13, tmph + 15, &HFFFEFE: SetPixel hdc, 14, tmph + 15, &HFFFDFD: SetPixel hdc, 15, tmph + 15, &HFFFEFE: SetPixel hdc, 16, tmph + 15, &HFFFDFD: SetPixel hdc, 17, tmph + 15, &HFFFEFE
    SetPixel hdc, 1, tmph + 16, &HF6F6F6: SetPixel hdc, 2, tmph + 16, &HD6D6D6: SetPixel hdc, 3, tmph + 16, &H7B7B7B: SetPixel hdc, 4, tmph + 16, &H8D8D8D: SetPixel hdc, 5, tmph + 16, &HE4E4E4: SetPixel hdc, 6, tmph + 16, &HF0F0F0: SetPixel hdc, 7, tmph + 16, &HF6F6F6: SetPixel hdc, 8, tmph + 16, &HFEFEFE: SetPixel hdc, 9, tmph + 16, &HFEFEFE: SetPixel hdc, 10, tmph + 16, &HFFFEFE: SetPixel hdc, 12, tmph + 16, &HFFFEFE: SetPixel hdc, 14, tmph + 16, &HFFFEFE: SetPixel hdc, 16, tmph + 16, &HFFFEFE
    SetPixel hdc, 1, tmph + 17, &HFDFDFD: SetPixel hdc, 2, tmph + 17, &HEDEDED: SetPixel hdc, 3, tmph + 17, &HBEBEBE: SetPixel hdc, 4, tmph + 17, &H727272: SetPixel hdc, 5, tmph + 17, &H898989: SetPixel hdc, 6, tmph + 17, &HEBEBEB: SetPixel hdc, 7, tmph + 17, &HF5F5F5: SetPixel hdc, 8, tmph + 17, &HFCFCFC: SetPixel hdc, 10, tmph + 17, &HFDFDFD: SetPixel hdc, 11, tmph + 17, &HFDFDFD: SetPixel hdc, 12, tmph + 17, &HFDFDFD: SetPixel hdc, 13, tmph + 17, &HFDFDFD: SetPixel hdc, 14, tmph + 17, &HFDFDFD: SetPixel hdc, 15, tmph + 17, &HFDFDFD: SetPixel hdc, 16, tmph + 17, &HFDFDFD: SetPixel hdc, 17, tmph + 17, &HFDFDFD
    SetPixel hdc, 2, tmph + 18, &HF9F9F9: SetPixel hdc, 3, tmph + 18, &HE6E6E6: SetPixel hdc, 4, tmph + 18, &HB9B9B9: SetPixel hdc, 5, tmph + 18, &H717171: SetPixel hdc, 6, tmph + 18, &H787878: SetPixel hdc, 7, tmph + 18, &HB6B6B6: SetPixel hdc, 8, tmph + 18, &HF7F7F7: SetPixel hdc, 9, tmph + 18, &HFCFCFC: SetPixel hdc, 10, tmph + 18, &HFEFEFE: SetPixel hdc, 11, tmph + 18, &HFEFEFE: SetPixel hdc, 12, tmph + 18, &HFEFEFE: SetPixel hdc, 13, tmph + 18, &HFEFEFE: SetPixel hdc, 14, tmph + 18, &HFEFEFE: SetPixel hdc, 15, tmph + 18, &HFEFEFE: SetPixel hdc, 16, tmph + 18, &HFEFEFE: SetPixel hdc, 17, tmph + 18, &HFEFEFE
    SetPixel hdc, 2, tmph + 19, &HFEFEFE: SetPixel hdc, 3, tmph + 19, &HF8F8F8: SetPixel hdc, 4, tmph + 19, &HE6E6E6: SetPixel hdc, 5, tmph + 19, &HC8C8C8: SetPixel hdc, 6, tmph + 19, &H8E8E8E: SetPixel hdc, 7, tmph + 19, &H6C6C6C: SetPixel hdc, 8, tmph + 19, &H757575: SetPixel hdc, 9, tmph + 19, &H9F9F9F: SetPixel hdc, 10, tmph + 19, &HC7C7C7: SetPixel hdc, 11, tmph + 19, &HE9E9E9: SetPixel hdc, 12, tmph + 19, &HFBFBFB: SetPixel hdc, 13, tmph + 19, &HFBFBFB: SetPixel hdc, 14, tmph + 19, &HFBFBFB: SetPixel hdc, 15, tmph + 19, &HFBFBFB: SetPixel hdc, 16, tmph + 19, &HFBFBFB: SetPixel hdc, 17, tmph + 19, &HFBFBFB
    SetPixel hdc, 3, tmph + 20, &HFEFEFE: SetPixel hdc, 4, tmph + 20, &HF9F9F9: SetPixel hdc, 5, tmph + 20, &HECECEC: SetPixel hdc, 6, tmph + 20, &HDADADA: SetPixel hdc, 7, tmph + 20, &HC1C1C1: SetPixel hdc, 8, tmph + 20, &H9D9D9D: SetPixel hdc, 9, tmph + 20, &H7B7B7B: SetPixel hdc, 10, tmph + 20, &H5E5E5E: SetPixel hdc, 11, tmph + 20, &H535353: SetPixel hdc, 12, tmph + 20, &H4D4D4D: SetPixel hdc, 13, tmph + 20, &H4B4B4B: SetPixel hdc, 14, tmph + 20, &H505050: SetPixel hdc, 15, tmph + 20, &H525252: SetPixel hdc, 16, tmph + 20, &H555555: SetPixel hdc, 17, tmph + 20, &H545454
    SetPixel hdc, 5, tmph + 21, &HFCFCFC: SetPixel hdc, 6, tmph + 21, &HF5F5F5: SetPixel hdc, 7, tmph + 21, &HEBEBEB: SetPixel hdc, 8, tmph + 21, &HE1E1E1: SetPixel hdc, 9, tmph + 21, &HD6D6D6: SetPixel hdc, 10, tmph + 21, &HCECECE: SetPixel hdc, 11, tmph + 21, &HC9C9C9: SetPixel hdc, 12, tmph + 21, &HC7C7C7: SetPixel hdc, 13, tmph + 21, &HC7C7C7: SetPixel hdc, 14, tmph + 21, &HC6C6C6: SetPixel hdc, 15, tmph + 21, &HC6C6C6: SetPixel hdc, 16, tmph + 21, &HC5C5C5: SetPixel hdc, 17, tmph + 21, &HC5C5C5
    SetPixel hdc, 7, tmph + 22, &HFDFDFD: SetPixel hdc, 8, tmph + 22, &HF9F9F9: SetPixel hdc, 9, tmph + 22, &HF4F4F4: SetPixel hdc, 10, tmph + 22, &HF0F0F0: SetPixel hdc, 11, tmph + 22, &HEEEEEE: SetPixel hdc, 12, tmph + 22, &HEDEDED: SetPixel hdc, 13, tmph + 22, &HECECEC: SetPixel hdc, 14, tmph + 22, &HECECEC: SetPixel hdc, 15, tmph + 22, &HECECEC: SetPixel hdc, 16, tmph + 22, &HECECEC: SetPixel hdc, 17, tmph + 22, &HECECEC
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, 0, &H67696A: SetPixel hdc, tmpw + 18, 0, &H666869: SetPixel hdc, tmpw + 19, 0, &H716F6F: SetPixel hdc, tmpw + 20, 0, &H6F6D6D: SetPixel hdc, tmpw + 21, 0, &H6F706E: SetPixel hdc, tmpw + 22, 0, &H727371: SetPixel hdc, tmpw + 23, 0, &H6E6E6E: SetPixel hdc, tmpw + 24, 0, &H707070: SetPixel hdc, tmpw + 25, 0, &HA6A6A6: SetPixel hdc, tmpw + 26, 0, &HEEEEEE: SetPixel hdc, tmpw + 34, 0, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 1, &HF5F4F6: SetPixel hdc, tmpw + 18, 1, &HF5F4F6: SetPixel hdc, tmpw + 19, 1, &HF5F4F6: SetPixel hdc, tmpw + 20, 1, &HF5F4F6: SetPixel hdc, tmpw + 21, 1, &HF4F3F5: SetPixel hdc, tmpw + 22, 1, &HF1F0F2: SetPixel hdc, tmpw + 23, 1, &HE0E0E0: SetPixel hdc, tmpw + 24, 1, &HC3C3C3: SetPixel hdc, tmpw + 25, 1, &H848484: SetPixel hdc, tmpw + 26, 1, &H6B6B6B: SetPixel hdc, tmpw + 27, 1, &HA0A0A0: SetPixel hdc, tmpw + 28, 1, &HF7F7F7: SetPixel hdc, tmpw + 34, 1, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 2, &HF3F2F4: SetPixel hdc, tmpw + 18, 2, &HF2F1F3: SetPixel hdc, tmpw + 19, 2, &HF3F2F4: SetPixel hdc, tmpw + 20, 2, &HF3F2F4: SetPixel hdc, tmpw + 21, 2, &HF0EFF1: SetPixel hdc, tmpw + 22, 2, &HF2F1F3: SetPixel hdc, tmpw + 23, 2, &HF6F6F6: SetPixel hdc, tmpw + 24, 2, &HE8E8E8: SetPixel hdc, tmpw + 25, 2, &HE0E0E0: SetPixel hdc, tmpw + 26, 2, &H999999: SetPixel hdc, tmpw + 27, 2, &H696969: SetPixel hdc, tmpw + 28, 2, &H717171: SetPixel hdc, tmpw + 29, 2, &HEBEBEB: SetPixel hdc, tmpw + 34, 2, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 3, &HEEEEEE: SetPixel hdc, tmpw + 18, 3, &HEDEDED: SetPixel hdc, tmpw + 19, 3, &HEEEEEE: SetPixel hdc, tmpw + 20, 3, &HEEEEEE: SetPixel hdc, tmpw + 21, 3, &HEEEEEE: SetPixel hdc, tmpw + 22, 3, &HEEEEEE: SetPixel hdc, tmpw + 23, 3, &HE9E9E9: SetPixel hdc, tmpw + 24, 3, &HEAEAEA: SetPixel hdc, tmpw + 25, 3, &HE7E7E7: SetPixel hdc, tmpw + 26, 3, &HD0D0D0: SetPixel hdc, tmpw + 27, 3, &H939393: SetPixel hdc, tmpw + 28, 3, &H727272: SetPixel hdc, tmpw + 29, 3, &H6F6F6F: SetPixel hdc, tmpw + 30, 3, &HEFEFEF: SetPixel hdc, tmpw + 34, 3, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 4, &HEBEBEB: SetPixel hdc, tmpw + 18, 4, &HEBEBEB: SetPixel hdc, tmpw + 19, 4, &HEBEBEB: SetPixel hdc, tmpw + 20, 4, &HEBEBEB: SetPixel hdc, tmpw + 21, 4, &HEDEDED: SetPixel hdc, tmpw + 22, 4, &HE6E6E6: SetPixel hdc, tmpw + 23, 4, &HE9E9E9: SetPixel hdc, tmpw + 24, 4, &HE6E6E6: SetPixel hdc, tmpw + 25, 4, &HDEDEDE: SetPixel hdc, tmpw + 26, 4, &HDCDCDC: SetPixel hdc, tmpw + 27, 4, &HB2B2B2: SetPixel hdc, tmpw + 28, 4, &H919191: SetPixel hdc, tmpw + 29, 4, &H6E6E6E: SetPixel hdc, tmpw + 30, 4, &H7F7F7F: SetPixel hdc, tmpw + 31, 4, &HFAFAFA: SetPixel hdc, tmpw + 34, 4, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 5, &HEBE8EA: SetPixel hdc, tmpw + 18, 5, &HEAE7E9: SetPixel hdc, tmpw + 19, 5, &HEBE8EA: SetPixel hdc, tmpw + 20, 5, &HEBE8EA: SetPixel hdc, tmpw + 21, 5, &HE5E8E6: SetPixel hdc, tmpw + 22, 5, &HE7EAE8: SetPixel hdc, tmpw + 23, 5, &HE5E5E5: SetPixel hdc, tmpw + 24, 5, &HE3E3E3: SetPixel hdc, tmpw + 25, 5, &HDFDFDF: SetPixel hdc, tmpw + 26, 5, &HDCDCDC: SetPixel hdc, tmpw + 27, 5, &HC3C3C3: SetPixel hdc, tmpw + 28, 5, &HA7A7A7: SetPixel hdc, tmpw + 29, 5, &H969696: SetPixel hdc, tmpw + 30, 5, &H717171: SetPixel hdc, tmpw + 31, 5, &HC5C5C5: SetPixel hdc, tmpw + 32, 5, &HFEFEFE: SetPixel hdc, tmpw + 34, 5, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 6, &HEBE8EA: SetPixel hdc, tmpw + 18, 6, &HEBE8EA: SetPixel hdc, tmpw + 19, 6, &HEBE8EA: SetPixel hdc, tmpw + 20, 6, &HEBE8EA: SetPixel hdc, tmpw + 21, 6, &HE8EBE9: SetPixel hdc, tmpw + 22, 6, &HE3E6E4: SetPixel hdc, tmpw + 23, 6, &HE5E5E5: SetPixel hdc, tmpw + 24, 6, &HE2E2E2: SetPixel hdc, tmpw + 25, 6, &HE0E0E0: SetPixel hdc, tmpw + 26, 6, &HDADADA: SetPixel hdc, tmpw + 27, 6, &HC7C7C7: SetPixel hdc, tmpw + 28, 6, &HB5B5B5: SetPixel hdc, tmpw + 29, 6, &HA6A6A6: SetPixel hdc, tmpw + 30, 6, &H8C8C8C: SetPixel hdc, tmpw + 31, 6, &H808080: SetPixel hdc, tmpw + 32, 6, &HF8F8F8: SetPixel hdc, tmpw + 34, 6, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 7, &HEAEAEA: SetPixel hdc, tmpw + 18, 7, &HEAEAEA: SetPixel hdc, tmpw + 19, 7, &HEAEAEA: SetPixel hdc, tmpw + 20, 7, &HEAEAEA: SetPixel hdc, tmpw + 21, 7, &HE9E6E8: SetPixel hdc, tmpw + 22, 7, &HE9E6E8: SetPixel hdc, tmpw + 23, 7, &HE4E4E4: SetPixel hdc, tmpw + 24, 7, &HE2E2E2: SetPixel hdc, tmpw + 25, 7, &HDFDFDF: SetPixel hdc, tmpw + 26, 7, &HD7D7D7: SetPixel hdc, tmpw + 27, 7, &HC4C4C4: SetPixel hdc, tmpw + 28, 7, &HB7B7B7: SetPixel hdc, tmpw + 29, 7, &HB4B5B3: SetPixel hdc, tmpw + 30, 7, &H9D9E9C: SetPixel hdc, tmpw + 31, 7, &H777777: SetPixel hdc, tmpw + 32, 7, &HE7E7E7: SetPixel hdc, tmpw + 34, 7, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 8, &HE1E1E1: SetPixel hdc, tmpw + 18, 8, &HE0E0E0: SetPixel hdc, tmpw + 19, 8, &HE1E1E1: SetPixel hdc, tmpw + 20, 8, &HE1E1E1: SetPixel hdc, tmpw + 21, 8, &HDFDCDE: SetPixel hdc, tmpw + 22, 8, &HDDDADC: SetPixel hdc, tmpw + 23, 8, &HDBDBDB: SetPixel hdc, tmpw + 24, 8, &HD6D6D6: SetPixel hdc, tmpw + 25, 8, &HD5D5D5: SetPixel hdc, tmpw + 26, 8, &HD1D1D1: SetPixel hdc, tmpw + 27, 8, &HC9C9C9: SetPixel hdc, tmpw + 28, 8, &HC4C4C4: SetPixel hdc, tmpw + 29, 8, &HC0C1BF: SetPixel hdc, tmpw + 30, 8, &HAFB0AE: SetPixel hdc, tmpw + 31, 8, &H818181: SetPixel hdc, tmpw + 32, 8, &HC3C3C3: SetPixel hdc, tmpw + 33, 8, &HFDFDFD: SetPixel hdc, tmpw + 34, 8, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 9, &HE5E2E3: SetPixel hdc, tmpw + 18, 9, &HE5E2E3: SetPixel hdc, tmpw + 19, 9, &HE5E2E3: SetPixel hdc, tmpw + 20, 9, &HE5E2E3: SetPixel hdc, tmpw + 21, 9, &HE1E1E1: SetPixel hdc, tmpw + 22, 9, &HE1E1E1: SetPixel hdc, tmpw + 23, 9, &HE1E1E1: SetPixel hdc, tmpw + 24, 9, &HDDDDDD: SetPixel hdc, tmpw + 25, 9, &HDBDBDB: SetPixel hdc, tmpw + 26, 9, &HD8D8D8: SetPixel hdc, tmpw + 27, 9, &HD2D2D2: SetPixel hdc, tmpw + 28, 9, &HCBCBCB: SetPixel hdc, tmpw + 29, 9, &HC4C4C4: SetPixel hdc, tmpw + 30, 9, &HBABABA: SetPixel hdc, tmpw + 31, 9, &H989898: SetPixel hdc, tmpw + 32, 9, &HA6A6A6: SetPixel hdc, tmpw + 33, 9, &HF9F9F9: SetPixel hdc, tmpw + 34, 9, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 10, &HEAE7E8: SetPixel hdc, tmpw + 18, 10, &HEAE7E8: SetPixel hdc, tmpw + 19, 10, &HEAE7E8: SetPixel hdc, tmpw + 20, 10, &HEAE7E8: SetPixel hdc, tmpw + 21, 10, &HE7E7E7: SetPixel hdc, tmpw + 22, 10, &HE6E6E6: SetPixel hdc, tmpw + 23, 10, &HE4E4E4: SetPixel hdc, tmpw + 24, 10, &HE0E0E0: SetPixel hdc, tmpw + 25, 10, &HE0E0E0: SetPixel hdc, tmpw + 26, 10, &HDEDEDE: SetPixel hdc, tmpw + 27, 10, &HD9D9D9: SetPixel hdc, tmpw + 28, 10, &HD3D3D3: SetPixel hdc, tmpw + 29, 10, &HCCCCCC: SetPixel hdc, tmpw + 30, 10, &HC3C3C3: SetPixel hdc, tmpw + 31, 10, &HA3A3A3: SetPixel hdc, tmpw + 32, 10, &H9C9C9C: SetPixel hdc, tmpw + 33, 10, &HF6F6F6: SetPixel hdc, tmpw + 34, 10, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 11, &HE8EBE9: SetPixel hdc, tmpw + 18, 11, &HE8EBE9: SetPixel hdc, tmpw + 19, 11, &HE8EBE9: SetPixel hdc, tmpw + 20, 11, &HE8EBE9: SetPixel hdc, tmpw + 21, 11, &HE9EAE8: SetPixel hdc, tmpw + 22, 11, &HE8E9E7: SetPixel hdc, tmpw + 23, 11, &HE9E9E9: SetPixel hdc, tmpw + 24, 11, &HE5E5E5: SetPixel hdc, tmpw + 25, 11, &HE4E4E4: SetPixel hdc, tmpw + 26, 11, &HE2E2E2: SetPixel hdc, tmpw + 27, 11, &HDBDBDB: SetPixel hdc, tmpw + 28, 11, &HD9D9D9: SetPixel hdc, tmpw + 29, 11, &HD1D1D1: SetPixel hdc, tmpw + 30, 11, &HC8C8C8: SetPixel hdc, tmpw + 31, 11, &HA4A4A4: SetPixel hdc, tmpw + 32, 11, &HA2A2A2: SetPixel hdc, tmpw + 33, 11, &HF4F4F4: SetPixel hdc, tmpw + 34, 11, &HFFFFFFFF
    SetPixel hdc, tmpw + 17, 12, &HEEF1EF: SetPixel hdc, tmpw + 18, 12, &HEEF1EF: SetPixel hdc, tmpw + 19, 12, &HEEF1EF: SetPixel hdc, tmpw + 20, 12, &HEEF1EF: SetPixel hdc, tmpw + 21, 12, &HEEEFED: SetPixel hdc, tmpw + 22, 12, &HEFF0EE: SetPixel hdc, tmpw + 23, 12, &HEEEEEE: SetPixel hdc, tmpw + 24, 12, &HECECEC: SetPixel hdc, tmpw + 25, 12, &HEAEAEA: SetPixel hdc, tmpw + 26, 12, &HE7E7E7: SetPixel hdc, tmpw + 27, 12, &HE2E2E2: SetPixel hdc, tmpw + 28, 12, &HDFDFDF: SetPixel hdc, tmpw + 29, 12, &HD8D8D8: SetPixel hdc, tmpw + 30, 12, &HD4D4D4: SetPixel hdc, tmpw + 31, 12, &H999999: SetPixel hdc, tmpw + 32, 12, &HAFAFAF: SetPixel hdc, tmpw + 33, 12, &HF5F5F5: SetPixel hdc, tmpw + 34, 12, &HFFFFFFFF
    tmph = lh - 22
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 18, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 19, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 20, tmph + 12, &HEEF1EF: SetPixel hdc, tmpw + 21, tmph + 12, &HEEEFED: SetPixel hdc, tmpw + 22, tmph + 12, &HEFF0EE: SetPixel hdc, tmpw + 23, tmph + 12, &HEEEEEE: SetPixel hdc, tmpw + 24, tmph + 12, &HECECEC: SetPixel hdc, tmpw + 25, tmph + 12, &HEAEAEA: SetPixel hdc, tmpw + 26, tmph + 12, &HE7E7E7: SetPixel hdc, tmpw + 27, tmph + 12, &HE2E2E2: SetPixel hdc, tmpw + 28, tmph + 12, &HDFDFDF: SetPixel hdc, tmpw + 29, tmph + 12, &HD8D8D8: SetPixel hdc, tmpw + 30, tmph + 12, &HD4D4D4: SetPixel hdc, tmpw + 31, tmph + 12, &H999999: SetPixel hdc, tmpw + 32, tmph + 12, &HAFAFAF: SetPixel hdc, tmpw + 33, tmph + 12, &HF5F5F5
    SetPixel hdc, tmpw + 17, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 18, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 19, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 20, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 21, tmph + 13, &HF5F4F6: SetPixel hdc, tmpw + 22, tmph + 13, &HF0EFF1: SetPixel hdc, tmpw + 23, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 24, tmph + 13, &HF2F2F2: SetPixel hdc, tmpw + 25, tmph + 13, &HECECEC: SetPixel hdc, tmpw + 26, tmph + 13, &HEAEAEA: SetPixel hdc, tmpw + 27, tmph + 13, &HEBEBEB: SetPixel hdc, tmpw + 28, tmph + 13, &HE3E3E3: SetPixel hdc, tmpw + 29, tmph + 13, &HDEDEDE: SetPixel hdc, tmpw + 30, tmph + 13, &HD1D1D1: SetPixel hdc, tmpw + 31, tmph + 13, &H8A8A8A: SetPixel hdc, tmpw + 32, tmph + 13, &HD5D5D5: SetPixel hdc, tmpw + 33, tmph + 13, &HF8F8F8
    SetPixel hdc, tmpw + 17, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 18, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 19, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 20, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 21, tmph + 14, &HF8F7F9: SetPixel hdc, tmpw + 22, tmph + 14, &HF7F6F8: SetPixel hdc, tmpw + 23, tmph + 14, &HF7F7F7: SetPixel hdc, tmpw + 24, tmph + 14, &HF5F5F5: SetPixel hdc, tmpw + 25, tmph + 14, &HEFEFEF: SetPixel hdc, tmpw + 26, tmph + 14, &HEEEEEE: SetPixel hdc, tmpw + 27, tmph + 14, &HECECEC: SetPixel hdc, tmpw + 28, tmph + 14, &HE5E5E5: SetPixel hdc, tmpw + 29, tmph + 14, &HDEDEDE: SetPixel hdc, tmpw + 30, tmph + 14, &HB3B3B3: SetPixel hdc, tmpw + 31, tmph + 14, &H808080: SetPixel hdc, tmpw + 32, tmph + 14, &HE8E8E8: SetPixel hdc, tmpw + 33, tmph + 14, &HFDFDFD
    SetPixel hdc, tmpw + 17, tmph + 15, &HFFFEFE: SetPixel hdc, tmpw + 18, tmph + 15, &HFFFDFD: SetPixel hdc, tmpw + 19, tmph + 15, &HFFFEFE: SetPixel hdc, tmpw + 20, tmph + 15, &HFFFEFE: SetPixel hdc, tmpw + 21, tmph + 15, &HFBFBFB: SetPixel hdc, tmpw + 22, tmph + 15, &HFCFCFC: SetPixel hdc, tmpw + 23, tmph + 15, &HFEFEFE: SetPixel hdc, tmpw + 24, tmph + 15, &HF8F8F8: SetPixel hdc, tmpw + 25, tmph + 15, &HF7F7F7: SetPixel hdc, tmpw + 26, tmph + 15, &HF5F5F5: SetPixel hdc, tmpw + 27, tmph + 15, &HEDEDED: SetPixel hdc, tmpw + 28, tmph + 15, &HEAEAEA: SetPixel hdc, tmpw + 29, tmph + 15, &HE0E0E0: SetPixel hdc, tmpw + 30, tmph + 15, &H8D8D8D: SetPixel hdc, tmpw + 31, tmph + 15, &HBABABA: SetPixel hdc, tmpw + 32, tmph + 15, &HF1F1F1
    SetPixel hdc, tmpw + 18, tmph + 16, &HFFFEFE: SetPixel hdc, tmpw + 22, tmph + 16, &HFEFEFE: SetPixel hdc, tmpw + 23, tmph + 16, &HFEFEFE: SetPixel hdc, tmpw + 25, tmph + 16, &HFCFCFC: SetPixel hdc, tmpw + 26, tmph + 16, &HF6F6F6: SetPixel hdc, tmpw + 27, tmph + 16, &HF2F2F2: SetPixel hdc, tmpw + 28, tmph + 16, &HE7E7E7: SetPixel hdc, tmpw + 29, tmph + 16, &H989898: SetPixel hdc, tmpw + 30, tmph + 16, &H828282: SetPixel hdc, tmpw + 31, tmph + 16, &HE2E2E2: SetPixel hdc, tmpw + 32, tmph + 16, &HF9F9F9
    SetPixel hdc, tmpw + 17, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 18, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 19, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 20, tmph + 17, &HFDFDFD: SetPixel hdc, tmpw + 21, tmph + 17, &HFEFEFE: SetPixel hdc, tmpw + 23, tmph + 17, &HFEFEFE: SetPixel hdc, tmpw + 25, tmph + 17, &HFEFEFE: SetPixel hdc, tmpw + 26, tmph + 17, &HF6F6F6: SetPixel hdc, tmpw + 27, tmph + 17, &HF1F1F1: SetPixel hdc, tmpw + 28, tmph + 17, &H979797: SetPixel hdc, tmpw + 29, tmph + 17, &H6F6F6F: SetPixel hdc, tmpw + 30, tmph + 17, &HD2D2D2: SetPixel hdc, tmpw + 31, tmph + 17, &HF2F2F2: SetPixel hdc, tmpw + 32, tmph + 17, &HFEFEFE
    SetPixel hdc, tmpw + 17, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 18, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 19, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 20, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 22, tmph + 18, &HFDFDFD: SetPixel hdc, tmpw + 23, tmph + 18, &HFEFEFE: SetPixel hdc, tmpw + 24, tmph + 18, &HFDFDFD: SetPixel hdc, tmpw + 25, tmph + 18, &HFCFCFC: SetPixel hdc, tmpw + 26, tmph + 18, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 18, &H838383: SetPixel hdc, tmpw + 28, tmph + 18, &H6F6F6F: SetPixel hdc, tmpw + 29, tmph + 18, &HC8C8C8: SetPixel hdc, tmpw + 30, tmph + 18, &HEBEBEB: SetPixel hdc, tmpw + 31, tmph + 18, &HFCFCFC
    SetPixel hdc, tmpw + 17, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 18, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 19, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 20, tmph + 19, &HFBFBFB: SetPixel hdc, tmpw + 21, tmph + 19, &HFAFAFA: SetPixel hdc, tmpw + 22, tmph + 19, &HEFEFEF: SetPixel hdc, tmpw + 23, tmph + 19, &HD0D0D0: SetPixel hdc, tmpw + 24, tmph + 19, &HA3A3A3: SetPixel hdc, tmpw + 25, tmph + 19, &H7E7E7E: SetPixel hdc, tmpw + 26, tmph + 19, &H6A6A6A: SetPixel hdc, tmpw + 27, tmph + 19, &H8F8F8F: SetPixel hdc, tmpw + 28, tmph + 19, &HCDCDCD: SetPixel hdc, tmpw + 29, tmph + 19, &HE8E8E8: SetPixel hdc, tmpw + 30, tmph + 19, &HFAFAFA
    SetPixel hdc, tmpw + 17, tmph + 20, &H545454: SetPixel hdc, tmpw + 18, tmph + 20, &H555555: SetPixel hdc, tmpw + 19, tmph + 20, &H525252: SetPixel hdc, tmpw + 20, tmph + 20, &H505050: SetPixel hdc, tmpw + 21, tmph + 20, &H535353: SetPixel hdc, tmpw + 22, tmph + 20, &H525252: SetPixel hdc, tmpw + 23, tmph + 20, &H616161: SetPixel hdc, tmpw + 24, tmph + 20, &H7A7A7A: SetPixel hdc, tmpw + 25, tmph + 20, &HA3A3A3: SetPixel hdc, tmpw + 26, tmph + 20, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 20, &HDADADA: SetPixel hdc, tmpw + 28, tmph + 20, &HEDEDED: SetPixel hdc, tmpw + 29, tmph + 20, &HFAFAFA
    SetPixel hdc, tmpw + 17, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 18, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 19, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 20, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 21, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 22, tmph + 21, &HC9C9C9: SetPixel hdc, tmpw + 23, tmph + 21, &HCECECE: SetPixel hdc, tmpw + 24, tmph + 21, &HD7D7D7: SetPixel hdc, tmpw + 25, tmph + 21, &HE1E1E1: SetPixel hdc, tmpw + 26, tmph + 21, &HECECEC: SetPixel hdc, tmpw + 27, tmph + 21, &HF6F6F6: SetPixel hdc, tmpw + 28, tmph + 21, &HFDFDFD
    SetPixel hdc, tmpw + 17, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 18, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 19, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 20, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 21, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 22, tmph + 22, &HEDEDED: SetPixel hdc, tmpw + 23, tmph + 22, &HF0F0F0: SetPixel hdc, tmpw + 24, tmph + 22, &HF4F4F4: SetPixel hdc, tmpw + 25, tmph + 22, &HFAFAFA: SetPixel hdc, tmpw + 26, tmph + 22, &HFDFDFD
    'Vlines
    tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
    DrawLineApi 0, tmph, 0, tmph1, &HF7F7F7: DrawLineApi 1, tmph, 1, tmph1, &HA0A0A0: DrawLineApi 2, tmph, 2, tmph1, &H999999: DrawLineApi 3, tmph, 3, tmph1, &HC3C3C3
    DrawLineApi 4, tmph, 4, tmph1, &HC9C9C9: DrawLineApi 5, tmph, 5, tmph1, &HD5D5D5: DrawLineApi 6, tmph, 6, tmph1, &HD7D7D7: DrawLineApi 7, tmph, 7, tmph1, &HDFDFDF
    DrawLineApi 8, tmph, 8, tmph1, &HE0E0E0: DrawLineApi 9, tmph, 9, tmph1, &HE0E0E0: DrawLineApi 10, tmph, 10, tmph1, &HE4E4E4: DrawLineApi 11, tmph, 11, tmph1, &HE6E8E6
    DrawLineApi 12, tmph, 12, tmph1, &HE8E7E7: DrawLineApi 13, tmph, 13, tmph1, &HEAE7E8: DrawLineApi 14, tmph, 14, tmph1, &HEAE7E8: DrawLineApi 15, tmph, 15, tmph1, &HEAE7E8
    DrawLineApi 16, tmph, 16, tmph1, &HEAE7E8: DrawLineApi 17, tmph, 17, tmph1, &HEAE7E8: DrawLineApi tmpw + 17, tmph, tmpw + 17, tmph1, &HEAE7E8: DrawLineApi tmpw + 18, tmph, tmpw + 18, tmph1, &HEAE7E8
    DrawLineApi tmpw + 19, tmph, tmpw + 19, tmph1, &HEAE7E8: DrawLineApi tmpw + 20, tmph, tmpw + 20, tmph1, &HEAE7E8: DrawLineApi tmpw + 21, tmph, tmpw + 21, tmph1, &HE7E7E7
    DrawLineApi tmpw + 22, tmph, tmpw + 22, tmph1, &HE6E6E6: DrawLineApi tmpw + 23, tmph, tmpw + 23, tmph1, &HE4E4E4: DrawLineApi tmpw + 24, tmph, tmpw + 24, tmph1, &HE0E0E0
    DrawLineApi tmpw + 25, tmph, tmpw + 25, tmph1, &HE0E0E0: DrawLineApi tmpw + 26, tmph, tmpw + 26, tmph1, &HDEDEDE: DrawLineApi tmpw + 27, tmph, tmpw + 27, tmph1, &HD9D9D9
    DrawLineApi tmpw + 28, tmph, tmpw + 28, tmph1, &HD3D3D3: DrawLineApi tmpw + 29, tmph, tmpw + 29, tmph1, &HCCCCCC: DrawLineApi tmpw + 30, tmph, tmpw + 30, tmph1, &HC3C3C3
    DrawLineApi tmpw + 31, tmph, tmpw + 31, tmph1, &HA3A3A3: DrawLineApi tmpw + 32, tmph, tmpw + 32, tmph1, &H9C9C9C: DrawLineApi tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6
    'HLines
    DrawLineApi 17, 0, lw - 17, 0, &H67696A
    DrawLineApi 17, 1, lw - 17, 1, &HF5F4F6
    DrawLineApi 17, 2, lw - 17, 2, &HF3F2F4
    DrawLineApi 17, 3, lw - 17, 3, &HEEEEEE
    DrawLineApi 17, 4, lw - 17, 4, &HEBEBEB
    DrawLineApi 17, 5, lw - 17, 5, &HEBE8EA
    DrawLineApi 17, 6, lw - 17, 6, &HEBE8EA
    DrawLineApi 17, 7, lw - 17, 7, &HEAEAEA
    DrawLineApi 17, 8, lw - 17, 8, &HE1E1E1
    DrawLineApi 17, 9, lw - 17, 9, &HE5E2E3
    DrawLineApi 17, 10, lw - 17, 10, &HEAE7E8
    DrawLineApi 17, 11, lw - 17, 11, &HE8EBE9
    tmph = lh - 22
    DrawLineApi 17, tmph + 11, lw - 17, tmph + 11, &HE8EBE9
    DrawLineApi 17, tmph + 12, lw - 17, tmph + 12, &HEEF1EF
    DrawLineApi 17, tmph + 13, lw - 17, tmph + 13, &HF2F2F2
    DrawLineApi 17, tmph + 14, lw - 17, tmph + 14, &HF5F5F5
    DrawLineApi 17, tmph + 15, lw - 17, tmph + 15, &HFFFEFE
    DrawLineApi 17, tmph + 16, lw - 17, tmph + 16, &HFFFFFF
    DrawLineApi 17, tmph + 17, lw - 17, tmph + 17, &HFDFDFD
    DrawLineApi 17, tmph + 18, lw - 17, tmph + 18, &HFEFEFE
    DrawLineApi 17, tmph + 19, lw - 17, tmph + 19, &HFBFBFB
    DrawLineApi 17, tmph + 20, lw - 17, tmph + 20, &H545454
    DrawLineApi 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    DrawLineApi 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub

Private Sub DrawAquaHot()

Dim tmph As Long, tmpw As Long
Dim tmph1 As Long, tmpw1 As Long
Dim lpRect As RECT

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    SetRect lpRect, 4, 4, lw - 4, lh - 4
    PaintRect &HE2A66A, lpRect

    SetPixel hdc, 6, 0, &HFEFEFE: SetPixel hdc, 7, 0, &HE6E5E5: SetPixel hdc, 8, 0, &HA9A5A5: SetPixel hdc, 9, 0, &H6C5E5E: SetPixel hdc, 10, 0, &H482729: SetPixel hdc, 11, 0, &H370D0C: SetPixel hdc, 12, 0, &H370706: SetPixel hdc, 13, 0, &H360605: SetPixel hdc, 14, 0, &H3A0606: SetPixel hdc, 15, 0, &H410807: SetPixel hdc, 16, 0, &H450707: SetPixel hdc, 17, 0, &H450608:
    SetPixel hdc, 5, 1, &HF0EFEF: SetPixel hdc, 6, 1, &HA38A8C: SetPixel hdc, 7, 1, &H6E342F: SetPixel hdc, 8, 1, &H661F1A: SetPixel hdc, 9, 1, &H9B6A63: SetPixel hdc, 10, 1, &HC9A29D: SetPixel hdc, 11, 1, &HE2BFBD: SetPixel hdc, 12, 1, &HE8C9C6: SetPixel hdc, 13, 1, &HEFD3CC: SetPixel hdc, 14, 1, &HEFD3CC: SetPixel hdc, 15, 1, &HF0D5C9: SetPixel hdc, 16, 1, &HF0D5C9: SetPixel hdc, 17, 1, &HF1D4C9:
    SetPixel hdc, 3, 2, &HFEFEFE: SetPixel hdc, 4, 2, &HE5E5E5: SetPixel hdc, 5, 2, &H755E5E: SetPixel hdc, 6, 2, &H41070C: SetPixel hdc, 7, 2, &H7F2D28: SetPixel hdc, 8, 2, &HEC9892: SetPixel hdc, 9, 2, &HECB6AF: SetPixel hdc, 10, 2, &HE3BBB6: SetPixel hdc, 11, 2, &HE3C0BD: SetPixel hdc, 12, 2, &HE1C2BF: SetPixel hdc, 13, 2, &HDFC3BC: SetPixel hdc, 14, 2, &HDFC3BC: SetPixel hdc, 15, 2, &HE4C9BD: SetPixel hdc, 16, 2, &HE4C9BD: SetPixel hdc, 17, 2, &HE5C8BD:
    SetPixel hdc, 3, 3, &HEEEEEE: SetPixel hdc, 4, 3, &H8A5A5A: SetPixel hdc, 5, 3, &H7A0702: SetPixel hdc, 6, 3, &H901501: SetPixel hdc, 7, 3, &HC38365: SetPixel hdc, 8, 3, &HE3B08F: SetPixel hdc, 9, 3, &HE1B394: SetPixel hdc, 10, 3, &HE5B798: SetPixel hdc, 11, 3, &HE6BC99: SetPixel hdc, 12, 3, &HE7BD9A: SetPixel hdc, 13, 3, &HE4BC99: SetPixel hdc, 14, 3, &HE7BF9C: SetPixel hdc, 15, 3, &HE9C1A1: SetPixel hdc, 16, 3, &HE8C0A1: SetPixel hdc, 17, 3, &HE8C0A1:
    SetPixel hdc, 2, 4, &HFBFBFB: SetPixel hdc, 3, 4, &H897879: SetPixel hdc, 4, 4, &H4D0909: SetPixel hdc, 5, 4, &H951905: SetPixel hdc, 6, 4, &HBF422E: SetPixel hdc, 7, 4, &HD49475: SetPixel hdc, 8, 4, &HD7A483: SetPixel hdc, 9, 4, &HDAAC8D: SetPixel hdc, 10, 4, &HDBAD8E: SetPixel hdc, 11, 4, &HD9AF8C: SetPixel hdc, 12, 4, &HDCB28F: SetPixel hdc, 13, 4, &HDDB592: SetPixel hdc, 14, 4, &HDCB491: SetPixel hdc, 15, 4, &HDFB797: SetPixel hdc, 16, 4, &HE0B898: SetPixel hdc, 17, 4, &HE0B898:
    SetPixel hdc, 1, 5, &HFEFEFE: SetPixel hdc, 2, 5, &HCDC9C9: SetPixel hdc, 3, 5, &H882517: SetPixel hdc, 4, 5, &H922100: SetPixel hdc, 5, 5, &HA13A00: SetPixel hdc, 6, 5, &HD57333: SetPixel hdc, 7, 5, &HDFA36F: SetPixel hdc, 8, 5, &HDDA876: SetPixel hdc, 9, 5, &HD8A573: SetPixel hdc, 10, 5, &HDFAE80: SetPixel hdc, 11, 5, &HDBAD7D: SetPixel hdc, 12, 5, &HDFB084: SetPixel hdc, 13, 5, &HDFB286: SetPixel hdc, 14, 5, &HDFB188: SetPixel hdc, 15, 5, &HE1B58D: SetPixel hdc, 16, 5, &HE3B58E: SetPixel hdc, 17, 5, &HE3B48E:
    SetPixel hdc, 1, 6, &HF9F9F9: SetPixel hdc, 2, 6, &H7B706E: SetPixel hdc, 3, 6, &H871405: SetPixel hdc, 4, 6, &HA5330E: SetPixel hdc, 5, 6, &HB34C0D: SetPixel hdc, 6, 6, &HD27030: SetPixel hdc, 7, 6, &HD89C68: SetPixel hdc, 8, 6, &HDAA573: SetPixel hdc, 9, 6, &HD9A674: SetPixel hdc, 10, 6, &HD9A87A: SetPixel hdc, 11, 6, &HDBAD7D: SetPixel hdc, 12, 6, &HDBAC80: SetPixel hdc, 13, 6, &HDCAF83: SetPixel hdc, 14, 6, &HDFB188: SetPixel hdc, 15, 6, &HDEB28A: SetPixel hdc, 16, 6, &HDFB18A: SetPixel hdc, 17, 6, &HE0B18B:
    SetPixel hdc, 1, 7, &HE8E8E7: SetPixel hdc, 2, 7, &H773F34: SetPixel hdc, 3, 7, &H9F2C00: SetPixel hdc, 4, 7, &HBA4B07: SetPixel hdc, 5, 7, &HC35E10: SetPixel hdc, 6, 7, &HCC7323: SetPixel hdc, 7, 7, &HDB8F46: SetPixel hdc, 8, 7, &HE8A763: SetPixel hdc, 9, 7, &HE3A76C: SetPixel hdc, 10, 7, &HE7AB70: SetPixel hdc, 11, 7, &HE8AE73: SetPixel hdc, 12, 7, &HE8AE73: SetPixel hdc, 13, 7, &HEDB17B: SetPixel hdc, 14, 7, &HEFB37D: SetPixel hdc, 15, 7, &HE9B57E: SetPixel hdc, 16, 7, &HE9B57E: SetPixel hdc, 17, 7, &HE9B47F:
    SetPixel hdc, 0, 8, &HFDFDFD: SetPixel hdc, 1, 8, &HCAC5C5: SetPixel hdc, 2, 8, &H682A1F: SetPixel hdc, 3, 8, &HB23E0C: SetPixel hdc, 4, 8, &HCC5D19: SetPixel hdc, 5, 8, &HCE691B: SetPixel hdc, 6, 8, &HCE7525: SetPixel hdc, 7, 8, &HCD8138: SetPixel hdc, 8, 8, &HC58440: SetPixel hdc, 9, 8, &HC5894E: SetPixel hdc, 10, 8, &HC98D52: SetPixel hdc, 11, 8, &HC88E53: SetPixel hdc, 12, 8, &HCC9257: SetPixel hdc, 13, 8, &HCF935D: SetPixel hdc, 14, 8, &HD0945E: SetPixel hdc, 15, 8, &HCE9963: SetPixel hdc, 16, 8, &HCE9963: SetPixel hdc, 17, 8, &HCE9963:
    SetPixel hdc, 0, 9, &HFAFAFA: SetPixel hdc, 1, 9, &HB9ADAB: SetPixel hdc, 2, 9, &H6E2B10: SetPixel hdc, 3, 9, &HB6580D: SetPixel hdc, 4, 9, &HCA6C20: SetPixel hdc, 5, 9, &HCE792B: SetPixel hdc, 6, 9, &HCE8132: SetPixel hdc, 7, 9, &HD08B42: SetPixel hdc, 8, 9, &HD3904B: SetPixel hdc, 9, 9, &HD3934C: SetPixel hdc, 10, 9, &HD89753: SetPixel hdc, 11, 9, &HDB9B5A: SetPixel hdc, 12, 9, &HDC9B5E: SetPixel hdc, 13, 9, &HDB9C60: SetPixel hdc, 14, 9, &HDB9C60: SetPixel hdc, 15, 9, &HDDA164: SetPixel hdc, 16, 9, &HDDA164: SetPixel hdc, 17, 9, &HDDA064:
    SetPixel hdc, 0, 10, &HF7F7F7: SetPixel hdc, 1, 10, &HB0A09E: SetPixel hdc, 2, 10, &H712E13: SetPixel hdc, 3, 10, &HBD5F14: SetPixel hdc, 4, 10, &HD17327: SetPixel hdc, 5, 10, &HD47F31: SetPixel hdc, 6, 10, &HD98C3D: SetPixel hdc, 7, 10, &HD9944B: SetPixel hdc, 8, 10, &HD7944F: SetPixel hdc, 9, 10, &HDC9C55: SetPixel hdc, 10, 10, &HDC9B57: SetPixel hdc, 11, 10, &HE3A362: SetPixel hdc, 12, 10, &HE3A265: SetPixel hdc, 13, 10, &HE2A367: SetPixel hdc, 14, 10, &HE0A165: SetPixel hdc, 15, 10, &HE3A66A: SetPixel hdc, 16, 10, &HE3A66A: SetPixel hdc, 17, 10, &HE2A66A
    tmph = lh - 22
    SetPixel hdc, 0, tmph + 10, &HF7F7F7: SetPixel hdc, 1, tmph + 10, &HB0A09E: SetPixel hdc, 2, tmph + 10, &H712E13: SetPixel hdc, 3, tmph + 10, &HBD5F14: SetPixel hdc, 4, tmph + 10, &HD17327: SetPixel hdc, 5, tmph + 10, &HD47F31: SetPixel hdc, 6, tmph + 10, &HD98C3D: SetPixel hdc, 7, tmph + 10, &HD9944B: SetPixel hdc, 8, tmph + 10, &HD7944F: SetPixel hdc, 9, tmph + 10, &HDC9C55: SetPixel hdc, 10, tmph + 10, &HDC9B57: SetPixel hdc, 11, tmph + 10, &HE3A362: SetPixel hdc, 12, tmph + 10, &HE3A265: SetPixel hdc, 13, tmph + 10, &HE2A367: SetPixel hdc, 14, tmph + 10, &HE0A165: SetPixel hdc, 15, tmph + 10, &HE3A66A: SetPixel hdc, 16, tmph + 10, &HE3A66A: SetPixel hdc, 17, tmph + 10, &HE2A66A:
    SetPixel hdc, 0, tmph + 11, &HF5F5F5: SetPixel hdc, 1, tmph + 11, &HACA39E: SetPixel hdc, 2, tmph + 11, &H744421: SetPixel hdc, 3, tmph + 11, &HC56F1F: SetPixel hdc, 4, tmph + 11, &HD17A2A: SetPixel hdc, 5, tmph + 11, &HD58C42: SetPixel hdc, 6, tmph + 11, &HD7914B: SetPixel hdc, 7, tmph + 11, &HDF9854: SetPixel hdc, 8, tmph + 11, &HE4A05F: SetPixel hdc, 9, tmph + 11, &HE29F66: SetPixel hdc, 10, tmph + 11, &HE4A56B: SetPixel hdc, 11, tmph + 11, &HDDA467: SetPixel hdc, 12, tmph + 11, &HE0A76A: SetPixel hdc, 13, tmph + 11, &HE2A96C: SetPixel hdc, 14, tmph + 11, &HE3A870: SetPixel hdc, 15, tmph + 11, &HE6AC76: SetPixel hdc, 16, tmph + 11, &HE6AC76: SetPixel hdc, 17, tmph + 11, &HE6AC76:
    SetPixel hdc, 0, tmph + 12, &HF5F5F5: SetPixel hdc, 1, tmph + 12, &HB1AAA7: SetPixel hdc, 2, tmph + 12, &H825533: SetPixel hdc, 3, tmph + 12, &HCF792A: SetPixel hdc, 4, tmph + 12, &HE48D3D: SetPixel hdc, 5, tmph + 12, &HDD944A: SetPixel hdc, 6, tmph + 12, &HE49E58: SetPixel hdc, 7, tmph + 12, &HEBA460: SetPixel hdc, 8, tmph + 12, &HEEAA69: SetPixel hdc, 9, tmph + 12, &HF3B077: SetPixel hdc, 10, tmph + 12, &HEEAF75: SetPixel hdc, 11, tmph + 12, &HEBB275: SetPixel hdc, 12, tmph + 12, &HEFB679: SetPixel hdc, 13, tmph + 12, &HF1B87B: SetPixel hdc, 14, tmph + 12, &HF1B67E: SetPixel hdc, 15, tmph + 12, &HF2B781: SetPixel hdc, 16, tmph + 12, &HF1B681: SetPixel hdc, 17, tmph + 12, &HF1B681:
    SetPixel hdc, 0, tmph + 13, &HF7F7F7: SetPixel hdc, 1, tmph + 13, &HC2C2C1: SetPixel hdc, 2, tmph + 13, &H6B5D4E: SetPixel hdc, 3, tmph + 13, &HC27831: SetPixel hdc, 4, tmph + 13, &HDA8E46: SetPixel hdc, 5, tmph + 13, &HE7A05C: SetPixel hdc, 6, tmph + 13, &HEAA665: SetPixel hdc, 7, tmph + 13, &HE9AF6E: SetPixel hdc, 8, tmph + 13, &HEFB377: SetPixel hdc, 9, tmph + 13, &HF3B579: SetPixel hdc, 10, tmph + 13, &HF7B97D: SetPixel hdc, 11, tmph + 13, &HF2BB7E: SetPixel hdc, 12, tmph + 13, &HF4BB83: SetPixel hdc, 13, tmph + 13, &HF5BE85: SetPixel hdc, 14, tmph + 13, &HF4BB87: SetPixel hdc, 15, tmph + 13, &HF5BE8A: SetPixel hdc, 16, tmph + 13, &HF5BD8A: SetPixel hdc, 17, tmph + 13, &HF3BD8A:
    SetPixel hdc, 0, tmph + 14, &HFBFBFB: SetPixel hdc, 1, tmph + 14, &HE1E1E1: SetPixel hdc, 2, tmph + 14, &H85796E: SetPixel hdc, 3, tmph + 14, &HB76F2B: SetPixel hdc, 4, tmph + 14, &HDE924A: SetPixel hdc, 5, tmph + 14, &HE8A15D: SetPixel hdc, 6, tmph + 14, &HF2AE6D: SetPixel hdc, 7, tmph + 14, &HF1B776: SetPixel hdc, 8, tmph + 14, &HF2B67A: SetPixel hdc, 9, tmph + 14, &HFBBD81: SetPixel hdc, 10, tmph + 14, &HFFC286: SetPixel hdc, 11, tmph + 14, &HFAC386: SetPixel hdc, 12, tmph + 14, &HFBC28A: SetPixel hdc, 13, tmph + 14, &HFAC38A: SetPixel hdc, 14, tmph + 14, &HFAC18D: SetPixel hdc, 15, tmph + 14, &HFDC592: SetPixel hdc, 16, tmph + 14, &HFDC592: SetPixel hdc, 17, tmph + 14, &HFCC592:
    SetPixel hdc, 0, tmph + 15, &HFEFEFE: SetPixel hdc, 1, tmph + 15, &HEDEDED: SetPixel hdc, 2, tmph + 15, &HA2A0A0: SetPixel hdc, 3, tmph + 15, &H816753: SetPixel hdc, 4, tmph + 15, &HC09068: SetPixel hdc, 5, tmph + 15, &HEDA55F: SetPixel hdc, 6, tmph + 15, &HFAB26C: SetPixel hdc, 7, tmph + 15, &HFCBF7D: SetPixel hdc, 8, tmph + 15, &HF7C182: SetPixel hdc, 9, tmph + 15, &HF8C38A: SetPixel hdc, 10, tmph + 15, &HFACA90: SetPixel hdc, 11, tmph + 15, &HF7CB8E: SetPixel hdc, 12, tmph + 15, &HF8CC8F: SetPixel hdc, 13, tmph + 15, &HFACC96: SetPixel hdc, 14, tmph + 15, &HF9CB95: SetPixel hdc, 15, tmph + 15, &HF9CE97: SetPixel hdc, 16, tmph + 15, &HF8CD97: SetPixel hdc, 17, tmph + 15, &HF8CE97:
    SetPixel hdc, 1, tmph + 16, &HF6F6F6: SetPixel hdc, 2, tmph + 16, &HD6D6D6: SetPixel hdc, 3, tmph + 16, &H8E7C6F: SetPixel hdc, 4, tmph + 16, &H946843: SetPixel hdc, 5, tmph + 16, &HEEA762: SetPixel hdc, 6, tmph + 16, &HFFB771: SetPixel hdc, 7, tmph + 16, &HFEC17F: SetPixel hdc, 8, tmph + 16, &HFFC98A: SetPixel hdc, 9, tmph + 16, &HFFCE95: SetPixel hdc, 10, tmph + 16, &HFBCB91: SetPixel hdc, 11, tmph + 16, &HFFD396: SetPixel hdc, 12, tmph + 16, &HFFD396: SetPixel hdc, 13, tmph + 16, &HFFD29C: SetPixel hdc, 14, tmph + 16, &HFFD39D: SetPixel hdc, 15, tmph + 16, &HFFD49E: SetPixel hdc, 16, tmph + 16, &HFFD49E: SetPixel hdc, 17, tmph + 16, &HFED59E:
    SetPixel hdc, 1, tmph + 17, &HFDFDFD: SetPixel hdc, 2, tmph + 17, &HEDEDED: SetPixel hdc, 3, tmph + 17, &HBEBEBE: SetPixel hdc, 4, tmph + 17, &H6C6C6C: SetPixel hdc, 5, tmph + 17, &H7C684F: SetPixel hdc, 6, tmph + 17, &HD1AE81: SetPixel hdc, 7, tmph + 17, &HF1C284: SetPixel hdc, 8, tmph + 17, &HFDCE90: SetPixel hdc, 9, tmph + 17, &HF8D193: SetPixel hdc, 10, tmph + 17, &HFBD899: SetPixel hdc, 11, tmph + 17, &HF5DC9E: SetPixel hdc, 12, tmph + 17, &HF8DFA1: SetPixel hdc, 13, tmph + 17, &HF8DFA1: SetPixel hdc, 14, tmph + 17, &HF8DFA1: SetPixel hdc, 15, tmph + 17, &HF8DEA3: SetPixel hdc, 16, tmph + 17, &HF7DDA3: SetPixel hdc, 17, tmph + 17, &HF7DDA3:
    SetPixel hdc, 2, tmph + 18, &HF9F9F9: SetPixel hdc, 3, tmph + 18, &HE6E6E6: SetPixel hdc, 4, tmph + 18, &HBABABA: SetPixel hdc, 5, tmph + 18, &H827666: SetPixel hdc, 6, tmph + 18, &H836743: SetPixel hdc, 7, tmph + 18, &HBE935B: SetPixel hdc, 8, tmph + 18, &HF4C78B: SetPixel hdc, 9, tmph + 18, &HFDD79A: SetPixel hdc, 10, tmph + 18, &HFFDFA0: SetPixel hdc, 11, tmph + 18, &HFBE2A4: SetPixel hdc, 12, tmph + 18, &HFFE7A9: SetPixel hdc, 13, tmph + 18, &HFFE9AB: SetPixel hdc, 14, tmph + 18, &HFFE7A9: SetPixel hdc, 15, tmph + 18, &HFFE6AC: SetPixel hdc, 16, tmph + 18, &HFFE6AD: SetPixel hdc, 17, tmph + 18, &HFFE6AD:
    SetPixel hdc, 2, tmph + 19, &HFEFEFE: SetPixel hdc, 3, tmph + 19, &HF8F8F8: SetPixel hdc, 4, tmph + 19, &HE6E6E6: SetPixel hdc, 5, tmph + 19, &HC8C8C8: SetPixel hdc, 6, tmph + 19, &H8F8F8F: SetPixel hdc, 7, tmph + 19, &H686462: SetPixel hdc, 8, tmph + 19, &H6D655E: SetPixel hdc, 9, tmph + 19, &H918472: SetPixel hdc, 10, tmph + 19, &HB3A88E: SetPixel hdc, 11, tmph + 19, &HDAD1B2: SetPixel hdc, 12, tmph + 19, &HE3DBBA: SetPixel hdc, 13, tmph + 19, &HE7E0C0: SetPixel hdc, 14, tmph + 19, &HE9E2C1: SetPixel hdc, 15, tmph + 19, &HE9E2C5: SetPixel hdc, 16, tmph + 19, &HE9E1C5: SetPixel hdc, 17, tmph + 19, &HE9E2C5:
    SetPixel hdc, 3, tmph + 20, &HFEFEFE: SetPixel hdc, 4, tmph + 20, &HF9F9F9: SetPixel hdc, 5, tmph + 20, &HECECEC: SetPixel hdc, 6, tmph + 20, &HDADADA: SetPixel hdc, 7, tmph + 20, &HC2C2C1: SetPixel hdc, 8, tmph + 20, &H9F9D9B: SetPixel hdc, 9, tmph + 20, &H827D75: SetPixel hdc, 10, tmph + 20, &H6A6353: SetPixel hdc, 11, tmph + 20, &H5F5941: SetPixel hdc, 12, tmph + 20, &H5D553B: SetPixel hdc, 13, tmph + 20, &H595338: SetPixel hdc, 14, tmph + 20, &H5E5739: SetPixel hdc, 15, tmph + 20, &H5F5A3C: SetPixel hdc, 16, tmph + 20, &H635E3F: SetPixel hdc, 17, tmph + 20, &H635D40:
    SetPixel hdc, 5, tmph + 21, &HFCFCFC: SetPixel hdc, 6, tmph + 21, &HF5F5F5: SetPixel hdc, 7, tmph + 21, &HEBEBEB: SetPixel hdc, 8, tmph + 21, &HE1E1E1: SetPixel hdc, 9, tmph + 21, &HD6D6D6: SetPixel hdc, 10, tmph + 21, &HCECECE: SetPixel hdc, 11, tmph + 21, &HC9C9C9: SetPixel hdc, 12, tmph + 21, &HC7C7C7: SetPixel hdc, 13, tmph + 21, &HC7C7C7: SetPixel hdc, 14, tmph + 21, &HC6C6C6: SetPixel hdc, 15, tmph + 21, &HC6C6C6: SetPixel hdc, 16, tmph + 21, &HC5C5C5: SetPixel hdc, 17, tmph + 21, &HC5C5C5:
    SetPixel hdc, 7, tmph + 22, &HFDFDFD: SetPixel hdc, 8, tmph + 22, &HF9F9F9: SetPixel hdc, 9, tmph + 22, &HF4F4F4: SetPixel hdc, 10, tmph + 22, &HF0F0F0: SetPixel hdc, 11, tmph + 22, &HEEEEEE: SetPixel hdc, 12, tmph + 22, &HEDEDED: SetPixel hdc, 13, tmph + 22, &HECECEC: SetPixel hdc, 14, tmph + 22, &HECECEC: SetPixel hdc, 15, tmph + 22, &HECECEC: SetPixel hdc, 16, tmph + 22, &HECECEC: SetPixel hdc, 17, tmph + 22, &HECECEC:
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, 0, &H450608: SetPixel hdc, tmpw + 18, 0, &H450608: SetPixel hdc, tmpw + 19, 0, &H3B0707: SetPixel hdc, tmpw + 20, 0, &H370706: SetPixel hdc, tmpw + 21, 0, &H360507: SetPixel hdc, tmpw + 22, 0, &H3B0F10: SetPixel hdc, tmpw + 23, 0, &H442526: SetPixel hdc, tmpw + 24, 0, &H604E4E: SetPixel hdc, tmpw + 25, 0, &HA29D9E: SetPixel hdc, tmpw + 26, 0, &HEEEEEE: SetPixel hdc, tmpw + 34, 0, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 1, &HF1D4C9: SetPixel hdc, tmpw + 18, 1, &HF1D4C9: SetPixel hdc, tmpw + 19, 1, &HEDD3CD: SetPixel hdc, tmpw + 20, 1, &HEBD1CB: SetPixel hdc, tmpw + 21, 1, &HE9CEC4: SetPixel hdc, tmpw + 22, 1, &HE5C1B9: SetPixel hdc, tmpw + 23, 1, &HCFA89F: SetPixel hdc, tmpw + 24, 1, &HAA6E68: SetPixel hdc, tmpw + 25, 1, &H73211B: SetPixel hdc, tmpw + 26, 1, &H702924: SetPixel hdc, tmpw + 27, 1, &HAA9897: SetPixel hdc, tmpw + 28, 1, &HF7F7F7: SetPixel hdc, tmpw + 34, 1, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 2, &HE5C8BD: SetPixel hdc, tmpw + 18, 2, &HE5C8BD: SetPixel hdc, tmpw + 19, 2, &HDEC4BE: SetPixel hdc, tmpw + 20, 2, &HDCC2BC: SetPixel hdc, tmpw + 21, 2, &HE2C7BD: SetPixel hdc, tmpw + 22, 2, &HE2BEB6: SetPixel hdc, tmpw + 23, 2, &HE8C1B8: SetPixel hdc, tmpw + 24, 2, &HF0B4AE: SetPixel hdc, tmpw + 25, 2, &HF29C96: SetPixel hdc, tmpw + 26, 2, &H822D27: SetPixel hdc, tmpw + 27, 2, &H400807: SetPixel hdc, tmpw + 28, 2, &H71585A: SetPixel hdc, tmpw + 29, 2, &HEBEBEB: SetPixel hdc, tmpw + 34, 2, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 3, &HE8C0A1: SetPixel hdc, tmpw + 18, 3, &HE8C0A1: SetPixel hdc, tmpw + 19, 3, &HE5C09A: SetPixel hdc, tmpw + 20, 3, &HE4BF99: SetPixel hdc, tmpw + 21, 3, &HE4BA97: SetPixel hdc, tmpw + 22, 3, &HE9BF9C: SetPixel hdc, tmpw + 23, 3, &HDFB695: SetPixel hdc, tmpw + 24, 3, &HDFB695: SetPixel hdc, tmpw + 25, 3, &HE0AE90: SetPixel hdc, tmpw + 26, 3, &HCB8469: SetPixel hdc, tmpw + 27, 3, &H941600: SetPixel hdc, tmpw + 28, 3, &H830800: SetPixel hdc, tmpw + 29, 3, &H895253: SetPixel hdc, tmpw + 30, 3, &HF0EFEF: SetPixel hdc, tmpw + 34, 3, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 4, &HE0B898: SetPixel hdc, tmpw + 18, 4, &HE0B897: SetPixel hdc, tmpw + 19, 4, &HDAB58F: SetPixel hdc, tmpw + 20, 4, &HDBB690: SetPixel hdc, tmpw + 21, 4, &HDBB18E: SetPixel hdc, tmpw + 22, 4, &HD7AD8A: SetPixel hdc, tmpw + 23, 4, &HDAB190: SetPixel hdc, tmpw + 24, 4, &HD2A988: SetPixel hdc, tmpw + 25, 4, &HD6A486: SetPixel hdc, tmpw + 26, 4, &HDA9378: SetPixel hdc, tmpw + 27, 4, &HBF4129: SetPixel hdc, tmpw + 28, 4, &H991B03: SetPixel hdc, tmpw + 29, 4, &H500709: SetPixel hdc, tmpw + 30, 4, &H826F70: SetPixel hdc, tmpw + 31, 4, &HFAFAFA: SetPixel hdc, tmpw + 34, 4, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 5, &HE3B48E: SetPixel hdc, tmpw + 18, 5, &HE3B48D: SetPixel hdc, tmpw + 19, 5, &HE0B387: SetPixel hdc, tmpw + 20, 5, &HDEB185: SetPixel hdc, tmpw + 21, 5, &HE1B084: SetPixel hdc, tmpw + 22, 5, &HE3AE83: SetPixel hdc, tmpw + 23, 5, &HE1AF7B: SetPixel hdc, tmpw + 24, 5, &HE0A976: SetPixel hdc, tmpw + 25, 5, &HDCA473: SetPixel hdc, tmpw + 26, 5, &HDEA372: SetPixel hdc, tmpw + 27, 5, &HCC712E: SetPixel hdc, tmpw + 28, 5, &HA53900: SetPixel hdc, tmpw + 29, 5, &H9D2200: SetPixel hdc, tmpw + 30, 5, &H9E2114: SetPixel hdc, tmpw + 31, 5, &HC7C5C4: SetPixel hdc, tmpw + 32, 5, &HFEFEFE: SetPixel hdc, tmpw + 34, 5, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 6, &HE0B18B: SetPixel hdc, tmpw + 18, 6, &HE0B18A: SetPixel hdc, tmpw + 19, 6, &HDEB185: SetPixel hdc, tmpw + 20, 6, &HDEB185: SetPixel hdc, tmpw + 21, 6, &HDCAB7F: SetPixel hdc, tmpw + 22, 6, &HE1AC81: SetPixel hdc, tmpw + 23, 6, &HDCAA76: SetPixel hdc, tmpw + 24, 6, &HDCA572: SetPixel hdc, tmpw + 25, 6, &HDBA372: SetPixel hdc, tmpw + 26, 6, &HD79C6B: SetPixel hdc, tmpw + 27, 6, &HD17633: SetPixel hdc, tmpw + 28, 6, &HB74B0B: SetPixel hdc, tmpw + 29, 6, &HAC310D: SetPixel hdc, tmpw + 30, 6, &H961507: SetPixel hdc, tmpw + 31, 6, &H736D6A: SetPixel hdc, tmpw + 32, 6, &HF8F8F8: SetPixel hdc, tmpw + 34, 6, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 7, &HE9B47F: SetPixel hdc, tmpw + 18, 7, &HEAB47E: SetPixel hdc, tmpw + 19, 7, &HEFB67E: SetPixel hdc, tmpw + 20, 7, &HE8AF77: SetPixel hdc, tmpw + 21, 7, &HE7AF74: SetPixel hdc, tmpw + 22, 7, &HE4AC71: SetPixel hdc, tmpw + 23, 7, &HEAAD6F: SetPixel hdc, tmpw + 24, 7, &HE9A968: SetPixel hdc, tmpw + 25, 7, &HE7A564: SetPixel hdc, tmpw + 26, 7, &HD9904C: SetPixel hdc, tmpw + 27, 7, &HC5711F: SetPixel hdc, tmpw + 28, 7, &HC16010: SetPixel hdc, tmpw + 29, 7, &HBB4D05: SetPixel hdc, tmpw + 30, 7, &HA02D00: SetPixel hdc, tmpw + 31, 7, &H774033: SetPixel hdc, tmpw + 32, 7, &HE7E6E6: SetPixel hdc, tmpw + 34, 7, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 8, &HCE9963: SetPixel hdc, tmpw + 18, 8, &HCF9963: SetPixel hdc, tmpw + 19, 8, &HCE955D: SetPixel hdc, tmpw + 20, 8, &HCE955D: SetPixel hdc, tmpw + 21, 8, &HCA9257: SetPixel hdc, tmpw + 22, 8, &HC89055: SetPixel hdc, tmpw + 23, 8, &HCB8E50: SetPixel hdc, tmpw + 24, 8, &HCB8B4A: SetPixel hdc, tmpw + 25, 8, &HC58342: SetPixel hdc, tmpw + 26, 8, &HC87F3B: SetPixel hdc, tmpw + 27, 8, &HCA7624: SetPixel hdc, tmpw + 28, 8, &HCA6919: SetPixel hdc, tmpw + 29, 8, &HCC5E16: SetPixel hdc, tmpw + 30, 8, &HB23E07: SetPixel hdc, tmpw + 31, 8, &H682B1D: SetPixel hdc, tmpw + 32, 8, &HC7C2C2: SetPixel hdc, tmpw + 33, 8, &HFDFDFD: SetPixel hdc, tmpw + 34, 8, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 9, &HDDA064: SetPixel hdc, tmpw + 18, 9, &HDCA064: SetPixel hdc, tmpw + 19, 9, &HDA9D5D: SetPixel hdc, tmpw + 20, 9, &HD99C5C: SetPixel hdc, tmpw + 21, 9, &HDA9D5D: SetPixel hdc, tmpw + 22, 9, &HDA9A5A: SetPixel hdc, tmpw + 23, 9, &HD89753: SetPixel hdc, tmpw + 24, 9, &HD7914E: SetPixel hdc, tmpw + 25, 9, &HD38E49: SetPixel hdc, tmpw + 26, 9, &HD38B43: SetPixel hdc, tmpw + 27, 9, &HCD8430: SetPixel hdc, tmpw + 28, 9, &HCA7826: SetPixel hdc, tmpw + 29, 9, &HCE6C1E: SetPixel hdc, tmpw + 30, 9, &HB9560C: SetPixel hdc, tmpw + 31, 9, &H742E0D: SetPixel hdc, tmpw + 32, 9, &HB3A6A4: SetPixel hdc, tmpw + 33, 9, &HF9F9F9: SetPixel hdc, tmpw + 34, 9, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 10, &HE2A66A: SetPixel hdc, tmpw + 18, 10, &HE2A66A: SetPixel hdc, tmpw + 19, 10, &HE1A464: SetPixel hdc, tmpw + 20, 10, &HE0A363: SetPixel hdc, tmpw + 21, 10, &HE0A363: SetPixel hdc, tmpw + 22, 10, &HE1A161: SetPixel hdc, tmpw + 23, 10, &HE09F5B: SetPixel hdc, tmpw + 24, 10, &HDE9855: SetPixel hdc, tmpw + 25, 10, &HDC9752: SetPixel hdc, tmpw + 26, 10, &HDB934B: SetPixel hdc, tmpw + 27, 10, &HD68D39: SetPixel hdc, tmpw + 28, 10, &HD17F2D: SetPixel hdc, tmpw + 29, 10, &HD67426: SetPixel hdc, tmpw + 30, 10, &HC05D13: SetPixel hdc, tmpw + 31, 10, &H7C3514: SetPixel hdc, tmpw + 32, 10, &HAB9B98: SetPixel hdc, tmpw + 33, 10, &HF6F6F6: SetPixel hdc, tmpw + 34, 10, &HFFFFFFFF:
    tmph = lh - 22
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, tmph + 10, &HE2A66A: SetPixel hdc, tmpw + 18, tmph + 10, &HE2A66A: SetPixel hdc, tmpw + 19, tmph + 10, &HE1A464: SetPixel hdc, tmpw + 20, tmph + 10, &HE0A363: SetPixel hdc, tmpw + 21, tmph + 10, &HE0A363: SetPixel hdc, tmpw + 22, tmph + 10, &HE1A161: SetPixel hdc, tmpw + 23, tmph + 10, &HE09F5B: SetPixel hdc, tmpw + 24, tmph + 10, &HDE9855: SetPixel hdc, tmpw + 25, tmph + 10, &HDC9752: SetPixel hdc, tmpw + 26, tmph + 10, &HDB934B: SetPixel hdc, tmpw + 27, tmph + 10, &HD68D39: SetPixel hdc, tmpw + 28, tmph + 10, &HD17F2D: SetPixel hdc, tmpw + 29, tmph + 10, &HD67426: SetPixel hdc, tmpw + 30, tmph + 10, &HC05D13: SetPixel hdc, tmpw + 31, tmph + 10, &H7C3514: SetPixel hdc, tmpw + 32, tmph + 10, &HAB9B98: SetPixel hdc, tmpw + 33, tmph + 10, &HF6F6F6:
    SetPixel hdc, tmpw + 17, tmph + 11, &HE6AC76: SetPixel hdc, tmpw + 18, tmph + 11, &HE6AC76: SetPixel hdc, tmpw + 19, tmph + 11, &HE2A86D: SetPixel hdc, tmpw + 20, tmph + 11, &HE5A66C: SetPixel hdc, tmpw + 21, tmph + 11, &HE1A56A: SetPixel hdc, tmpw + 22, tmph + 11, &HE4A46A: SetPixel hdc, tmpw + 23, tmph + 11, &HE1A266: SetPixel hdc, tmpw + 24, tmph + 11, &HE6A364: SetPixel hdc, tmpw + 25, tmph + 11, &HE19F5E: SetPixel hdc, tmpw + 26, tmph + 11, &HDF9A55: SetPixel hdc, tmpw + 27, tmph + 11, &HD89048: SetPixel hdc, tmpw + 28, tmph + 11, &HD88A3E: SetPixel hdc, tmpw + 29, tmph + 11, &HCF7927: SetPixel hdc, tmpw + 30, tmph + 11, &HC87220: SetPixel hdc, tmpw + 31, tmph + 11, &H77481E: SetPixel hdc, tmpw + 32, tmph + 11, &HABA39E: SetPixel hdc, tmpw + 33, tmph + 11, &HF4F4F4:
    SetPixel hdc, tmpw + 17, tmph + 12, &HF1B681: SetPixel hdc, tmpw + 18, tmph + 12, &HF0B780: SetPixel hdc, tmpw + 19, tmph + 12, &HF2B87D: SetPixel hdc, tmpw + 20, tmph + 12, &HF5B67C: SetPixel hdc, tmpw + 21, tmph + 12, &HF1B57A: SetPixel hdc, tmpw + 22, tmph + 12, &HF2B278: SetPixel hdc, tmpw + 23, tmph + 12, &HF0B175: SetPixel hdc, tmpw + 24, tmph + 12, &HF3B071: SetPixel hdc, tmpw + 25, tmph + 12, &HECAA69: SetPixel hdc, tmpw + 26, tmph + 12, &HE9A45F: SetPixel hdc, tmpw + 27, tmph + 12, &HE8A058: SetPixel hdc, tmpw + 28, tmph + 12, &HE5974B: SetPixel hdc, tmpw + 29, tmph + 12, &HE38D3B: SetPixel hdc, tmpw + 30, tmph + 12, &HD37D2B: SetPixel hdc, tmpw + 31, tmph + 12, &H895A32: SetPixel hdc, tmpw + 32, tmph + 12, &HB4AFAC: SetPixel hdc, tmpw + 33, tmph + 12, &HF5F5F5:
    SetPixel hdc, tmpw + 17, tmph + 13, &HF3BD8A: SetPixel hdc, tmpw + 18, tmph + 13, &HF3BD8A: SetPixel hdc, tmpw + 19, tmph + 13, &HF2BD84: SetPixel hdc, tmpw + 20, tmph + 13, &HF5BC84: SetPixel hdc, tmpw + 21, tmph + 13, &HF3BC83: SetPixel hdc, tmpw + 22, tmph + 13, &HF4B981: SetPixel hdc, tmpw + 23, tmph + 13, &HF2B97C: SetPixel hdc, tmpw + 24, tmph + 13, &HF5B77B: SetPixel hdc, tmpw + 25, tmph + 13, &HF1B476: SetPixel hdc, tmpw + 26, tmph + 13, &HEFAF6E: SetPixel hdc, tmpw + 27, tmph + 13, &HE5A45F: SetPixel hdc, tmpw + 28, tmph + 13, &HE49F5A: SetPixel hdc, tmpw + 29, tmph + 13, &HDA8F4A: SetPixel hdc, tmpw + 30, tmph + 13, &HC57A35: SetPixel hdc, tmpw + 31, tmph + 13, &H736353: SetPixel hdc, tmpw + 32, tmph + 13, &HD6D5D5: SetPixel hdc, tmpw + 33, tmph + 13, &HF8F8F8:
    SetPixel hdc, tmpw + 17, tmph + 14, &HFCC592: SetPixel hdc, tmpw + 18, tmph + 14, &HFBC592: SetPixel hdc, tmpw + 19, tmph + 14, &HF7C289: SetPixel hdc, tmpw + 20, tmph + 14, &HFCC38B: SetPixel hdc, tmpw + 21, tmph + 14, &HFAC38A: SetPixel hdc, tmpw + 22, tmph + 14, &HFDC28A: SetPixel hdc, tmpw + 23, tmph + 14, &HFBC285: SetPixel hdc, tmpw + 24, tmph + 14, &HFBBD81: SetPixel hdc, tmpw + 25, tmph + 14, &HF6B97B: SetPixel hdc, tmpw + 26, tmph + 14, &HF6B675: SetPixel hdc, tmpw + 27, tmph + 14, &HF0AF6A: SetPixel hdc, tmpw + 28, tmph + 14, &HE8A35E: SetPixel hdc, tmpw + 29, tmph + 14, &HDD924D: SetPixel hdc, tmpw + 30, tmph + 14, &HBA702B: SetPixel hdc, tmpw + 31, tmph + 14, &H847A70: SetPixel hdc, tmpw + 32, tmph + 14, &HE8E8E8: SetPixel hdc, tmpw + 33, tmph + 14, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 15, &HF8CE97: SetPixel hdc, tmpw + 18, tmph + 15, &HF9CD97: SetPixel hdc, tmpw + 19, tmph + 15, &HF9CE95: SetPixel hdc, tmpw + 20, tmph + 15, &HF7CC93: SetPixel hdc, tmpw + 21, tmph + 15, &HF6CB92: SetPixel hdc, tmpw + 22, tmph + 15, &HF9CA92: SetPixel hdc, tmpw + 23, tmph + 15, &HFCCD90: SetPixel hdc, tmpw + 24, tmph + 15, &HF8C488: SetPixel hdc, tmpw + 25, tmph + 15, &HF3BD80: SetPixel hdc, tmpw + 26, tmph + 15, &HFABD7D: SetPixel hdc, tmpw + 27, tmph + 15, &HF7B26D: SetPixel hdc, tmpw + 28, tmph + 15, &HEAA560: SetPixel hdc, tmpw + 29, tmph + 15, &HC0925D: SetPixel hdc, tmpw + 30, tmph + 15, &H896F54: SetPixel hdc, tmpw + 31, tmph + 15, &HBABABB: SetPixel hdc, tmpw + 32, tmph + 15, &HF1F1F1:
    SetPixel hdc, tmpw + 17, tmph + 16, &HFED59E: SetPixel hdc, tmpw + 18, tmph + 16, &HFFD59F: SetPixel hdc, tmpw + 19, tmph + 16, &HFED39A: SetPixel hdc, tmpw + 20, tmph + 16, &HFFD49B: SetPixel hdc, tmpw + 21, tmph + 16, &HFCD198: SetPixel hdc, tmpw + 22, tmph + 16, &HFFD098: SetPixel hdc, tmpw + 23, tmph + 16, &HFECF92: SetPixel hdc, tmpw + 24, tmph + 16, &HFFCB8F: SetPixel hdc, tmpw + 25, tmph + 16, &HFFC98C: SetPixel hdc, tmpw + 26, tmph + 16, &HFEC181: SetPixel hdc, tmpw + 27, tmph + 16, &HFBB671: SetPixel hdc, tmpw + 28, tmph + 16, &HF0AB66: SetPixel hdc, tmpw + 29, tmph + 16, &H9F733E: SetPixel hdc, tmpw + 30, tmph + 16, &H918478: SetPixel hdc, tmpw + 31, tmph + 16, &HE2E2E2: SetPixel hdc, tmpw + 32, tmph + 16, &HF9F9F9:
    SetPixel hdc, tmpw + 17, tmph + 17, &HF7DDA3: SetPixel hdc, tmpw + 18, tmph + 17, &HF8DDA4: SetPixel hdc, tmpw + 19, tmph + 17, &HF9E0A2: SetPixel hdc, tmpw + 20, tmph + 17, &HF5DC9E: SetPixel hdc, tmpw + 21, tmph + 17, &HF8DEA2: SetPixel hdc, tmpw + 22, tmph + 17, &HFBDDA2: SetPixel hdc, tmpw + 23, tmph + 17, &HF7D495: SetPixel hdc, tmpw + 24, tmph + 17, &HF8D193: SetPixel hdc, tmpw + 25, tmph + 17, &HFCCD90: SetPixel hdc, tmpw + 26, tmph + 17, &HF1C088: SetPixel hdc, tmpw + 27, tmph + 17, &HDBB186: SetPixel hdc, tmpw + 28, tmph + 17, &H8C7259: SetPixel hdc, tmpw + 29, tmph + 17, &H6D6B6B: SetPixel hdc, tmpw + 30, tmph + 17, &HD2D2D2: SetPixel hdc, tmpw + 31, tmph + 17, &HF2F2F2: SetPixel hdc, tmpw + 32, tmph + 17, &HFEFEFE:
    SetPixel hdc, tmpw + 17, tmph + 18, &HFFE6AD: SetPixel hdc, tmpw + 18, tmph + 18, &HFFE6AD: SetPixel hdc, tmpw + 19, tmph + 18, &HFFE7A9: SetPixel hdc, tmpw + 20, tmph + 18, &HFFEAAC: SetPixel hdc, tmpw + 21, tmph + 18, &HF7DDA1: SetPixel hdc, tmpw + 22, tmph + 18, &HFFE1A6: SetPixel hdc, tmpw + 23, tmph + 18, &HFFE1A2: SetPixel hdc, tmpw + 24, tmph + 18, &HFED799: SetPixel hdc, tmpw + 25, tmph + 18, &HFACC8F: SetPixel hdc, tmpw + 26, tmph + 18, &HC99A64: SetPixel hdc, tmpw + 27, tmph + 18, &H977048: SetPixel hdc, tmpw + 28, tmph + 18, &H817060: SetPixel hdc, tmpw + 29, tmph + 18, &HC8C8C8: SetPixel hdc, tmpw + 30, tmph + 18, &HEBEBEB: SetPixel hdc, tmpw + 31, tmph + 18, &HFCFCFC:
    SetPixel hdc, tmpw + 17, tmph + 19, &HE9E2C5: SetPixel hdc, tmpw + 18, tmph + 19, &HE9E2C5: SetPixel hdc, tmpw + 19, tmph + 19, &HEAE2C4: SetPixel hdc, tmpw + 20, tmph + 19, &HE7DFC1: SetPixel hdc, tmpw + 21, tmph + 19, &HEEE4C6: SetPixel hdc, tmpw + 22, tmph + 19, &HDBD1B4: SetPixel hdc, tmpw + 23, tmph + 19, &HB7AF93: SetPixel hdc, tmpw + 24, tmph + 19, &H8D8973: SetPixel hdc, tmpw + 25, tmph + 19, &H736D60: SetPixel hdc, tmpw + 26, tmph + 19, &H6A6660: SetPixel hdc, tmpw + 27, tmph + 19, &H8E9090: SetPixel hdc, tmpw + 28, tmph + 19, &HCDCDCD: SetPixel hdc, tmpw + 29, tmph + 19, &HE8E8E8: SetPixel hdc, tmpw + 30, tmph + 19, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 20, &H635D40: SetPixel hdc, tmpw + 18, tmph + 20, &H615B3F: SetPixel hdc, tmpw + 19, tmph + 20, &H60583C: SetPixel hdc, tmpw + 20, tmph + 20, &H5D563A: SetPixel hdc, tmpw + 21, tmph + 20, &H61583D: SetPixel hdc, tmpw + 22, tmph + 20, &H605840: SetPixel hdc, tmpw + 23, tmph + 20, &H6A6556: SetPixel hdc, tmpw + 24, tmph + 20, &H7F7D75: SetPixel hdc, tmpw + 25, tmph + 20, &HA4A3A1: SetPixel hdc, tmpw + 26, tmph + 20, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 20, &HDADADA: SetPixel hdc, tmpw + 28, tmph + 20, &HEDEDED: SetPixel hdc, tmpw + 29, tmph + 20, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 18, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 19, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 20, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 21, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 22, tmph + 21, &HC9C9C9: SetPixel hdc, tmpw + 23, tmph + 21, &HCECECE: SetPixel hdc, tmpw + 24, tmph + 21, &HD7D7D7: SetPixel hdc, tmpw + 25, tmph + 21, &HE1E1E1: SetPixel hdc, tmpw + 26, tmph + 21, &HECECEC: SetPixel hdc, tmpw + 27, tmph + 21, &HF6F6F6: SetPixel hdc, tmpw + 28, tmph + 21, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 18, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 19, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 20, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 21, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 22, tmph + 22, &HEDEDED: SetPixel hdc, tmpw + 23, tmph + 22, &HF0F0F0: SetPixel hdc, tmpw + 24, tmph + 22, &HF4F4F4: SetPixel hdc, tmpw + 25, tmph + 22, &HFAFAFA: SetPixel hdc, tmpw + 26, tmph + 22, &HFDFDFD:
    tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
    'Generar lineas intermedias
    DrawLineApi 0, tmph, 0, tmph1, &HF7F7F7: DrawLineApi 1, tmph, 1, tmph1, &HB0A09E: DrawLineApi 2, tmph, 2, tmph1, &H712E13: DrawLineApi 3, tmph, 3, tmph1, &HBD5F14:
    DrawLineApi 4, tmph, 4, tmph1, &HD17327: DrawLineApi 5, tmph, 5, tmph1, &HD47F31: DrawLineApi 6, tmph, 6, tmph1, &HD98C3D: DrawLineApi 7, tmph, 7, tmph1, &HD9944B:
    DrawLineApi 8, tmph, 8, tmph1, &HD7944F: DrawLineApi 9, tmph, 9, tmph1, &HDC9C55: DrawLineApi 10, tmph, 10, tmph1, &HDC9B57: DrawLineApi 11, tmph, 11, tmph1, &HE3A362:
    DrawLineApi 12, tmph, 12, tmph1, &HE3A265: DrawLineApi 13, tmph, 13, tmph1, &HE2A367: DrawLineApi 14, tmph, 14, tmph1, &HE0A165: DrawLineApi 15, tmph, 15, tmph1, &HE3A66A:
    DrawLineApi 16, tmph, 16, tmph1, &HE3A66A: DrawLineApi 17, tmph, 17, tmph1, &HE2A66A:
    DrawLineApi tmpw + 17, tmph, tmpw + 17, tmph1, &HE2A66A: DrawLineApi tmpw + 18, tmph, tmpw + 18, tmph1, &HE2A66A: DrawLineApi tmpw + 19, tmph, tmpw + 19, tmph1, &HE1A464:
    DrawLineApi tmpw + 20, tmph, tmpw + 20, tmph1, &HE0A363: DrawLineApi tmpw + 21, tmph, tmpw + 21, tmph1, &HE0A363: DrawLineApi tmpw + 22, tmph, tmpw + 22, tmph1, &HE1A161
    DrawLineApi tmpw + 23, tmph, tmpw + 23, tmph1, &HE09F5B: DrawLineApi tmpw + 24, tmph, tmpw + 24, tmph1, &HDE9855: DrawLineApi tmpw + 25, tmph, tmpw + 25, tmph1, &HDC9752:
    DrawLineApi tmpw + 26, tmph, tmpw + 26, tmph1, &HDB934B: DrawLineApi tmpw + 27, tmph, tmpw + 27, tmph1, &HD68D39: DrawLineApi tmpw + 28, tmph, tmpw + 28, tmph1, &HD17F2D:
    DrawLineApi tmpw + 29, tmph, tmpw + 29, tmph1, &HD67426: DrawLineApi tmpw + 30, tmph, tmpw + 30, tmph1, &HC05D13: DrawLineApi tmpw + 31, tmph, tmpw + 31, tmph1, &H7C3514:
    DrawLineApi tmpw + 32, tmph, tmpw + 32, tmph1, &HAB9B98: DrawLineApi tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6:
    'Lineas verticales
    DrawLineApi 17, 0, lw - 17, 0, &H450608
    DrawLineApi 17, 1, lw - 17, 1, &HF1D4C9
    DrawLineApi 17, 2, lw - 17, 2, &HE5C8BD
    DrawLineApi 17, 3, lw - 17, 3, &HE8C0A1
    DrawLineApi 17, 4, lw - 17, 4, &HE0B898
    DrawLineApi 17, 5, lw - 17, 5, &HE3B48E
    DrawLineApi 17, 6, lw - 17, 6, &HE0B18B
    DrawLineApi 17, 7, lw - 17, 7, &HE9B47F
    DrawLineApi 17, 8, lw - 17, 8, &HCE9963
    DrawLineApi 17, 9, lw - 17, 9, &HDDA064
    DrawLineApi 17, 10, lw - 17, 10, &HE2A66A
    DrawLineApi 17, 11, lw - 17, 11, &HE6AC76
    tmph = lh - 22
    DrawLineApi 17, tmph + 11, lw - 17, tmph + 11, &HE6AC76
    DrawLineApi 17, tmph + 12, lw - 17, tmph + 12, &HF1B681
    DrawLineApi 17, tmph + 13, lw - 17, tmph + 13, &HF3BD8A
    DrawLineApi 17, tmph + 14, lw - 17, tmph + 14, &HFCC592
    DrawLineApi 17, tmph + 15, lw - 17, tmph + 15, &HF8CE97
    DrawLineApi 17, tmph + 16, lw - 17, tmph + 16, &HFED59E
    DrawLineApi 17, tmph + 17, lw - 17, tmph + 17, &HF7DDA3
    DrawLineApi 17, tmph + 18, lw - 17, tmph + 18, &HFFE6AD
    DrawLineApi 17, tmph + 19, lw - 17, tmph + 19, &HE9E2C5
    DrawLineApi 17, tmph + 20, lw - 17, tmph + 20, &H635D40
    DrawLineApi 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    DrawLineApi 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub

Private Sub DrawAquaDown()

Dim tmph As Long, tmpw As Long
Dim tmph1 As Long, tmpw1 As Long
Dim lpRect As RECT

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    SetRect lpRect, 4, 4, lw - 4, lh - 4
    PaintRect &HCC9B6A, lpRect

    SetPixel hdc, 6, 0, &HFEFEFE: SetPixel hdc, 7, 0, &HE5E4E4: SetPixel hdc, 8, 0, &HA5A2A2: SetPixel hdc, 9, 0, &H675C5C: SetPixel hdc, 10, 0, &H422729: SetPixel hdc, 11, 0, &H300E0D: SetPixel hdc, 12, 0, &H300A09: SetPixel hdc, 13, 0, &H2F0908: SetPixel hdc, 14, 0, &H330909: SetPixel hdc, 15, 0, &H390A0A: SetPixel hdc, 16, 0, &H3C0A0A: SetPixel hdc, 17, 0, &H3C090A:
    SetPixel hdc, 5, 1, &HF0EEEE: SetPixel hdc, 6, 1, &H9D888A: SetPixel hdc, 7, 1, &H653531: SetPixel hdc, 8, 1, &H5A201D: SetPixel hdc, 9, 1, &H8D655F: SetPixel hdc, 10, 1, &HB99995: SetPixel hdc, 11, 1, &HD0B4B2: SetPixel hdc, 12, 1, &HD7BEBB: SetPixel hdc, 13, 1, &HDDC6C0: SetPixel hdc, 14, 1, &HDDC6C0: SetPixel hdc, 15, 1, &HDDC7BE: SetPixel hdc, 16, 1, &HDDC7BE: SetPixel hdc, 17, 1, &HDEC7BE:
    SetPixel hdc, 3, 2, &HFEFEFE: SetPixel hdc, 4, 2, &HE4E4E4: SetPixel hdc, 5, 2, &H6F5C5C: SetPixel hdc, 6, 2, &H390A0E: SetPixel hdc, 7, 2, &H712E2A: SetPixel hdc, 8, 2, &HD6928D: SetPixel hdc, 9, 2, &HD8ACA6: SetPixel hdc, 10, 2, &HD1B0AC: SetPixel hdc, 11, 2, &HD1B5B2: SetPixel hdc, 12, 2, &HD0B7B4: SetPixel hdc, 13, 2, &HCEB7B1: SetPixel hdc, 14, 2, &HCEB7B1: SetPixel hdc, 15, 2, &HD2BCB2: SetPixel hdc, 16, 2, &HD2BCB2: SetPixel hdc, 17, 2, &HD3BCB2:
    SetPixel hdc, 3, 3, &HEEEDED: SetPixel hdc, 4, 3, &H805858: SetPixel hdc, 5, 3, &H6A0D08: SetPixel hdc, 6, 3, &H7D1909: SetPixel hdc, 7, 3, &HB07B63: SetPixel hdc, 8, 3, &HCFA58A: SetPixel hdc, 9, 3, &HCDA78E: SetPixel hdc, 10, 3, &HD1AB92: SetPixel hdc, 11, 3, &HD2AF93: SetPixel hdc, 12, 3, &HD3B094: SetPixel hdc, 13, 3, &HD0AF93: SetPixel hdc, 14, 3, &HD3B296: SetPixel hdc, 15, 3, &HD4B49A: SetPixel hdc, 16, 3, &HD4B39A: SetPixel hdc, 17, 3, &HD4B39A:
    SetPixel hdc, 2, 4, &HFBFBFB: SetPixel hdc, 3, 4, &H837576: SetPixel hdc, 4, 4, &H440C0C: SetPixel hdc, 5, 4, &H821D0D: SetPixel hdc, 6, 4, &HA94433: SetPixel hdc, 7, 4, &HC08B72: SetPixel hdc, 8, 4, &HC49A7F: SetPixel hdc, 9, 4, &HC6A188: SetPixel hdc, 10, 4, &HC7A189: SetPixel hdc, 11, 4, &HC6A387: SetPixel hdc, 12, 4, &HC8A689: SetPixel hdc, 13, 4, &HC9A98C: SetPixel hdc, 14, 4, &HC8A88B: SetPixel hdc, 15, 4, &HCBAB91: SetPixel hdc, 16, 4, &HCCAC92: SetPixel hdc, 17, 4, &HCCAC92:
    SetPixel hdc, 1, 5, &HFEFEFE: SetPixel hdc, 2, 5, &HCAC8C7: SetPixel hdc, 3, 5, &H79281D: SetPixel hdc, 4, 5, &H7F2409: SetPixel hdc, 5, 5, &H8C3809: SetPixel hdc, 6, 5, &HBD6D39: SetPixel hdc, 7, 5, &HC9986E: SetPixel hdc, 8, 5, &HC89D74: SetPixel hdc, 9, 5, &HC49A71: SetPixel hdc, 10, 5, &HCAA17C: SetPixel hdc, 11, 5, &HC6A07A: SetPixel hdc, 12, 5, &HCAA480: SetPixel hdc, 13, 5, &HCAA582: SetPixel hdc, 14, 5, &HCBA584: SetPixel hdc, 15, 5, &HCDA989: SetPixel hdc, 16, 5, &HCFA98A: SetPixel hdc, 17, 5, &HCFA88A:
    SetPixel hdc, 1, 6, &HF9F9F9: SetPixel hdc, 2, 6, &H756C6B: SetPixel hdc, 3, 6, &H76190D: SetPixel hdc, 4, 6, &H913416: SetPixel hdc, 5, 6, &H9D4916: SetPixel hdc, 6, 6, &HBA6A36: SetPixel hdc, 7, 6, &HC39268: SetPixel hdc, 8, 6, &HC59A71: SetPixel hdc, 9, 6, &HC59B72: SetPixel hdc, 10, 6, &HC59C77: SetPixel hdc, 11, 6, &HC6A07A: SetPixel hdc, 12, 6, &HC6A07C: SetPixel hdc, 13, 6, &HC7A27F: SetPixel hdc, 14, 6, &HCBA584: SetPixel hdc, 15, 6, &HCAA686: SetPixel hdc, 16, 6, &HCBA586: SetPixel hdc, 17, 6, &HCCA587:
    SetPixel hdc, 1, 7, &HE8E7E7: SetPixel hdc, 2, 7, &H6C3E35: SetPixel hdc, 3, 7, &H8A2D09: SetPixel hdc, 4, 7, &HA34812: SetPixel hdc, 5, 7, &HAB591A: SetPixel hdc, 6, 7, &HB46B2B: SetPixel hdc, 7, 7, &HC3854A: SetPixel hdc, 8, 7, &HD19C64: SetPixel hdc, 9, 7, &HCD9C6C: SetPixel hdc, 10, 7, &HD1A070: SetPixel hdc, 11, 7, &HD2A272: SetPixel hdc, 12, 7, &HD2A272: SetPixel hdc, 13, 7, &HD6A57A: SetPixel hdc, 14, 7, &HD8A77C: SetPixel hdc, 15, 7, &HD2A87C: SetPixel hdc, 16, 7, &HD2A87C: SetPixel hdc, 17, 7, &HD2A77D:
    SetPixel hdc, 0, 8, &HFDFDFD: SetPixel hdc, 1, 8, &HC7C3C3: SetPixel hdc, 2, 8, &H5C2A21: SetPixel hdc, 3, 8, &H9C3E15: SetPixel hdc, 4, 8, &HB35A22: SetPixel hdc, 5, 8, &HB56324: SetPixel hdc, 6, 8, &HB66D2D: SetPixel hdc, 7, 8, &HB6783D: SetPixel hdc, 8, 8, &HB07B44: SetPixel hdc, 9, 8, &HB18050: SetPixel hdc, 10, 8, &HB58454: SetPixel hdc, 11, 8, &HB48554: SetPixel hdc, 12, 8, &HB78858: SetPixel hdc, 13, 8, &HBA895E: SetPixel hdc, 14, 8, &HBB8A5F: SetPixel hdc, 15, 8, &HB98E62: SetPixel hdc, 16, 8, &HB98E62: SetPixel hdc, 17, 8, &HB98E62:
    SetPixel hdc, 0, 9, &HFAFAFA: SetPixel hdc, 1, 9, &HB4ABA9: SetPixel hdc, 2, 9, &H612A14: SetPixel hdc, 3, 9, &HA05316: SetPixel hdc, 4, 9, &HB36628: SetPixel hdc, 5, 9, &HB67132: SetPixel hdc, 6, 9, &HB67738: SetPixel hdc, 7, 9, &HB98146: SetPixel hdc, 8, 9, &HBD864E: SetPixel hdc, 9, 9, &HBD894F: SetPixel hdc, 10, 9, &HC28D55: SetPixel hdc, 11, 9, &HC4905B: SetPixel hdc, 12, 9, &HC5905F: SetPixel hdc, 13, 9, &HC49161: SetPixel hdc, 14, 9, &HC49161: SetPixel hdc, 15, 9, &HC69564: SetPixel hdc, 16, 9, &HC69564: SetPixel hdc, 17, 9, &HC69464:
    SetPixel hdc, 0, 10, &HF7F7F7: SetPixel hdc, 1, 10, &HA99D9B: SetPixel hdc, 2, 10, &H632D17: SetPixel hdc, 3, 10, &HA65A1D: SetPixel hdc, 4, 10, &HB96C2E: SetPixel hdc, 5, 10, &HBC7738: SetPixel hdc, 6, 10, &HC18242: SetPixel hdc, 7, 10, &HC2894E: SetPixel hdc, 8, 10, &HC18A52: SetPixel hdc, 9, 10, &HC59157: SetPixel hdc, 10, 10, &HC59159: SetPixel hdc, 11, 10, &HCC9863: SetPixel hdc, 12, 10, &HCC9665: SetPixel hdc, 13, 10, &HCB9767: SetPixel hdc, 14, 10, &HC99565: SetPixel hdc, 15, 10, &HCC9A6A: SetPixel hdc, 16, 10, &HCC9A6A: SetPixel hdc, 17, 10, &HCC9B6A:
    tmph = lh - 22
    SetPixel hdc, 0, tmph + 10, &HF7F7F7: SetPixel hdc, 1, tmph + 10, &HA99D9B: SetPixel hdc, 2, tmph + 10, &H632D17: SetPixel hdc, 3, tmph + 10, &HA65A1D: SetPixel hdc, 4, tmph + 10, &HB96C2E: SetPixel hdc, 5, tmph + 10, &HBC7738: SetPixel hdc, 6, tmph + 10, &HC18242: SetPixel hdc, 7, tmph + 10, &HC2894E: SetPixel hdc, 8, tmph + 10, &HC18A52: SetPixel hdc, 9, tmph + 10, &HC59157: SetPixel hdc, 10, tmph + 10, &HC59159: SetPixel hdc, 11, tmph + 10, &HCC9863: SetPixel hdc, 12, tmph + 10, &HCC9665: SetPixel hdc, 13, tmph + 10, &HCB9767: SetPixel hdc, 14, tmph + 10, &HC99565: SetPixel hdc, 15, tmph + 10, &HCC9A6A: SetPixel hdc, 16, tmph + 10, &HCC9A6A: SetPixel hdc, 17, tmph + 10, &HCC9B6A:
    SetPixel hdc, 0, tmph + 11, &HF5F5F5: SetPixel hdc, 1, tmph + 11, &HA59F9A: SetPixel hdc, 2, tmph + 11, &H674024: SetPixel hdc, 3, tmph + 11, &HAE6827: SetPixel hdc, 4, tmph + 11, &HB97231: SetPixel hdc, 5, tmph + 11, &HBE8247: SetPixel hdc, 6, tmph + 11, &HC0874E: SetPixel hdc, 7, tmph + 11, &HC78E56: SetPixel hdc, 8, tmph + 11, &HCD9561: SetPixel hdc, 9, tmph + 11, &HCB9466: SetPixel hdc, 10, tmph + 11, &HCD9A6B: SetPixel hdc, 11, tmph + 11, &HC79867: SetPixel hdc, 12, tmph + 11, &HCA9B6A: SetPixel hdc, 13, tmph + 11, &HCC9D6C: SetPixel hdc, 14, tmph + 11, &HCD9D70: SetPixel hdc, 15, tmph + 11, &HD0A175: SetPixel hdc, 16, tmph + 11, &HD0A175: SetPixel hdc, 17, tmph + 11, &HD0A175:
    SetPixel hdc, 0, tmph + 12, &HF5F5F5: SetPixel hdc, 1, tmph + 12, &HACA7A4: SetPixel hdc, 2, tmph + 12, &H755035: SetPixel hdc, 3, tmph + 12, &HB77131: SetPixel hdc, 4, tmph + 12, &HCB8443: SetPixel hdc, 5, tmph + 12, &HC5894E: SetPixel hdc, 6, tmph + 12, &HCC935A: SetPixel hdc, 7, tmph + 12, &HD29962: SetPixel hdc, 8, tmph + 12, &HD69F6A: SetPixel hdc, 9, tmph + 12, &HDBA476: SetPixel hdc, 10, tmph + 12, &HD6A374: SetPixel hdc, 11, tmph + 12, &HD4A574: SetPixel hdc, 12, tmph + 12, &HD8A978: SetPixel hdc, 13, tmph + 12, &HDAAB7A: SetPixel hdc, 14, tmph + 12, &HDAAA7D: SetPixel hdc, 15, tmph + 12, &HDBAB7F: SetPixel hdc, 16, tmph + 12, &HDAAA7F: SetPixel hdc, 17, tmph + 12, &HDAAA7F:
    SetPixel hdc, 0, tmph + 13, &HF7F7F7: SetPixel hdc, 1, tmph + 13, &HC0C0BF: SetPixel hdc, 2, tmph + 13, &H63574B: SetPixel hdc, 3, tmph + 13, &HAC7036: SetPixel hdc, 4, tmph + 13, &HC2854A: SetPixel hdc, 5, tmph + 13, &HCF955E: SetPixel hdc, 6, tmph + 13, &HD29B66: SetPixel hdc, 7, tmph + 13, &HD1A26E: SetPixel hdc, 8, tmph + 13, &HD8A776: SetPixel hdc, 9, tmph + 13, &HDBA878: SetPixel hdc, 10, tmph + 13, &HDFAC7C: SetPixel hdc, 11, tmph + 13, &HDBAF7D: SetPixel hdc, 12, tmph + 13, &HDDAF81: SetPixel hdc, 13, tmph + 13, &HDEB183: SetPixel hdc, 14, tmph + 13, &HDDAF84: SetPixel hdc, 15, tmph + 13, &HDEB087: SetPixel hdc, 16, tmph + 13, &HDEB087: SetPixel hdc, 17, tmph + 13, &HDCB087:
    SetPixel hdc, 0, tmph + 14, &HFBFBFB: SetPixel hdc, 1, tmph + 14, &HE1E1E1: SetPixel hdc, 2, tmph + 14, &H7C7269: SetPixel hdc, 3, tmph + 14, &HA26830: SetPixel hdc, 4, tmph + 14, &HC6884E: SetPixel hdc, 5, tmph + 14, &HD0965F: SetPixel hdc, 6, tmph + 14, &HDAA26E: SetPixel hdc, 7, tmph + 14, &HD9AA75: SetPixel hdc, 8, tmph + 14, &HDBAA79: SetPixel hdc, 9, tmph + 14, &HE2AF7F: SetPixel hdc, 10, tmph + 14, &HE6B484: SetPixel hdc, 11, tmph + 14, &HE2B684: SetPixel hdc, 12, tmph + 14, &HE3B588: SetPixel hdc, 13, tmph + 14, &HE2B587: SetPixel hdc, 14, tmph + 14, &HE2B48A: SetPixel hdc, 15, tmph + 14, &HE5B78E: SetPixel hdc, 16, tmph + 14, &HE5B78E: SetPixel hdc, 17, tmph + 14, &HE4B88E:
    SetPixel hdc, 0, tmph + 15, &HFEFEFE: SetPixel hdc, 1, tmph + 15, &HEDEDED: SetPixel hdc, 2, tmph + 15, &H9E9C9C: SetPixel hdc, 3, tmph + 15, &H766051: SetPixel hdc, 4, tmph + 15, &HAD8666: SetPixel hdc, 5, tmph + 15, &HD49A61: SetPixel hdc, 6, tmph + 15, &HE0A66D: SetPixel hdc, 7, tmph + 15, &HE3B17C: SetPixel hdc, 8, tmph + 15, &HE0B380: SetPixel hdc, 9, tmph + 15, &HE0B587: SetPixel hdc, 10, tmph + 15, &HE2BC8C: SetPixel hdc, 11, tmph + 15, &HE0BB8B: SetPixel hdc, 12, tmph + 15, &HE0BC8B: SetPixel hdc, 13, tmph + 15, &HE3BD92: SetPixel hdc, 14, tmph + 15, &HE2BC91: SetPixel hdc, 15, tmph + 15, &HE2BF93: SetPixel hdc, 16, tmph + 15, &HE1BE93: SetPixel hdc, 17, tmph + 15, &HE1BF93:
    SetPixel hdc, 1, tmph + 16, &HF6F6F6: SetPixel hdc, 2, tmph + 16, &HD5D5D5: SetPixel hdc, 3, tmph + 16, &H86766C: SetPixel hdc, 4, tmph + 16, &H856144: SetPixel hdc, 5, tmph + 16, &HD59C63: SetPixel hdc, 6, tmph + 16, &HE5AB71: SetPixel hdc, 7, tmph + 16, &HE5B37E: SetPixel hdc, 8, tmph + 16, &HE7BB88: SetPixel hdc, 9, tmph + 16, &HE7BF91: SetPixel hdc, 10, tmph + 16, &HE3BC8D: SetPixel hdc, 11, tmph + 16, &HE7C392: SetPixel hdc, 12, tmph + 16, &HE7C392: SetPixel hdc, 13, tmph + 16, &HE8C398: SetPixel hdc, 14, tmph + 16, &HE8C499: SetPixel hdc, 15, tmph + 16, &HE8C599: SetPixel hdc, 16, tmph + 16, &HE8C599: SetPixel hdc, 17, tmph + 16, &HE7C699:
    SetPixel hdc, 1, tmph + 17, &HFDFDFD: SetPixel hdc, 2, tmph + 17, &HEDEDED: SetPixel hdc, 3, tmph + 17, &HBDBDBD: SetPixel hdc, 4, tmph + 17, &H676767: SetPixel hdc, 5, tmph + 17, &H71604C: SetPixel hdc, 6, tmph + 17, &HBEA17D: SetPixel hdc, 7, tmph + 17, &HDAB381: SetPixel hdc, 8, tmph + 17, &HE5BE8C: SetPixel hdc, 9, tmph + 17, &HE1C18F: SetPixel hdc, 10, tmph + 17, &HE4C895: SetPixel hdc, 11, tmph + 17, &HDFCA98: SetPixel hdc, 12, tmph + 17, &HE2CE9B: SetPixel hdc, 13, tmph + 17, &HE2CE9B: SetPixel hdc, 14, tmph + 17, &HE2CE9B: SetPixel hdc, 15, tmph + 17, &HE2CD9D: SetPixel hdc, 16, tmph + 17, &HE2CC9D: SetPixel hdc, 17, tmph + 17, &HE2CC9D:
    SetPixel hdc, 2, tmph + 18, &HF9F9F9: SetPixel hdc, 3, tmph + 18, &HE6E6E6: SetPixel hdc, 4, tmph + 18, &HB9B9B9: SetPixel hdc, 5, tmph + 18, &H7A7163: SetPixel hdc, 6, tmph + 18, &H776043: SetPixel hdc, 7, tmph + 18, &HAB885B: SetPixel hdc, 8, tmph + 18, &HDDB888: SetPixel hdc, 9, tmph + 18, &HE6C796: SetPixel hdc, 10, tmph + 18, &HE8CD9A: SetPixel hdc, 11, tmph + 18, &HE5D19E: SetPixel hdc, 12, tmph + 18, &HE9D6A3: SetPixel hdc, 13, tmph + 18, &HE9D6A5: SetPixel hdc, 14, tmph + 18, &HE9D6A3: SetPixel hdc, 15, tmph + 18, &HE9D5A6: SetPixel hdc, 16, tmph + 18, &HE9D5A6: SetPixel hdc, 17, tmph + 18, &HE9D5A6:
    SetPixel hdc, 2, tmph + 19, &HFEFEFE: SetPixel hdc, 3, tmph + 19, &HF8F8F8: SetPixel hdc, 4, tmph + 19, &HE6E6E6: SetPixel hdc, 5, tmph + 19, &HC8C8C8: SetPixel hdc, 6, tmph + 19, &H8C8C8C: SetPixel hdc, 7, tmph + 19, &H61605E: SetPixel hdc, 8, tmph + 19, &H656059: SetPixel hdc, 9, tmph + 19, &H857C6D: SetPixel hdc, 10, tmph + 19, &HA59C87: SetPixel hdc, 11, tmph + 19, &HC8C1A8: SetPixel hdc, 12, tmph + 19, &HD1CAB0: SetPixel hdc, 13, tmph + 19, &HD5CFB5: SetPixel hdc, 14, tmph + 19, &HD6D1B6: SetPixel hdc, 15, tmph + 19, &HD7D2BA: SetPixel hdc, 16, tmph + 19, &HD7D1BA: SetPixel hdc, 17, tmph + 19, &HD7D2BA:
    SetPixel hdc, 3, tmph + 20, &HFEFEFE: SetPixel hdc, 4, tmph + 20, &HF9F9F9: SetPixel hdc, 5, tmph + 20, &HECECEC: SetPixel hdc, 6, tmph + 20, &HDADADA: SetPixel hdc, 7, tmph + 20, &HC1C1C1: SetPixel hdc, 8, tmph + 20, &H9C9B99: SetPixel hdc, 9, tmph + 20, &H7D7A73: SetPixel hdc, 10, tmph + 20, &H635E50: SetPixel hdc, 11, tmph + 20, &H58533F: SetPixel hdc, 12, tmph + 20, &H554F39: SetPixel hdc, 13, tmph + 20, &H514D36: SetPixel hdc, 14, tmph + 20, &H554F37: SetPixel hdc, 15, tmph + 20, &H57523A: SetPixel hdc, 16, tmph + 20, &H5A563D: SetPixel hdc, 17, tmph + 20, &H5A563E:
    SetPixel hdc, 5, tmph + 21, &HFCFCFC: SetPixel hdc, 6, tmph + 21, &HF5F5F5: SetPixel hdc, 7, tmph + 21, &HEBEBEB: SetPixel hdc, 8, tmph + 21, &HE1E1E1: SetPixel hdc, 9, tmph + 21, &HD6D6D6: SetPixel hdc, 10, tmph + 21, &HCECECE: SetPixel hdc, 11, tmph + 21, &HC9C9C9: SetPixel hdc, 12, tmph + 21, &HC7C7C7: SetPixel hdc, 13, tmph + 21, &HC7C7C7: SetPixel hdc, 14, tmph + 21, &HC6C6C6: SetPixel hdc, 15, tmph + 21, &HC6C6C6: SetPixel hdc, 16, tmph + 21, &HC5C5C5: SetPixel hdc, 17, tmph + 21, &HC5C5C5:
    SetPixel hdc, 7, tmph + 22, &HFDFDFD: SetPixel hdc, 8, tmph + 22, &HF9F9F9: SetPixel hdc, 9, tmph + 22, &HF4F4F4: SetPixel hdc, 10, tmph + 22, &HF0F0F0: SetPixel hdc, 11, tmph + 22, &HEEEEEE: SetPixel hdc, 12, tmph + 22, &HEDEDED: SetPixel hdc, 13, tmph + 22, &HECECEC: SetPixel hdc, 14, tmph + 22, &HECECEC: SetPixel hdc, 15, tmph + 22, &HECECEC: SetPixel hdc, 16, tmph + 22, &HECECEC: SetPixel hdc, 17, tmph + 22, &HECECEC:
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, 0, &H3C090A: SetPixel hdc, tmpw + 18, 0, &H3C090A: SetPixel hdc, tmpw + 19, 0, &H340A0A: SetPixel hdc, tmpw + 20, 0, &H300A09: SetPixel hdc, tmpw + 21, 0, &H2F080A: SetPixel hdc, tmpw + 22, 0, &H341011: SetPixel hdc, tmpw + 23, 0, &H3E2526: SetPixel hdc, tmpw + 24, 0, &H5A4C4C: SetPixel hdc, tmpw + 25, 0, &H9E9B9B: SetPixel hdc, tmpw + 26, 0, &HEEEEEE: SetPixel hdc, tmpw + 34, 0, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 1, &HDEC7BE: SetPixel hdc, tmpw + 18, 1, &HDEC7BE: SetPixel hdc, tmpw + 19, 1, &HDBC6C1: SetPixel hdc, tmpw + 20, 1, &HD9C4BF: SetPixel hdc, tmpw + 21, 1, &HD7C1B9: SetPixel hdc, tmpw + 22, 1, &HD3B5AF: SetPixel hdc, tmpw + 23, 1, &HBE9F97: SetPixel hdc, tmpw + 24, 1, &H9B6A65: SetPixel hdc, tmpw + 25, 1, &H65231E: SetPixel hdc, tmpw + 26, 1, &H642A26: SetPixel hdc, tmpw + 27, 1, &HA59696: SetPixel hdc, tmpw + 28, 1, &HF7F7F7: SetPixel hdc, tmpw + 34, 1, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 2, &HD3BCB2: SetPixel hdc, tmpw + 18, 2, &HD3BCB2: SetPixel hdc, tmpw + 19, 2, &HCDB8B3: SetPixel hdc, tmpw + 20, 2, &HCBB6B1: SetPixel hdc, tmpw + 21, 2, &HD0BBB2: SetPixel hdc, tmpw + 22, 2, &HD0B2AC: SetPixel hdc, tmpw + 23, 2, &HD6B6AF: SetPixel hdc, tmpw + 24, 2, &HDCABA6: SetPixel hdc, tmpw + 25, 2, &HDC9691: SetPixel hdc, tmpw + 26, 2, &H732E29: SetPixel hdc, tmpw + 27, 2, &H380A0A: SetPixel hdc, tmpw + 28, 2, &H6A5556: SetPixel hdc, tmpw + 29, 2, &HEAEBEA: SetPixel hdc, tmpw + 34, 2, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 3, &HD4B39A: SetPixel hdc, tmpw + 18, 3, &HD4B39A: SetPixel hdc, tmpw + 19, 3, &HD1B294: SetPixel hdc, tmpw + 20, 3, &HD0B193: SetPixel hdc, tmpw + 21, 3, &HD0AE91: SetPixel hdc, tmpw + 22, 3, &HD4B296: SetPixel hdc, tmpw + 23, 3, &HCBAA8F: SetPixel hdc, tmpw + 24, 3, &HCBAA8F: SetPixel hdc, tmpw + 25, 3, &HCCA38B: SetPixel hdc, tmpw + 26, 3, &HB77E68: SetPixel hdc, tmpw + 27, 3, &H811B09: SetPixel hdc, tmpw + 28, 3, &H720E08: SetPixel hdc, tmpw + 29, 3, &H7D5051: SetPixel hdc, tmpw + 30, 3, &HEFEEEE: SetPixel hdc, tmpw + 34, 3, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 4, &HCCAC92: SetPixel hdc, tmpw + 18, 4, &HCCAC91: SetPixel hdc, tmpw + 19, 4, &HC6A889: SetPixel hdc, tmpw + 20, 4, &HC7A98A: SetPixel hdc, tmpw + 21, 4, &HC7A589: SetPixel hdc, tmpw + 22, 4, &HC4A185: SetPixel hdc, tmpw + 23, 4, &HC6A58A: SetPixel hdc, tmpw + 24, 4, &HBF9E83: SetPixel hdc, tmpw + 25, 4, &HC39A82: SetPixel hdc, tmpw + 26, 4, &HC58C76: SetPixel hdc, tmpw + 27, 4, &HA9432F: SetPixel hdc, tmpw + 28, 4, &H861F0C: SetPixel hdc, tmpw + 29, 4, &H460B0C: SetPixel hdc, tmpw + 30, 4, &H7B6B6C: SetPixel hdc, tmpw + 31, 4, &HFAFAFA: SetPixel hdc, tmpw + 34, 4, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 5, &HCFA88A: SetPixel hdc, tmpw + 18, 5, &HCFA889: SetPixel hdc, tmpw + 19, 5, &HCBA683: SetPixel hdc, tmpw + 20, 5, &HC9A481: SetPixel hdc, tmpw + 21, 5, &HCCA480: SetPixel hdc, tmpw + 22, 5, &HCEA280: SetPixel hdc, tmpw + 23, 5, &HCCA379: SetPixel hdc, tmpw + 24, 5, &HCA9E74: SetPixel hdc, tmpw + 25, 5, &HC69971: SetPixel hdc, tmpw + 26, 5, &HC89870: SetPixel hdc, tmpw + 27, 5, &HB46A34: SetPixel hdc, tmpw + 28, 5, &H90380A: SetPixel hdc, tmpw + 29, 5, &H892509: SetPixel hdc, tmpw + 30, 5, &H8A251B: SetPixel hdc, tmpw + 31, 5, &HC4C2C2: SetPixel hdc, tmpw + 32, 5, &HFEFEFE: SetPixel hdc, tmpw + 34, 5, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 6, &HCCA587: SetPixel hdc, tmpw + 18, 6, &HCCA586: SetPixel hdc, tmpw + 19, 6, &HC9A481: SetPixel hdc, tmpw + 20, 6, &HC9A481: SetPixel hdc, tmpw + 21, 6, &HC7A07C: SetPixel hdc, tmpw + 22, 6, &HCCA17E: SetPixel hdc, tmpw + 23, 6, &HC79F74: SetPixel hdc, tmpw + 24, 6, &HC69A70: SetPixel hdc, tmpw + 25, 6, &HC59870: SetPixel hdc, tmpw + 26, 6, &HC2926A: SetPixel hdc, tmpw + 27, 6, &HB96F39: SetPixel hdc, tmpw + 28, 6, &HA04814: SetPixel hdc, tmpw + 29, 6, &H973215: SetPixel hdc, tmpw + 30, 6, &H831A0F: SetPixel hdc, tmpw + 31, 6, &H6E6966: SetPixel hdc, tmpw + 32, 6, &HF8F8F8: SetPixel hdc, tmpw + 34, 6, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 7, &HD2A77D: SetPixel hdc, tmpw + 18, 7, &HD3A77C: SetPixel hdc, tmpw + 19, 7, &HD8AA7D: SetPixel hdc, tmpw + 20, 7, &HD2A376: SetPixel hdc, tmpw + 21, 7, &HD1A373: SetPixel hdc, tmpw + 22, 7, &HCEA070: SetPixel hdc, tmpw + 23, 7, &HD2A06F: SetPixel hdc, tmpw + 24, 7, &HD19D68: SetPixel hdc, tmpw + 25, 7, &HD09A65: SetPixel hdc, tmpw + 26, 7, &HC2864F: SetPixel hdc, tmpw + 27, 7, &HAE6927: SetPixel hdc, tmpw + 28, 7, &HA95A19: SetPixel hdc, tmpw + 29, 7, &HA44A10: SetPixel hdc, tmpw + 30, 7, &H8B2E09: SetPixel hdc, tmpw + 31, 7, &H6B3E34: SetPixel hdc, tmpw + 32, 7, &HE7E6E6: SetPixel hdc, tmpw + 34, 7, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 8, &HB98E62: SetPixel hdc, tmpw + 18, 8, &HBA8E62: SetPixel hdc, tmpw + 19, 8, &HB98B5E: SetPixel hdc, tmpw + 20, 8, &HB98B5E: SetPixel hdc, tmpw + 21, 8, &HB68858: SetPixel hdc, tmpw + 22, 8, &HB48656: SetPixel hdc, tmpw + 23, 8, &HB58452: SetPixel hdc, tmpw + 24, 8, &HB5814C: SetPixel hdc, tmpw + 25, 8, &HB07A46: SetPixel hdc, tmpw + 26, 8, &HB2773F: SetPixel hdc, tmpw + 27, 8, &HB36E2C: SetPixel hdc, tmpw + 28, 8, &HB26221: SetPixel hdc, tmpw + 29, 8, &HB35A20: SetPixel hdc, tmpw + 30, 8, &H9C3E11: SetPixel hdc, tmpw + 31, 8, &H5C2A1F: SetPixel hdc, tmpw + 32, 8, &HC4C1C0: SetPixel hdc, tmpw + 33, 8, &HFDFDFD: SetPixel hdc, tmpw + 34, 8, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 9, &HC69464: SetPixel hdc, tmpw + 18, 9, &HC69564: SetPixel hdc, tmpw + 19, 9, &HC3925E: SetPixel hdc, tmpw + 20, 9, &HC3915D: SetPixel hdc, tmpw + 21, 9, &HC3925E: SetPixel hdc, tmpw + 22, 9, &HC38F5B: SetPixel hdc, tmpw + 23, 9, &HC28D55: SetPixel hdc, tmpw + 24, 9, &HC08751: SetPixel hdc, tmpw + 25, 9, &HBC844C: SetPixel hdc, tmpw + 26, 9, &HBC8147: SetPixel hdc, tmpw + 27, 9, &HB57936: SetPixel hdc, tmpw + 28, 9, &HB3702D: SetPixel hdc, tmpw + 29, 9, &HB56626: SetPixel hdc, tmpw + 30, 9, &HA25115: SetPixel hdc, tmpw + 31, 9, &H662D12: SetPixel hdc, tmpw + 32, 9, &HAEA3A1: SetPixel hdc, tmpw + 33, 9, &HF9F9F9: SetPixel hdc, tmpw + 34, 9, &HFFFFFFFF:
    SetPixel hdc, tmpw + 17, 10, &HCC9B6A: SetPixel hdc, tmpw + 18, 10, &HCC9B6A: SetPixel hdc, tmpw + 19, 10, &HCA9864: SetPixel hdc, tmpw + 20, 10, &HC99763: SetPixel hdc, tmpw + 21, 10, &HC99763: SetPixel hdc, tmpw + 22, 10, &HCA9562: SetPixel hdc, tmpw + 23, 10, &HC9945D: SetPixel hdc, tmpw + 24, 10, &HC68E57: SetPixel hdc, tmpw + 25, 10, &HC48C55: SetPixel hdc, tmpw + 26, 10, &HC3884E: SetPixel hdc, tmpw + 27, 10, &HBE823E: SetPixel hdc, tmpw + 28, 10, &HB97634: SetPixel hdc, tmpw + 29, 10, &HBD6D2D: SetPixel hdc, tmpw + 30, 10, &HA8581C: SetPixel hdc, tmpw + 31, 10, &H6D3319: SetPixel hdc, tmpw + 32, 10, &HA49794: SetPixel hdc, tmpw + 33, 10, &HF6F6F6: SetPixel hdc, tmpw + 34, 10, &HFFFFFFFF:
    tmph = lh - 22
    tmpw = lw - 34
    SetPixel hdc, tmpw + 17, tmph + 10, &HCC9B6A: SetPixel hdc, tmpw + 18, tmph + 10, &HCC9B6A: SetPixel hdc, tmpw + 19, tmph + 10, &HCA9864: SetPixel hdc, tmpw + 20, tmph + 10, &HC99763: SetPixel hdc, tmpw + 21, tmph + 10, &HC99763: SetPixel hdc, tmpw + 22, tmph + 10, &HCA9562: SetPixel hdc, tmpw + 23, tmph + 10, &HC9945D: SetPixel hdc, tmpw + 24, tmph + 10, &HC68E57: SetPixel hdc, tmpw + 25, tmph + 10, &HC48C55: SetPixel hdc, tmpw + 26, tmph + 10, &HC3884E: SetPixel hdc, tmpw + 27, tmph + 10, &HBE823E: SetPixel hdc, tmpw + 28, tmph + 10, &HB97634: SetPixel hdc, tmpw + 29, tmph + 10, &HBD6D2D: SetPixel hdc, tmpw + 30, tmph + 10, &HA8581C: SetPixel hdc, tmpw + 31, tmph + 10, &H6D3319: SetPixel hdc, tmpw + 32, tmph + 10, &HA49794: SetPixel hdc, tmpw + 33, tmph + 10, &HF6F6F6:
    SetPixel hdc, tmpw + 17, tmph + 11, &HD0A175: SetPixel hdc, tmpw + 18, tmph + 11, &HD0A175: SetPixel hdc, tmpw + 19, tmph + 11, &HCC9D6D: SetPixel hdc, tmpw + 20, tmph + 11, &HCE9B6C: SetPixel hdc, tmpw + 21, tmph + 11, &HCB9A6A: SetPixel hdc, tmpw + 22, tmph + 11, &HCD996A: SetPixel hdc, tmpw + 23, tmph + 11, &HCA9666: SetPixel hdc, tmpw + 24, tmph + 11, &HCF9865: SetPixel hdc, tmpw + 25, tmph + 11, &HCA9460: SetPixel hdc, tmpw + 26, tmph + 11, &HC78F57: SetPixel hdc, tmpw + 27, tmph + 11, &HC1864B: SetPixel hdc, tmpw + 28, tmph + 11, &HC08143: SetPixel hdc, tmpw + 29, tmph + 11, &HB7712E: SetPixel hdc, tmpw + 30, tmph + 11, &HB16A28: SetPixel hdc, tmpw + 31, tmph + 11, &H694321: SetPixel hdc, tmpw + 32, tmph + 11, &HA59F9B: SetPixel hdc, tmpw + 33, tmph + 11, &HF4F4F4:
    SetPixel hdc, tmpw + 17, tmph + 12, &HDAAA7F: SetPixel hdc, tmpw + 18, tmph + 12, &HD9AB7E: SetPixel hdc, tmpw + 19, tmph + 12, &HDBAC7C: SetPixel hdc, tmpw + 20, tmph + 12, &HDDAA7B: SetPixel hdc, tmpw + 21, tmph + 12, &HDAA979: SetPixel hdc, tmpw + 22, tmph + 12, &HDAA677: SetPixel hdc, tmpw + 23, tmph + 12, &HD8A474: SetPixel hdc, tmpw + 24, tmph + 12, &HDBA471: SetPixel hdc, tmpw + 25, tmph + 12, &HD49F6A: SetPixel hdc, tmpw + 26, tmph + 12, &HD09861: SetPixel hdc, tmpw + 27, tmph + 12, &HD0955A: SetPixel hdc, tmpw + 28, tmph + 12, &HCC8D4F: SetPixel hdc, tmpw + 29, tmph + 12, &HCA8441: SetPixel hdc, tmpw + 30, tmph + 12, &HBB7532: SetPixel hdc, tmpw + 31, tmph + 12, &H7B5434: SetPixel hdc, tmpw + 32, tmph + 12, &HB1ACAA: SetPixel hdc, tmpw + 33, tmph + 12, &HF5F5F5:
    SetPixel hdc, tmpw + 17, tmph + 13, &HDCB087: SetPixel hdc, tmpw + 18, tmph + 13, &HDCB087: SetPixel hdc, tmpw + 19, tmph + 13, &HDBAF81: SetPixel hdc, tmpw + 20, tmph + 13, &HDEAF82: SetPixel hdc, tmpw + 21, tmph + 13, &HDCAF81: SetPixel hdc, tmpw + 22, tmph + 13, &HDDAD7F: SetPixel hdc, tmpw + 23, tmph + 13, &HDBAC7B: SetPixel hdc, tmpw + 24, tmph + 13, &HDDAA7A: SetPixel hdc, tmpw + 25, tmph + 13, &HD9A775: SetPixel hdc, tmpw + 26, tmph + 13, &HD7A26E: SetPixel hdc, tmpw + 27, tmph + 13, &HCE9961: SetPixel hdc, tmpw + 28, tmph + 13, &HCC945C: SetPixel hdc, tmpw + 29, tmph + 13, &HC2854D: SetPixel hdc, tmpw + 30, tmph + 13, &HAF7239: SetPixel hdc, tmpw + 31, tmph + 13, &H695C4F: SetPixel hdc, tmpw + 32, tmph + 13, &HD5D5D5: SetPixel hdc, tmpw + 33, tmph + 13, &HF8F8F8:
    SetPixel hdc, tmpw + 17, tmph + 14, &HE4B88E: SetPixel hdc, tmpw + 18, tmph + 14, &HE3B88E: SetPixel hdc, tmpw + 19, tmph + 14, &HE0B486: SetPixel hdc, tmpw + 20, tmph + 14, &HE4B689: SetPixel hdc, tmpw + 21, tmph + 14, &HE2B587: SetPixel hdc, tmpw + 22, tmph + 14, &HE5B588: SetPixel hdc, tmpw + 23, tmph + 14, &HE3B483: SetPixel hdc, tmpw + 24, tmph + 14, &HE2AF7F: SetPixel hdc, tmpw + 25, tmph + 14, &HDEAC7A: SetPixel hdc, tmpw + 26, tmph + 14, &HDEAA75: SetPixel hdc, tmpw + 27, tmph + 14, &HD8A36B: SetPixel hdc, tmpw + 28, tmph + 14, &HD09860: SetPixel hdc, tmpw + 29, tmph + 14, &HC58850: SetPixel hdc, tmpw + 30, tmph + 14, &HA56930: SetPixel hdc, tmpw + 31, tmph + 14, &H7B746C: SetPixel hdc, tmpw + 32, tmph + 14, &HE8E8E8: SetPixel hdc, tmpw + 33, tmph + 14, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 15, &HE1BF93: SetPixel hdc, tmpw + 18, tmph + 15, &HE2BE93: SetPixel hdc, tmpw + 19, tmph + 15, &HE2BF91: SetPixel hdc, tmpw + 20, tmph + 15, &HE1BD8F: SetPixel hdc, tmpw + 21, tmph + 15, &HE0BC8E: SetPixel hdc, tmpw + 22, tmph + 15, &HE2BC8E: SetPixel hdc, tmpw + 23, tmph + 15, &HE4BD8C: SetPixel hdc, tmpw + 24, tmph + 15, &HE0B685: SetPixel hdc, tmpw + 25, tmph + 15, &HDCB07E: SetPixel hdc, tmpw + 26, tmph + 15, &HE1AF7C: SetPixel hdc, tmpw + 27, tmph + 15, &HDEA66E: SetPixel hdc, tmpw + 28, tmph + 15, &HD19962: SetPixel hdc, tmpw + 29, tmph + 15, &HAD875D: SetPixel hdc, tmpw + 30, tmph + 15, &H7D6851: SetPixel hdc, tmpw + 31, tmph + 15, &HB9B9B9: SetPixel hdc, tmpw + 32, tmph + 15, &HF1F1F1:
    SetPixel hdc, tmpw + 17, tmph + 16, &HE7C699: SetPixel hdc, tmpw + 18, tmph + 16, &HE8C69A: SetPixel hdc, tmpw + 19, tmph + 16, &HE7C496: SetPixel hdc, tmpw + 20, tmph + 16, &HE8C597: SetPixel hdc, tmpw + 21, tmph + 16, &HE5C294: SetPixel hdc, tmpw + 22, tmph + 16, &HE8C194: SetPixel hdc, tmpw + 23, tmph + 16, &HE6BF8E: SetPixel hdc, tmpw + 24, tmph + 16, &HE7BC8C: SetPixel hdc, tmpw + 25, tmph + 16, &HE7BB8A: SetPixel hdc, tmpw + 26, tmph + 16, &HE5B37F: SetPixel hdc, tmpw + 27, tmph + 16, &HE1A971: SetPixel hdc, tmpw + 28, tmph + 16, &HD79F67: SetPixel hdc, tmpw + 29, tmph + 16, &H8E6A40: SetPixel hdc, tmpw + 30, tmph + 16, &H8A8076: SetPixel hdc, tmpw + 31, tmph + 16, &HE2E2E2: SetPixel hdc, tmpw + 32, tmph + 16, &HF9F9F9:
    SetPixel hdc, tmpw + 17, tmph + 17, &HE2CC9D: SetPixel hdc, tmpw + 18, tmph + 17, &HE2CC9E: SetPixel hdc, tmpw + 19, tmph + 17, &HE3CF9C: SetPixel hdc, tmpw + 20, tmph + 17, &HDFCA98: SetPixel hdc, tmpw + 21, tmph + 17, &HE2CD9C: SetPixel hdc, tmpw + 22, tmph + 17, &HE4CC9C: SetPixel hdc, tmpw + 23, tmph + 17, &HE1C491: SetPixel hdc, tmpw + 24, tmph + 17, &HE1C18F: SetPixel hdc, tmpw + 25, tmph + 17, &HE4BD8C: SetPixel hdc, tmpw + 26, tmph + 17, &HDAB285: SetPixel hdc, tmpw + 27, tmph + 17, &HC7A582: SetPixel hdc, tmpw + 28, tmph + 17, &H806A56: SetPixel hdc, tmpw + 29, tmph + 17, &H676565: SetPixel hdc, tmpw + 30, tmph + 17, &HD2D2D2: SetPixel hdc, tmpw + 31, tmph + 17, &HF2F2F2: SetPixel hdc, tmpw + 32, tmph + 17, &HFEFEFE:
    SetPixel hdc, tmpw + 17, tmph + 18, &HE9D5A6: SetPixel hdc, tmpw + 18, tmph + 18, &HE9D5A6: SetPixel hdc, tmpw + 19, tmph + 18, &HE9D6A3: SetPixel hdc, tmpw + 20, tmph + 18, &HE9D7A6: SetPixel hdc, tmpw + 21, tmph + 18, &HE2CC9B: SetPixel hdc, tmpw + 22, tmph + 18, &HE8D0A0: SetPixel hdc, tmpw + 23, tmph + 18, &HE8CF9C: SetPixel hdc, tmpw + 24, tmph + 18, &HE7C795: SetPixel hdc, tmpw + 25, tmph + 18, &HE2BC8B: SetPixel hdc, tmpw + 26, tmph + 18, &HB68F63: SetPixel hdc, tmpw + 27, tmph + 18, &H886948: SetPixel hdc, tmpw + 28, tmph + 18, &H786A5D: SetPixel hdc, tmpw + 29, tmph + 18, &HC7C7C7: SetPixel hdc, tmpw + 30, tmph + 18, &HEBEBEB: SetPixel hdc, tmpw + 31, tmph + 18, &HFCFCFC:
    SetPixel hdc, tmpw + 17, tmph + 19, &HD7D2BA: SetPixel hdc, tmpw + 18, tmph + 19, &HD7D2BA: SetPixel hdc, tmpw + 19, tmph + 19, &HD7D1B9: SetPixel hdc, tmpw + 20, tmph + 19, &HD5CEB6: SetPixel hdc, tmpw + 21, tmph + 19, &HDBD3BB: SetPixel hdc, tmpw + 22, tmph + 19, &HC9C1AA: SetPixel hdc, tmpw + 23, tmph + 19, &HA9A28B: SetPixel hdc, tmpw + 24, tmph + 19, &H827E6C: SetPixel hdc, tmpw + 25, tmph + 19, &H6A665B: SetPixel hdc, tmpw + 26, tmph + 19, &H625F5A: SetPixel hdc, tmpw + 27, tmph + 19, &H8B8C8C: SetPixel hdc, tmpw + 28, tmph + 19, &HCDCDCD: SetPixel hdc, tmpw + 29, tmph + 19, &HE8E8E8: SetPixel hdc, tmpw + 30, tmph + 19, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 20, &H5A563E: SetPixel hdc, tmpw + 18, tmph + 20, &H59543D: SetPixel hdc, tmpw + 19, tmph + 20, &H58513A: SetPixel hdc, tmpw + 20, tmph + 20, &H554F38: SetPixel hdc, tmpw + 21, tmph + 20, &H59513B: SetPixel hdc, tmpw + 22, tmph + 20, &H58513E: SetPixel hdc, tmpw + 23, tmph + 20, &H646053: SetPixel hdc, tmpw + 24, tmph + 20, &H7B7973: SetPixel hdc, tmpw + 25, tmph + 20, &HA2A19F: SetPixel hdc, tmpw + 26, tmph + 20, &HC5C5C5: SetPixel hdc, tmpw + 27, tmph + 20, &HDADADA: SetPixel hdc, tmpw + 28, tmph + 20, &HEDEDED: SetPixel hdc, tmpw + 29, tmph + 20, &HFAFAFA:
    SetPixel hdc, tmpw + 17, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 18, tmph + 21, &HC5C5C5: SetPixel hdc, tmpw + 19, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 20, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 21, tmph + 21, &HC6C6C6: SetPixel hdc, tmpw + 22, tmph + 21, &HC9C9C9: SetPixel hdc, tmpw + 23, tmph + 21, &HCECECE: SetPixel hdc, tmpw + 24, tmph + 21, &HD7D7D7: SetPixel hdc, tmpw + 25, tmph + 21, &HE1E1E1: SetPixel hdc, tmpw + 26, tmph + 21, &HECECEC: SetPixel hdc, tmpw + 27, tmph + 21, &HF6F6F6: SetPixel hdc, tmpw + 28, tmph + 21, &HFDFDFD:
    SetPixel hdc, tmpw + 17, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 18, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 19, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 20, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 21, tmph + 22, &HECECEC: SetPixel hdc, tmpw + 22, tmph + 22, &HEDEDED: SetPixel hdc, tmpw + 23, tmph + 22, &HF0F0F0: SetPixel hdc, tmpw + 24, tmph + 22, &HF4F4F4: SetPixel hdc, tmpw + 25, tmph + 22, &HFAFAFA: SetPixel hdc, tmpw + 26, tmph + 22, &HFDFDFD:
    tmph = 11:     tmph1 = lh - 10:     tmpw = lw - 34
    'Generar lineas intermedias
    DrawLineApi 0, tmph, 0, tmph1, &HF7F7F7: DrawLineApi 1, tmph, 1, tmph1, &HA99D9B: DrawLineApi 2, tmph, 2, tmph1, &H632D17: DrawLineApi 3, tmph, 3, tmph1, &HA65A1D: DrawLineApi 4, tmph, 4, tmph1, &HB96C2E
    DrawLineApi 5, tmph, 5, tmph1, &HBC7738: DrawLineApi 6, tmph, 6, tmph1, &HC18242: DrawLineApi 7, tmph, 7, tmph1, &HC2894E: DrawLineApi 8, tmph, 8, tmph1, &HC18A52: DrawLineApi 9, tmph, 9, tmph1, &HC59157
    DrawLineApi 10, tmph, 10, tmph1, &HC59159: DrawLineApi 11, tmph, 11, tmph1, &HCC9863: DrawLineApi 12, tmph, 12, tmph1, &HCC9665: DrawLineApi 13, tmph, 13, tmph1, &HCB9767: DrawLineApi 14, tmph, 14, tmph1, &HC99565
    DrawLineApi 15, tmph, 15, tmph1, &HCC9A6A: DrawLineApi 16, tmph, 16, tmph1, &HCC9A6A: DrawLineApi 17, tmph, 17, tmph1, &HCC9B6A: DrawLineApi tmpw + 17, tmph, tmpw + 17, tmph1, &HCC9B6A: DrawLineApi tmpw + 18, tmph, tmpw + 18, tmph1, &HCC9B6A:
    DrawLineApi tmpw + 19, tmph, tmpw + 19, tmph1, &HCA9864: DrawLineApi tmpw + 20, tmph, tmpw + 20, tmph1, &HC99763: DrawLineApi tmpw + 21, tmph, tmpw + 21, tmph1, &HC99763: DrawLineApi tmpw + 22, tmph, tmpw + 22, tmph1, &HCA9562: DrawLineApi tmpw + 23, tmph, tmpw + 23, tmph1, &HC9945D
    DrawLineApi tmpw + 24, tmph, tmpw + 24, tmph1, &HC68E57: DrawLineApi tmpw + 25, tmph, tmpw + 25, tmph1, &HC48C55: DrawLineApi tmpw + 26, tmph, tmpw + 26, tmph1, &HC3884E: DrawLineApi tmpw + 27, tmph, tmpw + 27, tmph1, &HBE823E: DrawLineApi tmpw + 28, tmph, tmpw + 28, tmph1, &HB97634
    DrawLineApi tmpw + 29, tmph, tmpw + 29, tmph1, &HBD6D2D: DrawLineApi tmpw + 30, tmph, tmpw + 30, tmph1, &HA8581C: DrawLineApi tmpw + 31, tmph, tmpw + 31, tmph1, &H6D3319: DrawLineApi tmpw + 32, tmph, tmpw + 32, tmph1, &HA49794: DrawLineApi tmpw + 33, tmph, tmpw + 33, tmph1, &HF6F6F6
    'Lineas verticales
    DrawLineApi 17, 0, lw - 17, 0, &H3C090A
    DrawLineApi 17, 1, lw - 17, 1, &HDEC7BE
    DrawLineApi 17, 2, lw - 17, 2, &HD3BCB2
    DrawLineApi 17, 3, lw - 17, 3, &HD4B39A
    DrawLineApi 17, 4, lw - 17, 4, &HCCAC92
    DrawLineApi 17, 5, lw - 17, 5, &HCFA88A
    DrawLineApi 17, 6, lw - 17, 6, &HCCA587
    DrawLineApi 17, 7, lw - 17, 7, &HD2A77D
    DrawLineApi 17, 8, lw - 17, 8, &HB98E62
    DrawLineApi 17, 9, lw - 17, 9, &HC69464
    DrawLineApi 17, 10, lw - 17, 10, &HCC9B6A
    DrawLineApi 17, 11, lw - 17, 11, &HD0A175
    tmph = lh - 22
    DrawLineApi 17, tmph + 11, lw - 17, tmph + 11, &HD0A175
    DrawLineApi 17, tmph + 12, lw - 17, tmph + 12, &HDAAA7F
    DrawLineApi 17, tmph + 13, lw - 17, tmph + 13, &HDCB087
    DrawLineApi 17, tmph + 14, lw - 17, tmph + 14, &HE4B88E
    DrawLineApi 17, tmph + 15, lw - 17, tmph + 15, &HE1BF93
    DrawLineApi 17, tmph + 16, lw - 17, tmph + 16, &HE7C699
    DrawLineApi 17, tmph + 17, lw - 17, tmph + 17, &HE2CC9D
    DrawLineApi 17, tmph + 18, lw - 17, tmph + 18, &HE9D5A6
    DrawLineApi 17, tmph + 19, lw - 17, tmph + 19, &HD7D2BA
    DrawLineApi 17, tmph + 20, lw - 17, tmph + 20, &H5A563E
    DrawLineApi 17, tmph + 21, lw - 17, tmph + 21, &HC5C5C5
    DrawLineApi 17, tmph + 22, lw - 17, tmph + 22, &HECECEC

End Sub

Private Sub CreateRegion()

'***************************************************************************
'*  Create region everytime you redraw a button.                           *
'*  Because some settings may have changed the button regions              *
'***************************************************************************

'If m_lButtonRgn Then DeleteObject m_lButtonRgn':(?-> replaced by:

    If m_lButtonRgn Then
        DeleteObject m_lButtonRgn
    End If
    m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 18, 18)
    SetWindowRgn UserControl.hWnd, m_lButtonRgn, True       'Set Button Region
    DeleteObject m_lButtonRgn                               'Free memory

End Sub

Private Sub DrawSymbol(ByVal eArrow As enumAquaSymbol)

Dim hOldFont As Long
Dim hNewFont As Long
Dim sSign As String
Dim BtnSymbol As enumAquaSymbol

    hNewFont = BuildSymbolFont(14)
    hOldFont = SelectObject(hdc, hNewFont)

    sSign = eArrow
    DrawText hdc, sSign, 1, lpSignRect, DT_WORDBREAK '!!
    DeleteObject hNewFont

End Sub

Private Function BuildSymbolFont(lFontSize As Long) As Long

Const SYMBOL_CHARSET = 2
Dim lpFont As tLOGFONT

    With lpFont
        .lfFaceName = "Marlett" + vbNullChar    'Standard Marlett Font
        .lfHeight = lFontSize                   'I was using Webdings first,
        .lfCharSet = SYMBOL_CHARSET             'but I am not sure whether
    End With                                    'it is installed in every machine!
    'Still Im not sure about Marlet :)
    BuildSymbolFont = CreateFontIndirect(lpFont) 'I got inspirations from
    'Light Templer's Project

End Function

Private Sub DrawPicwithCaption()

Dim lpRect   As RECT                        'RECT to draw caption
Dim pRect As RECT

    lw = ScaleWidth                         'ScaleHeight of Button
    lh = ScaleHeight                        'ScaleWidth of Button

    If (m_Buttonstate = eStateDown Or (m_ButtonMode <> ebmCommandButton And m_bValue = True)) Then
        '-- Mouse down
        If Not m_PictureDown Is Nothing Then
            Set tmppic = m_PictureDown
        Else
            If Not m_PictureHot Is Nothing Then
                Set tmppic = m_PictureHot
            Else
                Set tmppic = m_Picture
            End If
        End If
    ElseIf (m_Buttonstate = eStateOver) Then
        '-- Mouse in (over)
        If Not m_PictureHot Is Nothing Then
            Set tmppic = m_PictureHot
        Else
            Set tmppic = m_Picture
        End If
    Else
        '-- Mouse out (normal)
        Set tmppic = m_Picture
    End If

    ' --Adjust Picture Sizes
    PicH = ScaleX(tmppic.Height, vbHimetric, vbPixels)
    PicW = ScaleX(tmppic.Width, vbHimetric, vbPixels)

    ' --Get the drawing area of caption
    If m_DropDownSymbol <> ebsNone Or m_bDropDownSep Then
        If m_PictureAlign = epRightEdge Or m_PictureAlign = epRightOfCaption Then
            SetRect m_TextRect, 0, 0, lw - 24, lh
        Else
            SetRect m_TextRect, 0, 0, lw - 16, lh
        End If
    Else
        SetRect m_TextRect, 0, 0, lw - 8, lh
    End If

    ' --Calc rects for multiline
    If m_WindowsNT Then
        DrawTextW hdc, StrPtr(m_Caption), -1, m_TextRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTL, DT_RTLREADING, 0)
    Else
        DrawText hdc, m_Caption, -1, m_TextRect, DT_CALCRECT Or DT_WORDBREAK Or IIf(m_bRTL, DT_RTLREADING, 0)
    End If

    ' --Copy rect into temp var
    CopyRect lpRect, m_TextRect

    ' --Move the caption area according to Caption alignments
    Select Case m_CaptionAlign
    Case ecLeftAlign
        OffsetRect lpRect, 2, (lh - lpRect.Bottom) \ 2

    Case ecCenterAlign
        OffsetRect lpRect, (lw - lpRect.Right + PicW) \ 2, (lh - lpRect.Bottom) \ 2
        If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
            OffsetRect lpRect, -8, 0
        End If
        If m_PictureAlign = epBottomEdge Or m_PictureAlign = epBottomOfCaption Or m_PictureAlign = epTopOfCaption Or m_PictureAlign = epTopEdge Then
            OffsetRect lpRect, -(PicW \ 2), 0
        End If

    Case ecRightAlign
        OffsetRect lpRect, (lw - lpRect.Right - 4), (lh - lpRect.Bottom) \ 2

    End Select

    With lpRect

        If Not m_Picture Is Nothing Then
            Select Case m_PictureAlign
            Case epLeftEdge, epLeftOfCaption
                If m_CaptionAlign <> ecCenterAlign Then
                    If .Left < PicW + 4 Then
                        .Left = PicW + 4: .Right = .Right + PicW + 4
                    End If
                End If

            Case epRightEdge, epRightOfCaption
                If .Right > lw - PicW - 4 Then
                    .Right = lw - PicW - 4: .Left = .Left - PicW - 4
                End If
                If m_CaptionAlign = ecCenterAlign Then
                    OffsetRect lpRect, -12, 0
                End If

            Case epTopOfCaption, epTopEdge
                OffsetRect lpRect, 0, PicH \ 2

            Case epBottomOfCaption, epBottomEdge
                OffsetRect lpRect, 0, -PicH \ 2

            Case epBehindCaption
                If m_CaptionAlign = ecCenterAlign Then
                    OffsetRect lpRect, -16, 0
                End If
            End Select
        End If

        If m_CaptionAlign = ecRightAlign Then
            If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                OffsetRect lpRect, -16, 0
            End If
        End If

    End With

    ' --Save the caption rect
    CopyRect m_TextRect, lpRect
    OffsetRect m_TextRect, 0, -1

    With m_TextRect
        If .Top <= 4 Then
            .Top = 4
        End If
        If .Bottom >= ScaleHeight - 4 Then
            .Bottom = ScaleHeight - 4
        End If
        If .Right >= ScaleWidth - 4 Then
            .Right = ScaleWidth - 4
        End If
        If .Left <= 4 Then
            .Left = 4
        End If
    End With

    ' --Calculate Pictures positions once we have caption rects
    CalcPicRects

    ' --Calculate rects with the dropdown symbol
    If m_DropDownSymbol <> ebsNone Then
        ' --Drawing area for dropdown symbol  (the symbol is optional;)
        SetRect lpSignRect, lw - 15, lh / 2 - 7, lw, lh / 2 + 8
    End If

    If m_bDropDownSep Then
        If m_PictureAlign <> epRightEdge Or m_PictureAlign <> epRightOfCaption Then
            If m_TextRect.Right < ScaleWidth - 8 Then
                DrawLineApi lw - 16, 3, lw - 16, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), -0.1)
                DrawLineApi lw - 15, 3, lw - 15, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), 0.1)
            End If
        ElseIf m_PictureAlign = epRightEdge Or m_PictureAlign = epRightOfCaption Then
            DrawLineApi lw - 16, 3, lw - 16, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), -0.1)
            DrawLineApi lw - 15, 3, lw - 15, lh - 3, ShiftColor(GetPixel(hdc, 7, 7), 0.1)
        End If
    End If

   ' --Draw Pictures
   If m_bPicPushOnHover And m_Buttonstate = eStateOver Then
      DrawPicture m_PicRect, TranslateColor(&HC0C0C0)
      CopyRect pRect, m_PicRect
      OffsetRect pRect, -2, -2
      DrawPicture pRect
   Else
      DrawPicture m_PicRect
   End If

   If m_PictureShadow Then
      If Not (m_bPicPushOnHover And m_Buttonstate = eStateOver) Then
         DrawPicShadow
      End If
   End If

    ' --Text Effects
    If m_CaptionEffects <> eseNone Then
        DrawCaptionEffect
    End If

    ' --At Last, draw the Captions
    If m_bEnabled Then
        If m_Buttonstate = eStateOver Then
            DrawCaptionEx m_TextRect, TranslateColor(m_bColors.tForeColorOver), 0, 0
        Else
            DrawCaptionEx m_TextRect, TranslateColor(m_bColors.tForeColor), 0, 0
        End If
    Else
        DrawCaptionEx m_TextRect, GetSysColor(COLOR_GRAYTEXT), 0, 0
    End If

    If m_DropDownSymbol <> ebsNone Then
        DrawSymbol m_DropDownSymbol
    End If

End Sub

Private Sub CalcPicRects()

'If m_Picture Is Nothing Then Exit Sub':(?-> replaced by:

    If m_Picture Is Nothing Then
        Exit Sub
    End If

    With m_PicRect

        If Trim$(m_Caption) <> "" And m_PictureAlign <> epBehindCaption Then

            Select Case m_PictureAlign

            Case epLeftEdge
                .Left = 3
                .Top = (lh - PicH) \ 2
                If m_PicRect.Left < 0 Then
                    OffsetRect m_PicRect, PicW, 0
                    OffsetRect m_TextRect, PicW, 0
                End If

            Case epLeftOfCaption
                .Left = m_TextRect.Left - PicW - 4
                .Top = (lh - PicH) \ 2

            Case epRightEdge
                .Left = lw - PicW - 3
                .Top = (lh - PicH) \ 2
                ' --If picture overlaps text
                If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                    OffsetRect m_PicRect, -16, 0
                End If
                If .Left < m_TextRect.Right + 2 Then
                    .Left = m_TextRect.Right + 2
                End If

            Case epRightOfCaption
                .Left = m_TextRect.Right + 4
                .Top = (lh - PicH) \ 2
                If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                    OffsetRect m_PicRect, -16, 0
                End If
                ' --If picture overlaps text
                If .Left < m_TextRect.Right + 2 Then
                    .Left = m_TextRect.Right + 2
                End If

            Case epTopOfCaption
                .Left = (lw - PicW) \ 2
                .Top = m_TextRect.Top - PicH - 2
                If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                    OffsetRect m_PicRect, -8, 0
                End If

            Case epTopEdge
                .Left = (lw - PicW) \ 2
                .Top = 4
                If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                    OffsetRect m_PicRect, -8, 0
                End If

            Case epBottomOfCaption
                .Left = (lw - PicW) \ 2
                .Top = m_TextRect.Bottom + 2
                If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                    OffsetRect m_PicRect, -8, 0
                End If

            Case epBottomEdge
                .Left = (lw - PicW) \ 2
                .Top = lh - PicH - 4
                If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                    OffsetRect m_PicRect, -8, 0
                End If

            End Select
        Else
            .Left = (lw - PicW) \ 2
            .Top = (lh - PicH) \ 2
            If m_bDropDownSep Or m_DropDownSymbol <> ebsNone Then
                OffsetRect m_PicRect, -8, 0
            End If
        End If

        ' --Set the height and width
        If tmppic.Type = vbPicTypeIcon Then
            .Right = PicW
            .Bottom = PicH
        Else
            .Right = .Left + PicW
            .Bottom = .Top + PicH
        End If

        If .Left <= 1 Then
            .Left = 1
        End If
        If .Top <= 1 Then
            .Top = 1
        End If
        If .Bottom >= ScaleHeight - 4 Then
            .Bottom = ScaleHeight - 4
        End If
        If .Right >= ScaleWidth - 4 Then
            .Right = ScaleWidth - 4
        End If

    End With

End Sub

Private Sub DrawPicture(lpRect As RECT, Optional lBrushColor As Long = -1)

   Dim tmpMaskColor     As Long

   ' --Draw picture
   If tmppic.Type = vbPicTypeIcon Then
      tmpMaskColor = TranslateColor(&HC0C0C0)
   Else
      tmpMaskColor = m_lMaskColor
   End If
   
   If Is32BitBMP(tmppic) Then
      TransBlt32 hdc, lpRect.Left, lpRect.Top, PicW, PicH, tmppic, lBrushColor
   Else
      TransBlt hdc, lpRect.Left, lpRect.Top, PicW, PicH, tmppic, tmpMaskColor, lBrushColor
   End If
   
End Sub

Private Sub DrawPicShadow()

Dim bClr             As Long
Dim lShadowClr       As Long
Dim lPixelClr        As Long
Dim lpRect           As RECT

    If m_bPicPushOnHover And m_Buttonstate = eStateOver Then
        OffsetRect m_PicRect, -2, -2
    End If

    bClr = TranslateColor(m_bColors.tBackColor)

    ' --Get the pixel of the rightBottom corner of the picture and move slitghly from there --5 pixels
    lPixelClr = GetPixel(hdc, lpRect.Right + 5, lpRect.Bottom + 5)
    lShadowClr = BlendColors(TranslateColor(&HC0C0C0), lPixelClr)
    CopyRect lpRect, m_PicRect

    OffsetRect lpRect, 2, 2
    DrawPicture lpRect, ShiftColor(lShadowClr, -0.02)
    OffsetRect lpRect, -1, -1
    DrawPicture lpRect, ShiftColor(lShadowClr, -0.15)

    DrawPicture m_PicRect

End Sub

Private Sub DrawCaptionEffect()

'****************************************************************************
'* Draws the caption with/without unicode along with the special effects    *
'****************************************************************************

Dim bColor           As Long                                  'BackColor

    bColor = TranslateColor(m_bColors.tBackColor)

    ' --Set new colors according to effects
    Select Case m_CaptionEffects
    Case eseEmbossed
        DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.14), -1, -1
    Case eseEngraved
        DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.14), 1, 1
    Case eseShadowed
        DrawCaptionEx m_TextRect, TranslateColor(&HC0C0C0), 1, 1
    Case eseOutline
        DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), 1, 1
        DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), 1, -1
        DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), -1, 1
        DrawCaptionEx m_TextRect, ShiftColor(bColor, 0.1), -1, -1
    Case eseCover
        DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), 1, 1
        DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), 1, -1
        DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), -1, 1
        DrawCaptionEx m_TextRect, ShiftColor(bColor, -0.1), -1, -1

    End Select

    If m_bEnabled Then
        DrawCaptionEx m_TextRect, TranslateColor(m_bColors.tForeColor), 0, 0
    Else
        DrawCaptionEx m_TextRect, GetSysColor(COLOR_GRAYTEXT), 0, 0
    End If

End Sub

Private Sub DrawCaptionEx(lpRect As RECT, lColor As Long, OffsetX As Long, OffsetY As Long)

Dim tRect            As RECT
Dim lOldForeColor    As Long

' --Get current forecolor

    lOldForeColor = GetTextColor(hdc)

    CopyRect tRect, lpRect
    OffsetRect tRect, OffsetX, OffsetY

    SetTextColor hdc, lColor

    If m_WindowsNT Then
        DrawTextW hdc, StrPtr(m_Caption), -1, tRect, DT_DRAWFLAG Or IIf(m_bRTL, DT_RTLREADING, 0)
    Else
        DrawText hdc, m_Caption, -1, tRect, DT_DRAWFLAG Or IIf(m_bRTL, DT_RTLREADING, 0)
    End If

    ' --Restore previous forecolor
    SetTextColor hdc, lOldForeColor

End Sub

Private Sub UncheckAllValues()

' --Many Thanks to Morgan Haueisen

  Dim objButton As Object
   ' Check all controls in parent
   For Each objButton In Parent.Controls
       ' Is it a jcbutton?
      If TypeOf objButton Is jcbutton Then
         ' Is the button in the same container?
         If objButton.Container.hWnd = UserControl.ContainerHwnd Then
            ' is the button type Option?
            If objButton.Mode = [ebmOptionButton] Then
               ' is it not this button
               If Not objButton.hWnd = UserControl.hWnd Then
                  objButton.Value = False
               End If
            End If
         End If
      End If
    Next objButton
    
End Sub

Private Sub SetAccessKey()

Dim i As Long

    UserControl.AccessKeys = vbNullString
    If Len(m_Caption) > 1 Then
        i = InStr(1, m_Caption, "&", vbTextCompare)
        If (i < Len(m_Caption)) And (i > 0) Then
            If Mid$(m_Caption, i + 1, 1) <> "&" Then
                AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
            Else
                i = InStr(i + 2, m_Caption, "&", vbTextCompare)
                If Mid$(m_Caption, i + 1, 1) <> "&" Then
                    AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
                End If
            End If
        End If
    End If

End Sub

Private Sub DrawCorners(Color As Long)

'****************************************************************************
'* Draws four Corners of the button specified by Color                      *
'****************************************************************************

    lh = ScaleHeight
    lw = ScaleWidth

    SetPixel hdc, 1, 1, Color
    SetPixel hdc, 1, lh - 2, Color
    SetPixel hdc, lw - 2, 1, Color
    SetPixel hdc, lw - 2, lh - 2, Color

End Sub

Private Sub PaintRect(ByVal lColor As Long, lpRect As RECT)

'Fills a region with specified color

Dim hOldBrush   As Long
Dim hBrush      As Long

    hBrush = CreateSolidBrush(lColor)
    hOldBrush = SelectObject(hdc, hBrush)

    FillRect hdc, lpRect, hBrush

    SelectObject hdc, hOldBrush
    DeleteObject hBrush

End Sub

Private Sub ShowPopupMenu()

'* Shows a popupmenu
'* Inspired from Noel Dacara's dcbutton

Const TPM_BOTTOMALIGN As Long = &H20

Dim Menu        As VB.Menu
Dim Align       As enumAquaMenuAlign
Dim Flags       As Long
Dim DefaultMenu As VB.Menu

Dim x As Long
Dim Y As Long

    Set Menu = DropDownMenu
    Align = MenuAlign
    Flags = MenuFlags
    Set DefaultMenu = DefaultMenu

    lh = ScaleHeight: lw = ScaleWidth

    m_bPopupInit = True

    ' --Set the drop down menu position
    Select Case Align
    Case edaBottom
        Y = lh

    Case edaLeft, edaBottomLeft
        MenuFlags = MenuFlags Or vbPopupMenuRightAlign
        If (MenuAlign = edaBottomLeft) Then
            Y = lh
        End If

    Case edaRight, edaBottomRight
        x = lw
        If (MenuAlign = edaBottomRight) Then
            Y = lh
        End If

    Case edaTop, edaTopRight, edaTopLeft
        MenuFlags = TPM_BOTTOMALIGN
        If (MenuAlign = edaTopRight) Then
            x = lw
        ElseIf (MenuAlign = edaTopLeft) Then
            MenuFlags = MenuFlags Or vbPopupMenuRightAlign
        End If

    Case Else
        m_bPopupInit = False

    End Select

    If (m_bPopupInit) Then

        ' /--Show the dropdown menu
        If (DefaultMenu Is Nothing) Then
            UserControl.PopupMenu DropDownMenu, MenuFlags, x, Y
        Else
            UserControl.PopupMenu DropDownMenu, MenuFlags, x, Y, DefaultMenu
        End If

Dim lpPoint As POINT
        GetCursorPos lpPoint

        If (WindowFromPoint(lpPoint.x, lpPoint.Y) = UserControl.hWnd) Then
            m_bPopupShown = True
        Else
            m_bIsDown = False
            m_bMouseInCtl = False
            m_bIsSpaceBarDown = False
            m_Buttonstate = eStateNormal
            m_bPopupShown = False
            m_bPopupInit = False
            RedrawButton
        End If
    End If

End Sub

Private Function ShiftColor(Color As Long, PercentInDecimal As Single) As Long

'****************************************************************************
'* This routine shifts a color value specified by PercentInDecimal          *
'* Function inspired from DCbutton                                          *
'* All Credits goes to Noel Dacara                                          *
'* A Littlebit modified by me                                               *
'****************************************************************************

Dim r As Long
Dim g As Long
Dim B As Long

'  Add or remove a certain color quantity by how many percent.

    r = Color And 255
    g = (Color \ 256) And 255
    B = (Color \ 65536) And 255

    r = r + PercentInDecimal * 255       ' Percent should already
    g = g + PercentInDecimal * 255       ' be translated.
    B = B + PercentInDecimal * 255       ' Ex. 50% -> 50 / 100 = 0.5

    '  When overflow occurs, ....
    If (PercentInDecimal > 0) Then       ' RGB values must be between 0-255 only
        'If (r > 255) Then r = 255':(?-> replaced by:
        If (r > 255) Then
            r = 255
        End If
        'If (g > 255) Then g = 255':(?-> replaced by:
        If (g > 255) Then
            g = 255
        End If
        'If (B > 255) Then B = 255':(?-> replaced by:
        If (B > 255) Then
            B = 255
        End If
    Else
        'If (r < 0) Then r = 0':(?-> replaced by:
        If (r < 0) Then
            r = 0
        End If
        'If (g < 0) Then g = 0':(?-> replaced by:
        If (g < 0) Then
            g = 0
        End If
        'If (B < 0) Then B = 0':(?-> replaced by:
        If (B < 0) Then
            B = 0
        End If
    End If

    ShiftColor = r + 256& * g + 65536 * B ' Return shifted color value

End Function

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    If m_bEnabled Then                           'Disabled?? get out!!
        If m_bIsSpaceBarDown Then
            m_bIsSpaceBarDown = False
            m_bIsDown = False
        End If
        If m_ButtonMode = ebmCheckBox Then       'Checkbox Mode?
            'If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape':(?-> replaced by:
            If KeyAscii = 13 Or KeyAscii = 27 Then
                Exit Sub 'Checkboxes dont repond to Enter/Escape
            End If
            m_bValue = Not m_bValue             'Change Value (Checked/Unchecked)
            If Not m_bValue Then                'If value unchecked then
                m_Buttonstate = eStateNormal     'Normal State
            End If
            RedrawButton
        ElseIf m_ButtonMode = ebmOptionButton Then
            'If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape':(?-> replaced by:
            If KeyAscii = 13 Or KeyAscii = 27 Then
                Exit Sub 'Checkboxes dont repond to Enter/Escape
            End If
            UncheckAllValues
            m_bValue = True
            RedrawButton
        End If
        DoEvents                               'To remove focus from other button and Do events before click event
        RaiseEvent Click                       'Now Raiseevent
    End If

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    m_bDefault = Ambient.DisplayAsDefault
    If PropertyName = "DisplayAsDefault" Then
        RedrawButton
    End If

    If PropertyName = "BackColor" Then
        RedrawButton
    End If

End Sub

Private Sub UserControl_DblClick()

    If m_bHandPointer Then
        SetCursor m_lCursor
    End If

    If m_lDownButton = 1 Then                    'React to only Left button

        SetCapture (hWnd)                         'Preserve Hwnd on DoubleClick
        'If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown':(?-> replaced by:
        If m_Buttonstate <> eStateDown Then
            m_Buttonstate = eStateDown
        End If
        RedrawButton
        UserControl_MouseDown m_lDownButton, m_lDShift, m_lDX, m_lDY
        If Not m_bPopupEnabled Then
            RaiseEvent DblClick
        Else
            If Not m_bPopupShown Then
                ShowPopupMenu
            End If
        End If
    End If

End Sub

Private Sub UserControl_GotFocus()

    m_bHasFocus = True
    If m_bMouseInCtl Then
        'If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver':(?-> replaced by:
        If m_Buttonstate <> eStateOver Then
            m_Buttonstate = eStateOver
        End If
    Else
        'If Not m_bIsDown Then m_Buttonstate = eStateNormal':(?-> replaced by:
        If Not m_bIsDown Then
            m_Buttonstate = eStateNormal
        End If
    End If

End Sub

Private Sub UserControl_Hide()

    UserControl.Extender.ToolTipText = m_sTooltipText

End Sub

Private Sub UserControl_Initialize()

Dim i                As Long
Dim OS               As OSVERSIONINFO

'Prebuid Lighten/Darken arrays

    For i = 0 To 255
        aLighten(i) = Lighten(i)
        aDarken(i) = Darken(i)
    Next

    ' --Get the operating system version for text drawing purposes.
    m_hMode = LoadLibraryA("shell32.dll")
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    m_WindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub UserControl_InitProperties()

'Initialize Properties for User Control
'Called on designtime everytime a control is added

    m_bEnabled = True
    m_Caption = Ambient.DisplayName
    UserControl.FontName = "Tahoma"
    Set mFont = UserControl.Font
    mFont_FontChanged ("")
    m_PictureOpacity = 255
    m_PicOpacityOnOver = 255
    m_PictureAlign = epLeftOfCaption
    m_bUseMaskColor = True
    m_lMaskColor = &HE0E0E0
    m_CaptionAlign = ecCenterAlign
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    Refresh

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case 13                                    'Enter Key
        RaiseEvent Click
    Case 37, 38                                'Left and Up Arrows
        SendKeys "+{TAB}"                      'Button should transfer focus to other ctl
    Case 39, 40                                'Right and Down Arrows
        SendKeys "{TAB}"                       'Button should transfer focus to other ctl
    Case 32                                    'SpaceBar held down
        'If Shift = 4 Then Exit Sub             'System Menu Should pop up':(?-> replaced by:
        If Shift = 4 Then
            Exit Sub             'System Menu Should pop up
        End If
        If Not m_bIsDown Then
            m_bIsSpaceBarDown = True           'Set space bar as pressed

            If (m_ButtonMode = ebmCheckBox) Then 'Is CheckBoxMode??
                m_bValue = Not m_bValue         'Toggle Check Value
            ElseIf m_ButtonMode = ebmOptionButton Then
                UncheckAllValues                'Option Button Mode
                m_bValue = True                 'Pressed button Checked
            End If

            If m_Buttonstate <> eStateDown Then
                m_Buttonstate = eStateDown 'Button state should be down
                RedrawButton
            End If
        End If

        If (Not GetCapture = UserControl.hWnd) Then
            ReleaseCapture
            SetCapture UserControl.hWnd     'No other processing until spacebar is released
        End If                              'Thanks to APIGuide

    Case Else
        If m_bIsSpaceBarDown Then
            m_bIsSpaceBarDown = False
            m_Buttonstate = eStateNormal
            RedrawButton
        End If
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

' --Simply raise the event =)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
        
        ReleaseCapture                          'Now you can process further
                                                'as the spacebar is released
        If m_bMouseInCtl And m_bIsDown Then
            If m_Buttonstate <> eStateDown Then
                m_Buttonstate = eStateDown
                RedrawButton
            End If
        ElseIf m_bMouseInCtl Then                'If spacebar released over ctl
            If m_Buttonstate <> eStateOver Then
                m_Buttonstate = eStateOver 'Draw Hover State
                RedrawButton
            End If
            If Not m_bIsDown And m_bIsSpaceBarDown Then
                RaiseEvent Click
            End If
        Else                                         'If Spacebar released outside ctl
            If m_Buttonstate <> eStateNormal Then
                m_Buttonstate = eStateNormal
                RedrawButton
            End If
            If Not m_bIsDown And m_bIsSpaceBarDown Then
                RaiseEvent Click
            End If
        End If

        RaiseEvent KeyUp(KeyCode, Shift)
        m_bIsSpaceBarDown = False
        m_bIsDown = False
    End If

End Sub

Private Sub UserControl_LostFocus()

   m_bHasFocus = False                                 'No focus
   m_bIsDown = False                                   'No down state
   m_bIsSpaceBarDown = False                           'No spacebar held
   If Not m_bParentActive Then
      If m_Buttonstate <> eStateNormal Then
          m_Buttonstate = eStateNormal
      End If
   ElseIf m_bMouseInCtl Then
      If m_Buttonstate <> eStateOver Then
         m_Buttonstate = eStateOver
      End If
   Else
      If m_Buttonstate <> eStateNormal Then
         m_Buttonstate = eStateNormal
      End If
   End If
   RedrawButton

   If m_bDefault Then                                  'If default button,
      RedrawButton                                    'Show Focus
   End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    m_lDownButton = Button                       'Button pressed for Dblclick
    m_lDX = x
    m_lDY = Y
    m_lDShift = Shift

    ' --Set HandPointer if any!
    If m_bHandPointer Then
        SetCursor m_lCursor
    End If

    If Button = vbLeftButton Or m_bPopupShown Then
        m_bHasFocus = True
        m_bIsDown = True

        If (Not m_bIsSpaceBarDown) Then
            If m_Buttonstate <> eStateDown Then
                m_Buttonstate = eStateDown
                RedrawButton
            End If
        End If

        If Not m_bPopupEnabled Then
            RaiseEvent MouseDown(Button, Shift, x, Y)
        Else
            If Not m_bPopupShown Then
                ShowPopupMenu
            End If
        End If
    End If

End Sub

Private Sub CreateToolTip()

'****************************************************************************
'* A very nice and flexible sub to create balloon tool tips
'* Author :- Fred.CPP
'* Added as requested by many users
'* Modified by me to support unicode
'* Thanks Alfredo ;)
'****************************************************************************

Dim lpRect           As RECT
Dim lWinStyle        As Long
Dim lPtr             As Long
Dim ttip             As TOOLINFO
Dim ttipW            As TOOLINFOW
Const CS_DROPSHADOW     As Long = &H20000
Const GCL_STYLE         As Long = (-26)

' --Dont show tooltips if disabled

'If (Not m_bEnabled) Or m_bPopupShown Or m_Buttonstate = eStateDown Then Exit Sub':(?-> replaced by:

    If (Not m_bEnabled) Or m_bPopupShown Or m_Buttonstate = eStateDown Then
        Exit Sub
    End If

    ' --Destroy any previous tooltip
    If m_lttHwnd <> 0 Then
        DestroyWindow m_lttHwnd
    End If

    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX

    ''create baloon style if desired
    'If m_lTooltipType = TooltipBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON':(?-> replaced by:
    If m_lTooltipType = TooltipBalloon Then
        lWinStyle = lWinStyle Or TTS_BALLOON
    End If

    If m_bttRTL Then
        m_lttHwnd = CreateWindowEx(WS_EX_LAYOUTRTL, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
    Else
        m_lttHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hWnd, 0&, App.hInstance, 0&)
    End If

    SetClassLong m_lttHwnd, GCL_STYLE, GetClassLong(m_lttHwnd, GCL_STYLE) Or CS_DROPSHADOW

    'make our tooltip window a topmost window
    'SetWindowPos m_lttHwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE

    ''get the rect of the parent control
    GetClientRect UserControl.hWnd, lpRect

    If m_WindowsNT Then
        ' --set our tooltip info structure  for UNICODE SUPPORT >> WinNT
        With ttipW
            ' --if we want it centered, then set that flag
            If m_lttCentered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
            Else
                .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
            End If

            ' --set the hwnd prop to our parent control's hwnd
            .lHwnd = UserControl.hWnd
            .lId = hWnd
            .lSize = Len(ttipW)
            .hInstance = App.hInstance
            .lpStrW = StrPtr(m_sTooltipText)
            .lpRect = lpRect
        End With
        ' --add the tooltip structure
        SendMessage m_lttHwnd, TTM_ADDTOOLW, 0&, ttipW
    Else
        ' --set our tooltip info structure for << WinNT
        With ttip
            ''if we want it centered, then set that flag
            If m_lttCentered Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
            Else
                .lFlags = TTF_SUBCLASS
            End If

            ' --set the hwnd prop to our parent control's hwnd
            .lHwnd = UserControl.hWnd
            .lId = hWnd
            .lSize = Len(ttip)
            .hInstance = App.hInstance
            .lpStr = m_sTooltipText
            .lpRect = lpRect
        End With
        ' --add the tooltip structure
        SendMessage m_lttHwnd, TTM_ADDTOOLA, 0&, ttip
    End If

    'if we want a title or we want an icon
    If m_sTooltiptitle <> vbNullString Or m_lToolTipIcon <> TTNoIcon Then
        If m_WindowsNT Then
            lPtr = StrPtr(m_sTooltiptitle)
            If lPtr Then
                SendMessage m_lttHwnd, TTM_SETTITLEW, m_lToolTipIcon, ByVal lPtr
            End If
        Else
            SendMessage m_lttHwnd, TTM_SETTITLE, CLng(m_lToolTipIcon), ByVal m_sTooltiptitle
        End If

    End If
    SendMessage m_lttHwnd, TTM_SETMAXTIPWIDTH, 0, 240 'for Multiline capability
    If m_lttBackColor <> Empty Then
        SendMessage m_lttHwnd, TTM_SETTIPBKCOLOR, TranslateColor(m_lttBackColor), 0&
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim lp               As POINT

    GetCursorPos lp
    ' --Set hand pointer if any!
    If m_bHandPointer Then
        SetCursor m_lCursor
    End If

    If Not (WindowFromPoint(lp.x, lp.Y) = UserControl.hWnd) Then
        ' --Mouse yet not entered in the control
        m_bMouseInCtl = False
    Else
        m_bMouseInCtl = True
        ' --Check when the Mouse leaves the control
        TrackMouseLeave hWnd
        ' --Raise a MouseEnter event(it's Same as mouseMove)
        RaiseEvent MouseEnter
    End If

    ' --Proceed only if spacebar is not pressed
    'If m_bIsSpaceBarDown Then Exit Sub':(?-> replaced by:
    If m_bIsSpaceBarDown Then
        Exit Sub
    End If

    ' --We are inside button
    If m_bMouseInCtl Then

        ' --Mouse button is pressed down
        If m_bIsDown Then
            If m_Buttonstate <> eStateDown Then
                m_Buttonstate = eStateDown
                RedrawButton
            End If
        Else
            ' --Button should be in hot state if user leaves the button
            ' --with mouse button pressed
            If m_Buttonstate <> eStateOver Then
                m_Buttonstate = eStateOver
                RedrawButton
                ' --Create Tooltip Here
                If m_Buttonstate <> eStateDown Then
                    CreateToolTip
                End If
            End If
        End If

    Else
        If m_Buttonstate <> eStateNormal Then
            m_Buttonstate = eStateNormal
            RedrawButton
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, x, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If m_bHandPointer Then
        SetCursor m_lCursor
    End If

    ' --Popupmenu enabled
    If m_bPopupEnabled Then
        m_bIsDown = False
        m_bPopupShown = False
        m_Buttonstate = eStateNormal
        RedrawButton
        Exit Sub
    End If

    ' --React only to Left mouse button
    If Button = vbLeftButton Then
        '--Button released
        m_bIsDown = False
        ' --If button released in button area
        If (x > 0 And Y > 0) And (x < ScaleWidth And Y < ScaleHeight) Then

            ' --If check box mode
            If m_ButtonMode = ebmCheckBox Then
                m_bValue = Not m_bValue
                RedrawButton
                ' --If option button mode
            ElseIf m_ButtonMode = ebmOptionButton Then
                UncheckAllValues
                m_bValue = True
            End If

            ' --redraw Normal State
            m_Buttonstate = eStateNormal
            RedrawButton
            RaiseEvent Click

        End If
    End If

    RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

Private Sub UserControl_Resize()

' --At least, a checkbox will also need this much of size!!!!

'If Height < 345 Then Height = 345':(?-> replaced by:

    If Height < 345 Then
        Height = 345
    End If
    'If Width < 615 Then Width = 615':(?-> replaced by:
    If Width < 615 Then
        Width = 615
    End If

    ' --On resize, create button region again
    CreateRegion
    RedrawButton                'then redraw

End Sub

Private Sub UserControl_Paint()

' --this routine typically called by Windows when another window covering
'   this button is removed, or when the parent is moved/minimized/etc.

    RedrawButton

End Sub

'Load property values from storage

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set mFont = .ReadProperty("Font", Ambient.Font)
        Set UserControl.Font = mFont
        m_bEnabled = .ReadProperty("Enabled", True)
        m_Caption = .ReadProperty("Caption", "jcbutton")
        m_bValue = .ReadProperty("Value", False)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        m_bHandPointer = .ReadProperty("HandPointer", False)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Set m_Picture = .ReadProperty("PictureNormal", Nothing)
        Set m_PictureHot = .ReadProperty("PictureHot", Nothing)
        Set m_PictureDown = .ReadProperty("PictureDown", Nothing)
        m_PicEffectonOver = .ReadProperty("PictureEffectOnOver", epeLighter)
        m_PicEffectonDown = .ReadProperty("PictureEffectOnDown", epeDarker)
        m_bPicPushOnHover = .ReadProperty("PicturePushOnHover", False)
        m_PictureShadow = .ReadProperty("PictureShadow", False)
        m_PicOpacityOnOver = .ReadProperty("PictureOpacityOnOver", 255)
        m_PictureOpacity = .ReadProperty("PictureOpacity", 255)
        m_PicDisabledMode = .ReadProperty("DisabledPictureMode", edpBlended)
        m_lMaskColor = .ReadProperty("MaskColor", &HE0E0E0)
        m_bUseMaskColor = .ReadProperty("UseMaskColor", True)
        m_CaptionEffects = .ReadProperty("CaptionEffects", eseNone)
        m_ButtonMode = .ReadProperty("Mode", ebmCommandButton)
        m_PictureAlign = .ReadProperty("PictureAlign", epLeftOfCaption)
        m_CaptionAlign = .ReadProperty("CaptionAlign", ecCenterAlign)
        m_bColors.tForeColor = .ReadProperty("ForeColor", TranslateColor(vbButtonText))
        m_bColors.tForeColorOver = .ReadProperty("ForeColorHover", TranslateColor(vbButtonText))
        UserControl.ForeColor = m_bColors.tForeColor
        m_bDropDownSep = .ReadProperty("DropDownSeparator", False)
        m_sTooltipText = .ReadProperty("ToolTip", vbNullString)
        m_sTooltiptitle = .ReadProperty("TooltipTitle", vbNullString)
        m_lToolTipIcon = .ReadProperty("TooltipIcon", TTNoIcon)
        m_lttBackColor = .ReadProperty("TooltipBackColor", TranslateColor(vbInfoBackground))
        m_lTooltipType = .ReadProperty("TooltipType", TooltipStandard)
        m_bRTL = .ReadProperty("RightToLeft", False)
        m_bttRTL = .ReadProperty("RightToLeft", False)
        m_DropDownSymbol = .ReadProperty("DropDownSymbol", ebsNone)
        UserControl.Enabled = m_bEnabled
        SetAccessKey
        lh = UserControl.ScaleHeight
        lw = UserControl.ScaleWidth
        m_lParenthWnd = UserControl.Parent.hWnd
    End With

    UserControl_Resize

    If Ambient.UserMode Then                                                              'If we're not in design mode

        If m_bHandPointer Then
            m_lCursor = LoadCursor(0, IDC_HAND)     'Load System Hand pointer
            m_bHandPointer = (Not m_lCursor = 0)
        End If

        On Error GoTo h:

        If Ambient.UserMode Then                                                              'If we're not in design mode
            TrackUser32 = IsFunctionSupported("TrackMouseEvent", "User32")

            'If Not TrackUser32 Then IsFunctionSupported "_TrackMouseEvent", "ComCtl32"':(?-> replaced by:
            If Not TrackUser32 Then
                IsFunctionSupported "_TrackMouseEvent", "ComCtl32"
            End If

            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                Subclass_Initialize .hWnd
                Subclass_Initialize m_lParenthWnd
                Subclass_AddMsg .hWnd, WM_MOUSELEAVE, MSG_AFTER
                Subclass_AddMsg .hWnd, WM_THEMECHANGED, MSG_AFTER
                On Error Resume Next
                    If UserControl.Parent.MDIChild Then
                        Subclass_AddMsg m_lParenthWnd, WM_NCACTIVATE, MSG_AFTER
                    Else
                        Subclass_AddMsg m_lParenthWnd, WM_ACTIVATE, MSG_AFTER
                    End If
                End With
            End If
        End If

h:

End Sub

Private Sub UserControl_Show()

    UserControl.Extender.ToolTipText = vbNullString

End Sub

'A nice place to stop subclasser

Private Sub UserControl_Terminate()

    On Error GoTo Crash:
    'If m_lButtonRgn Then DeleteObject m_lButtonRgn      'Delete button region':(?-> replaced by:
    If m_lButtonRgn Then
        DeleteObject m_lButtonRgn      'Delete button region
    End If
    Set mFont = Nothing                                 'Clean up Font (StdFont)
    FreeLibrary m_hMode
    UnsetPopupMenu
    If Ambient.UserMode Then
        Subclass_Terminate
        Subclass_Terminate
    End If
Crash:

End Sub

'Write property values to storage

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "ShowFocusRect", m_bShowFocus, False
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "Font", mFont, Ambient.Font
        .WriteProperty "Caption", m_Caption, "jcbutton1"
        .WriteProperty "ForeColor", m_bColors.tForeColor, TranslateColor(vbButtonText)
        .WriteProperty "ForeColorHover", m_bColors.tForeColorOver, TranslateColor(vbButtonText)
        .WriteProperty "Mode", m_ButtonMode, ebmCommandButton
        .WriteProperty "Value", m_bValue, False
        .WriteProperty "MousePointer", UserControl.MousePointer, 0
        .WriteProperty "HandPointer", m_bHandPointer, False
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "PictureNormal", m_Picture, Nothing
        .WriteProperty "PictureHot", m_PictureHot, Nothing
        .WriteProperty "PictureDown", m_PictureDown, Nothing
        .WriteProperty "PictureAlign", m_PictureAlign, epLeftOfCaption
        .WriteProperty "PictureShadow", m_PictureShadow, False
        .WriteProperty "PictureOpacity", m_PictureOpacity, 255
        .WriteProperty "PictureOpacityOnOver", m_PicOpacityOnOver, 255
        .WriteProperty "DisabledPictureMode", m_PicDisabledMode, edpBlended
        .WriteProperty "PicturePushOnHover", m_bPicPushOnHover, False
        .WriteProperty "PictureEffectOnOver", m_PicEffectonOver, epeLighter
        .WriteProperty "PictureEffectOnDown", m_PicEffectonDown, epeDarker
        .WriteProperty "CaptionEffects", m_CaptionEffects, vbNullString
        .WriteProperty "UseMaskCOlor", m_bUseMaskColor, True
        .WriteProperty "MaskColor", m_lMaskColor, &HE0E0E0
        .WriteProperty "CaptionAlign", m_CaptionAlign, ecCenterAlign
        .WriteProperty "ToolTip", m_sTooltipText, vbNullString
        .WriteProperty "TooltipType", m_lTooltipType, TooltipStandard
        .WriteProperty "TooltipIcon", m_lToolTipIcon, TTNoIcon
        .WriteProperty "TooltipTitle", m_sTooltiptitle, vbNullString
        .WriteProperty "TooltipBackColor", m_lttBackColor, TranslateColor(vbInfoBackground)
        .WriteProperty "RightToLeft", m_bRTL, False
        .WriteProperty "DropDownSymbol", m_DropDownSymbol, ebsNone
        .WriteProperty "DropDownSeparator", m_bDropDownSep, False
    End With

End Sub

Private Function Is32BitBMP(obj As Object) As Boolean

Dim uBI              As BITMAP

    If obj.Type = vbPicTypeBitmap Then
        GetObject obj.Handle, Len(uBI), uBI
        Is32BitBMP = uBI.bmBitsPixel = 32
    End If

End Function

'Determine if the passed function is supported

Private Function IsFunctionSupported(ByVal sFunction As String, ByVal sModule As String) As Boolean

Dim lngModule As Long

    lngModule = GetModuleHandle(sModule)

    'If lngModule = 0 Then lngModule = LoadLibraryA(sModule)':(?-> replaced by:
    If lngModule = 0 Then
        lngModule = LoadLibraryA(sModule)
    End If

    If lngModule Then
        IsFunctionSupported = GetProcAddress(lngModule, sFunction)
        FreeLibrary lngModule
    End If

End Function

'Track the mouse leaving the indicated window

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

Dim tme              As TRACKMOUSEEVENT_STRUCT

    If TrackUser32 Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If TrackUser32 Then
            TrackMouseEvent tme
        Else
            TrackMouseEventComCtl tme
        End If
    End If

End Sub

'=========================================================================
'PUBLIC ROUTINES including subclassing & public button properties

' CREDITS: Paul Caton
'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lHwnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

'Parameters:
'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'hWnd     - The window handle
'uMsg     - The message number
'wParam   - Message related data
'lParam   - Message related data
'Notes:
'If you really know what you're doing, it's possible to change the values of the
'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'values get passed to the default handler.. and optionaly, the 'after' callback

Static bMoving As Boolean

    Select Case uMsg

    Case WM_MOUSELEAVE

        m_bMouseInCtl = False
        If m_bPopupEnabled Then
            If m_bPopupInit Then
                m_bPopupInit = False
                m_bPopupShown = True
                Exit Sub
            Else
                m_bPopupShown = False
            End If
        End If

        'If m_bIsSpaceBarDown Then Exit Sub':(?-> replaced by:
        If m_bIsSpaceBarDown Then
            Exit Sub
        End If
        If m_Buttonstate <> eStateNormal Then
            m_Buttonstate = eStateNormal
            RedrawButton
        End If
        RaiseEvent MouseLeave

    Case WM_NCACTIVATE, WM_ACTIVATE
        If wParam Then
            m_bParentActive = True
            'If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal':(?-> replaced by:
            If m_Buttonstate <> eStateNormal Then
                m_Buttonstate = eStateNormal
            End If
            If m_bDefault Then
                RedrawButton
            End If
            RedrawButton
        Else
            m_bIsDown = False
            m_bIsSpaceBarDown = False
            m_bHasFocus = False
            m_bParentActive = False
            'If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal':(?-> replaced by:
            If m_Buttonstate <> eStateNormal Then
                m_Buttonstate = eStateNormal
            End If
            RedrawButton
        End If

    Case WM_THEMECHANGED
        RedrawButton

    Case WM_SYSCOLORCHANGE
        RedrawButton
    End Select

End Sub

Public Sub SetPopupMenu(Menu As Object, Optional Align As enumAquaMenuAlign, Optional Flags = 0, Optional DefaultMenu = Nothing)
Attribute SetPopupMenu.VB_Description = "Sets a dropdown menu to the button."

    If Not (Menu Is Nothing) Then
        If (TypeOf Menu Is VB.Menu) Then

            Set DropDownMenu = Menu
            MenuAlign = Align
            MenuFlags = Flags
            Set DefaultMenu = DefaultMenu
            m_bPopupEnabled = True
        End If
    End If

End Sub

Public Sub UnsetPopupMenu()

' --Free the popup menu

    Set DropDownMenu = Nothing
    Set DefaultMenu = Nothing
    m_bPopupEnabled = False
    m_bPopupShown = False

End Sub

Public Sub About()
Attribute About.VB_Description = "Displays information about the control and its author."
Attribute About.VB_UserMemId = -552

    MsgBox "AquaButton (MAC OS X)" & vbNewLine & _
                        "Author: Juned S. Chhipa" & vbNewLine & _
                        "Contact: juned.chhipa@yahoo.com" & vbNewLine & vbNewLine & _
                        "Copyright ?2008-2009 Juned Chhipa. All rights reserved.", vbInformation + vbOKOnly, "About"

End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the button."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    SetAccessKey
    RedrawButton
    PropertyChanged "Caption"

End Property

Public Property Get CaptionAlign() As enumAquaCaptionAlign
Attribute CaptionAlign.VB_Description = "Returns/Sets the position of the Caption."
Attribute CaptionAlign.VB_ProcData.VB_Invoke_Property = ";Position"

    CaptionAlign = m_CaptionAlign

End Property

Public Property Let CaptionAlign(ByVal New_CaptionAlign As enumAquaCaptionAlign)

    m_CaptionAlign = New_CaptionAlign
    RedrawButton
    PropertyChanged "CaptionAlign"

End Property

Public Property Get DisabledPictureMode() As enumAquaDisabledPicMode
Attribute DisabledPictureMode.VB_Description = "Returns/Sets the effect to be used for picture when button is disabled."
Attribute DisabledPictureMode.VB_ProcData.VB_Invoke_Property = ";Appearance"

    DisabledPictureMode = m_PicDisabledMode

End Property

Public Property Let DisabledPictureMode(ByVal New_mode As enumAquaDisabledPicMode)

    m_PicDisabledMode = New_mode
    RedrawButton
    PropertyChanged "DisabledPictureMode"

End Property

Public Property Get DropDownSymbol() As enumAquaSymbol
Attribute DropDownSymbol.VB_Description = "Returns/Sets the Symbol to be used for displaying PopupMenu."
Attribute DropDownSymbol.VB_ProcData.VB_Invoke_Property = ";Appearance"

    DropDownSymbol = m_DropDownSymbol

End Property

Public Property Let DropDownSymbol(ByVal New_Align As enumAquaSymbol)

    m_DropDownSymbol = New_Align
    RedrawButton
    PropertyChanged "DropDownSymbol"

End Property

Public Property Get DropDownSeparator() As Boolean
Attribute DropDownSeparator.VB_Description = "Returns/Sets the value whether to display DropDown Separator."
Attribute DropDownSeparator.VB_ProcData.VB_Invoke_Property = ";Appearance"

    DropDownSeparator = m_bDropDownSep

End Property

Public Property Let DropDownSeparator(ByVal New_Value As Boolean)

    m_bDropDownSep = New_Value
    RedrawButton
    PropertyChanged "DropDownSeparator"

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value to determine whether the button can respond to events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514

    Enabled = m_bEnabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    m_bEnabled = New_Enabled
    UserControl.Enabled = m_bEnabled
    RedrawButton
    PropertyChanged "Enabled"

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/sets the Font used to display text on the button."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512

    Set Font = mFont

End Property

Public Property Set Font(ByVal New_Font As StdFont)

    Set mFont = New_Font
    Refresh
    RedrawButton
    PropertyChanged "Font"
    mFont_FontChanged ""

End Property

Private Sub mFont_FontChanged(ByVal PropertyName As String)

    Set UserControl.Font = mFont
    Refresh
    RedrawButton
    PropertyChanged "Font"

End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the text color of the button caption."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513

    ForeColor = m_bColors.tForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

    m_bColors.tForeColor = New_ForeColor
    UserControl.ForeColor = m_bColors.tForeColor
    RedrawButton
    PropertyChanged "ForeColor"

End Property

Public Property Get ForeColorHover() As OLE_COLOR
Attribute ForeColorHover.VB_Description = "Returns/sets the text color of the button caption when Mouse is over the control."
Attribute ForeColorHover.VB_ProcData.VB_Invoke_Property = ";Appearance"

    ForeColorHover = m_bColors.tForeColorOver

End Property

Public Property Let ForeColorHover(ByVal New_ForeColorHover As OLE_COLOR)

    m_bColors.tForeColorOver = New_ForeColorHover
    UserControl.ForeColor = m_bColors.tForeColorOver
    RedrawButton
    PropertyChanged "ForeColorHover"

End Property

Public Property Get HandPointer() As Boolean
Attribute HandPointer.VB_Description = "Returns/sets a value to determine whether the control uses the system's hand pointer as its cursor."
Attribute HandPointer.VB_ProcData.VB_Invoke_Property = ";Misc"

    HandPointer = m_bHandPointer

End Property

Public Property Let HandPointer(ByVal New_HandPointer As Boolean)

    m_bHandPointer = New_HandPointer
    If m_bHandPointer Then
        UserControl.MousePointer = 0
    End If
    PropertyChanged "HandPointer"

End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle that uniquely identifies the control."
Attribute hWnd.VB_UserMemId = -515

' --Handle that uniquely identifies the control

    hWnd = UserControl.hWnd

End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in a button's picture to be transparent."
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"

    MaskColor = m_lMaskColor

End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)

    m_lMaskColor = New_MaskColor
    RedrawButton
    PropertyChanged "MaskColor"

End Property

Public Property Get Mode() As enumAquaButtonModes
Attribute Mode.VB_Description = "Returns/sets the type of control the button will observe."
Attribute Mode.VB_ProcData.VB_Invoke_Property = ";Behavior"

    Mode = m_ButtonMode

End Property

Public Property Let Mode(ByVal New_mode As enumAquaButtonModes)

    m_ButtonMode = New_mode
    If m_ButtonMode = ebmCommandButton Then
        m_Buttonstate = eStateNormal        'Force Normal State for command buttons
    End If
    RedrawButton
    PropertyChanged "Value"
    PropertyChanged "Mode"

End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon for the button."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_Icon As IPictureDisp)

    On Error Resume Next
        Set UserControl.MouseIcon = New_Icon
        If (New_Icon Is Nothing) Then
            UserControl.MousePointer = 0 ' vbDefault
        Else
            m_bHandPointer = False
            PropertyChanged "HandPointer"
            UserControl.MousePointer = 99 ' vbCustom
        End If
        PropertyChanged "MouseIcon"

End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when cursor over the button."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_Cursor As MousePointerConstants)

    UserControl.MousePointer = New_Cursor
    PropertyChanged "MousePointer"

End Property

Public Property Get PictureNormal() As StdPicture
Attribute PictureNormal.VB_Description = "Returns/sets the picture displayed on a normal state button."
Attribute PictureNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"

    Set PictureNormal = m_Picture

End Property

Public Property Set PictureNormal(ByVal New_Picture As StdPicture)

    Set m_Picture = New_Picture
    If Not New_Picture Is Nothing Then
        RedrawButton
        PropertyChanged "PictureNormal"
    Else
        UserControl_Resize
        Set m_PictureHot = Nothing
        Set m_PictureDown = Nothing
        PropertyChanged "PictureHot"
        PropertyChanged "PictureDown"
    End If

End Property

Public Property Get PictureHot() As StdPicture
Attribute PictureHot.VB_Description = "Returns/sets the picture displayed when the cursor is over the control."
Attribute PictureHot.VB_ProcData.VB_Invoke_Property = ";Appearance"

    Set PictureHot = m_PictureHot

End Property

Public Property Set PictureHot(ByVal New_Hot As StdPicture)

    If m_Picture Is Nothing Then
        Set m_Picture = New_Hot
        PropertyChanged "PictureNormal"
        Exit Property
    End If

    Set m_PictureHot = New_Hot
    PropertyChanged "PictureHot"
    RedrawButton

End Property

Public Property Get PictureDown() As StdPicture
Attribute PictureDown.VB_Description = "Returns/sets the picture displayed when the control is pressed down or in checked state."
Attribute PictureDown.VB_ProcData.VB_Invoke_Property = ";Appearance"

    Set PictureDown = m_PictureDown

End Property

Public Property Set PictureDown(ByVal New_Down As StdPicture)

    If m_Picture Is Nothing Then
        Set m_Picture = New_Down
        PropertyChanged "PictureNormal"
        Exit Property
    End If

    Set m_PictureDown = New_Down
    PropertyChanged "PictureDown"
    RedrawButton

End Property

Public Property Get PictureShadow() As Boolean
Attribute PictureShadow.VB_Description = "Returns/Sets a value to determine whether to display Picture Shadow"
Attribute PictureShadow.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PictureShadow = m_PictureShadow

End Property

Public Property Let PictureShadow(ByVal New_Shadow As Boolean)

    m_PictureShadow = New_Shadow
    RedrawButton
    PropertyChanged "PictureShadow"

End Property

Public Property Get PictureAlign() As enumAquaPictureAlign
Attribute PictureAlign.VB_Description = "Returns/sets a value to determine where to draw the picture in the button."
Attribute PictureAlign.VB_ProcData.VB_Invoke_Property = ";Position"

    PictureAlign = m_PictureAlign

End Property

Public Property Let PictureAlign(ByVal New_PictureAlign As enumAquaPictureAlign)

    m_PictureAlign = New_PictureAlign
    If Not m_Picture Is Nothing Then
        RedrawButton
    End If
    PropertyChanged "PictureAlign"

End Property

Public Property Get PictureEffectOnOver() As enumAquaPicEffect
Attribute PictureEffectOnOver.VB_Description = "Returns/Sets the Picture Effects to be applied when the mouseis over the control."
Attribute PictureEffectOnOver.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PictureEffectOnOver = m_PicEffectonOver

End Property

Public Property Let PictureEffectOnOver(ByVal New_Effect As enumAquaPicEffect)

    m_PicEffectonOver = New_Effect
    RedrawButton
    PropertyChanged "PictureEffectOnOver"

End Property

Public Property Get PictureEffectOnDown() As enumAquaPicEffect
Attribute PictureEffectOnDown.VB_Description = "Returns/Sets the Picture Effects to be applied when the Button is pressed down."
Attribute PictureEffectOnDown.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PictureEffectOnDown = m_PicEffectonDown

End Property

Public Property Let PictureEffectOnDown(ByVal New_Effect As enumAquaPicEffect)

    m_PicEffectonDown = New_Effect
    RedrawButton
    PropertyChanged "PictureEffectOnDown"

End Property

Public Property Get PicturePushOnHover() As Boolean
Attribute PicturePushOnHover.VB_Description = "Returns/Sets a value to determine whether to Push picture when Mouse is over the control."
Attribute PicturePushOnHover.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PicturePushOnHover = m_bPicPushOnHover

End Property

Public Property Let PicturePushOnHover(ByVal Value As Boolean)

    m_bPicPushOnHover = Value
    RedrawButton
    PropertyChanged "PicturePushOnHover"

End Property

Public Property Get PictureOpacity() As Byte
Attribute PictureOpacity.VB_Description = "Returns/Sets a byte value to control the Opacity of the Picture."
Attribute PictureOpacity.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PictureOpacity = m_PictureOpacity

End Property

Public Property Let PictureOpacity(ByVal New_Opacity As Byte)

    m_PictureOpacity = New_Opacity
    RedrawButton
    PropertyChanged "PictureOpacity"

End Property

Public Property Get PictureOpacityOnOver() As Byte
Attribute PictureOpacityOnOver.VB_Description = "Returns/Sets a byte value to control the Opacity of the Picture when Mouse is over the button."
Attribute PictureOpacityOnOver.VB_ProcData.VB_Invoke_Property = ";Appearance"

    PictureOpacityOnOver = m_PicOpacityOnOver

End Property

Public Property Let PictureOpacityOnOver(ByVal New_Opacity As Byte)

    m_PicOpacityOnOver = New_Opacity
    RedrawButton
    PropertyChanged "PictureOpacityOnOver"

End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Returns/Sets a value to determine whether to display text in RTL mode."
Attribute RightToLeft.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute RightToLeft.VB_UserMemId = -611

    RightToLeft = m_bRTL

End Property

Public Property Let RightToLeft(ByVal Value As Boolean)

    m_bttRTL = Value
    m_bRTL = Value
    RedrawButton
    PropertyChanged "RightToLeft"

End Property

Public Property Get CaptionEffects() As enumAquaCaptionEffects
Attribute CaptionEffects.VB_Description = "Returns/Sets the Special Effects apply to the caption."
Attribute CaptionEffects.VB_ProcData.VB_Invoke_Property = ";Appearance"

    CaptionEffects = m_CaptionEffects

End Property

Public Property Let CaptionEffects(ByVal New_Effects As enumAquaCaptionEffects)

    m_CaptionEffects = New_Effects
    RedrawButton
    PropertyChanged "CaptionEffects"

End Property

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_Description = "Returns/Sets a value to show Focusrect when the button has focus"
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = ";Appearance"

    ShowFocusRect = m_bShowFocus

End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)

    m_bShowFocus = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"

End Property

Public Property Get UseMaskColor() As Boolean

    UseMaskColor = m_bUseMaskColor

End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)

    m_bUseMaskColor = New_UseMaskColor
    If Not m_Picture Is Nothing Then
        RedrawButton
    End If
    PropertyChanged "UseMaskColor"

End Property

Public Property Get Value() As Boolean

    Value = m_bValue

End Property

Public Property Let Value(ByVal New_Value As Boolean)

    If m_ButtonMode <> ebmCommandButton Then
        m_bValue = New_Value
        'If Not m_bValue Then m_Buttonstate = eStateNormal
        If Not m_bValue Then
            m_Buttonstate = eStateNormal
        End If
        RedrawButton
        PropertyChanged "Value"
    Else
        m_Buttonstate = eStateNormal
        RedrawButton
    End If

End Property

Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Returns/Sets the text displayed when mouse is paused over the control."
Attribute ToolTip.VB_ProcData.VB_Invoke_Property = ";Appearance"

    ToolTip = m_sTooltipText

End Property

Public Property Let ToolTip(ByVal New_Tooltip As String)

    m_sTooltipText = New_Tooltip
    RedrawButton
    PropertyChanged "ToolTip"

End Property

Public Property Get TooltipTitle() As String

    TooltipTitle = m_sTooltiptitle

End Property

Public Property Let TooltipTitle(ByVal New_title As String)

    m_sTooltiptitle = New_title
    PropertyChanged "TooltipTitle"

End Property

Public Property Get TooltipBackColor() As OLE_COLOR

    TooltipBackColor = m_lttBackColor
    
End Property

Public Property Let TooltipBackColor(ByVal New_Color As OLE_COLOR)

    m_lttBackColor = New_Color
    RedrawButton
    PropertyChanged "TooltipBackcolor"
    
End Property

Public Property Let ToolTipIcon(lTooltipIcon As enumAquaIconType)

    m_lToolTipIcon = lTooltipIcon
    PropertyChanged "TooltipIcon"

End Property

Public Property Get ToolTipIcon() As enumAquaIconType

    ToolTipIcon = m_lToolTipIcon

End Property

Public Property Get ToolTipType() As enumAquaTooltipStyle

    ToolTipType = m_lTooltipType

End Property

Public Property Let ToolTipType(ByVal lNewTTType As enumAquaTooltipStyle)

    m_lTooltipType = lNewTTType
    PropertyChanged "ToolTipType"

End Property

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Stop subclassing the passed window handle

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

    Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
    Debug.Assert Subclass_AddrFunc

End Function

Private Function Subclass_Index(ByVal lHwnd As Long, Optional ByVal bAdd As Boolean) As Long

    For Subclass_Index = UBound(SubclassData) To 0 Step -1
        If SubclassData(Subclass_Index).hWnd = lHwnd Then
            'If Not bAdd Then Exit Function':(?-> replaced by:
            If Not bAdd Then
                Exit Function
            End If

        ElseIf SubclassData(Subclass_Index).hWnd = 0 Then
            'If bAdd Then Exit Function':(?-> replaced by:
            If bAdd Then
                Exit Function
            End If
        End If
    Next 'Subclass_Index

    'If Not bAdd Then Debug.Assert False':(?-> replaced by:
    If Not bAdd Then
        Debug.Assert False
    End If

End Function

Private Function Subclass_InIDE() As Boolean

    Debug.Assert Subclass_SetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Initialize(ByVal lHwnd As Long) As Long

Const CODE_LEN                  As Long = 200
Const GMEM_FIXED                As Long = 0
Const PATCH_01                  As Long = 18
Const PATCH_02                  As Long = 68
Const PATCH_03                  As Long = 78
Const PATCH_06                  As Long = 116
Const PATCH_07                  As Long = 121
Const PATCH_0A                  As Long = 186
Const FUNC_CWP                  As String = "CallWindowProcA"
Const FUNC_EBM                  As String = "EbMode"
Const FUNC_SWL                  As String = "SetWindowLongA"
Const MOD_USER                  As String = "User32"
Const MOD_VBA5                  As String = "vba5"
Const MOD_VBA6                  As String = "vba6"

Static bytBuffer(1 To CODE_LEN) As Byte
Static lngCWP                   As Long
Static lngEbMode                As Long
Static lngSWL                   As Long

Dim lngCount                    As Long
Dim lngIndex                    As Long
Dim strHex                      As String

    If bytBuffer(1) Then
        lngIndex = Subclass_Index(lHwnd, True)

        If lngIndex = -1 Then
            lngIndex = UBound(SubclassData()) + 1

            ReDim Preserve SubclassData(lngIndex) As SubClassDatatype
        End If

        Subclass_Initialize = lngIndex

    Else
        strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

        For lngCount = 1 To CODE_LEN
            bytBuffer(lngCount) = Val("&H" & Left$(strHex, 2))
            strHex = Mid$(strHex, 3)
        Next 'lngCount

        If Subclass_InIDE Then
            bytBuffer(16) = &H90
            bytBuffer(17) = &H90
            lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)

            'If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)':(?-> replaced by:
            If lngEbMode = 0 Then
                lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If

        lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
        lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)

        ReDim SubclassData(0) As SubClassDatatype
    End If

    With SubclassData(lngIndex)
        .hWnd = lHwnd
        .nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSclass)

        CopyMemory ByVal .nAddrSclass, bytBuffer(1), CODE_LEN
        Subclass_PatchRel .nAddrSclass, PATCH_01, lngEbMode
        Subclass_PatchVal .nAddrSclass, PATCH_02, .nAddrOrig
        Subclass_PatchRel .nAddrSclass, PATCH_03, lngSWL
        Subclass_PatchVal .nAddrSclass, PATCH_06, .nAddrOrig
        Subclass_PatchRel .nAddrSclass, PATCH_07, lngCWP
        Subclass_PatchVal .nAddrSclass, PATCH_0A, ObjPtr(Me)
    End With

End Function

Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean

    Subclass_SetTrue = True
    bValue = True

End Function

Private Sub Subclass_AddMsg(ByVal lHwnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

    With SubclassData(Subclass_Index(lHwnd))
        'If When And MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)':(?-> replaced by:
        If When And MSG_BEFORE Then
            Subclass_DoAddMsg uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass
        End If
        'If When And MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)':(?-> replaced by:
        If When And MSG_AFTER Then
            Subclass_DoAddMsg uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass
        End If
    End With

End Sub

Private Sub Subclass_DelMsg(ByVal lHwnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

    With SubclassData(Subclass_Index(lHwnd))
        'If When And MSG_BEFORE Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)':(?-> replaced by:
        If When And MSG_BEFORE Then
            Subclass_DoDelMsg uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass
        End If
        'If When And MSG_AFTER Then Call Subclass_DoDelMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)':(?-> replaced by:
        If When And MSG_AFTER Then
            Subclass_DoDelMsg uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass
        End If
    End With

End Sub

Private Sub Subclass_DoAddMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Const PATCH_04 As Long = 88
Const PATCH_08 As Long = 132

Dim lngEntry   As Long

    ReDim lngOffset(1) As Long

    If uMsg = ALL_MESSAGES Then
        nMsgCount = ALL_MESSAGES

    Else
        For lngEntry = 1 To nMsgCount - 1
            If aMsgTabel(lngEntry) = 0 Then
                aMsgTabel(lngEntry) = uMsg

                GoTo ExitSub

            ElseIf aMsgTabel(lngEntry) = uMsg Then
                GoTo ExitSub
            End If
        Next 'lngEntry

        nMsgCount = nMsgCount + 1

        ReDim Preserve aMsgTabel(1 To nMsgCount) As Long

        aMsgTabel(nMsgCount) = uMsg
    End If

    If When = MSG_BEFORE Then
        lngOffset(0) = PATCH_04
        lngOffset(1) = PATCH_05

    Else
        lngOffset(0) = PATCH_08
        lngOffset(1) = PATCH_09
    End If

    'If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))':(?-> replaced by:
    If uMsg <> ALL_MESSAGES Then
        Subclass_PatchVal nAddr, lngOffset(0), VarPtr(aMsgTabel(1))
    End If

    Subclass_PatchVal nAddr, lngOffset(1), nMsgCount

ExitSub:
    Erase lngOffset

End Sub

Private Sub Subclass_DoDelMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Dim lngEntry As Long

    If uMsg = ALL_MESSAGES Then
        nMsgCount = 0

        If When = MSG_BEFORE Then
            lngEntry = PATCH_05

        Else
            lngEntry = PATCH_09
        End If

        Subclass_PatchVal nAddr, lngEntry, 0

    Else
        For lngEntry = 1 To nMsgCount - 1
            If aMsgTabel(lngEntry) = uMsg Then
                aMsgTabel(lngEntry) = 0
                Exit For
            End If
        Next 'lngEntry
    End If

End Sub

Private Sub Subclass_PatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

    CopyMemory ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4

End Sub

Private Sub Subclass_PatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

    CopyMemory ByVal nAddr + nOffset, nValue, 4

End Sub

Private Sub Subclass_Stop(ByVal lHwnd As Long)

    With SubclassData(Subclass_Index(lHwnd))
        SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig

        Subclass_PatchVal .nAddrSclass, PATCH_05, 0
        Subclass_PatchVal .nAddrSclass, PATCH_09, 0

        GlobalFree .nAddrSclass
        .hWnd = 0
        .nMsgCountA = 0
        .nMsgCountB = 0
        Erase .aMsgTabelA, .aMsgTabelB
    End With

End Sub

Private Sub Subclass_Terminate()

Dim lngCount As Long

    For lngCount = UBound(SubclassData) To 0 Step -1
        'If SubclassData(lngCount).hWnd Then Call Subclass_Stop(SubclassData(lngCount).hWnd)':(?-> replaced by:
        If SubclassData(lngCount).hWnd Then
            Subclass_Stop SubclassData(lngCount).hWnd
        End If
    Next 'lngCount

End Sub

'-------------------------x----------------------x--------------------x---------

