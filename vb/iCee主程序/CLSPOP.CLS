VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSPOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'文本框冒泡提示
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTM_TRACKACTIVATE = WM_USER + 17
Private Const TTM_TRACKPOSITION = WM_USER + 18
Private Const WS_POPUP = &H80000000

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttStyleEnum
    TTStandard
    TTBalloon
End Enum

Private mvarBackColor As Long
Private mvarTitle As String
Private mvarForeColor As Long
Private mvarIcon As ttIconType
Private mvarCentered As Boolean
Private mvarStyle As ttStyleEnum
Private mvarTipText As String
Private mvarVisibleTime As Long
Private mvarDelayTime As Long
Private mvarPopupOnDemand As Boolean

Private m_lTTHwnd As Long
Private m_lParentHwnd As Long
Private Ti As TOOLINFO

Private Sub Class_Initialize()
    mvarDelayTime = 500
    mvarVisibleTime = 5000
    mvarPopupOnDemand = False
End Sub
Private Sub Class_Terminate()
    Destroy
End Sub
Public Property Get VisibleTime() As Long
    VisibleTime = mvarVisibleTime
End Property
Public Property Let VisibleTime(ByVal lData As Long)
    mvarVisibleTime = lData
End Property
Public Property Get DelayTime() As Long
    DelayTime = mvarDelayTime
End Property
Public Property Let DelayTime(ByVal lData As Long)
    mvarDelayTime = lData
End Property
Public Property Let Icon(ByVal vData As ttIconType)
    mvarIcon = vData
    If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If
End Property
Public Property Get Icon() As ttIconType
    Icon = mvarIcon
End Property
Public Property Let ForeColor(ByVal vData As Long)
    mvarForeColor = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If
End Property
Public Property Get ForeColor() As Long
    ForeColor = mvarForeColor
End Property
Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
    If m_lTTHwnd <> 0 And mvarTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If
End Property
Public Property Get Title() As String
    Title = Ti.lpStr
End Property
Public Property Let PopupOnDemand(ByVal vData As Boolean)
    mvarPopupOnDemand = vData
End Property
Public Property Get PopupOnDemand() As Boolean
    PopupOnDemand = mvarPopupOnDemand
End Property
Public Property Let BackColor(ByVal vData As Long)
    mvarBackColor = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If
End Property
Public Property Get BackColor() As Long
    BackColor = mvarBackColor
End Property
Public Property Let TipText(ByVal vData As String)
    mvarTipText = vData
    Ti.lpStr = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, Ti
    End If
End Property
Public Property Get TipText() As String
    TipText = mvarTipText
End Property
Public Property Let Style(ByVal vData As ttStyleEnum)
    mvarStyle = vData
End Property
Public Property Get Style() As ttStyleEnum
    Style = mvarStyle
End Property
Public Property Let Centered(ByVal vData As Boolean)
    mvarCentered = vData
End Property
Public Property Get Centered() As Boolean
    Centered = mvarCentered
End Property
Public Sub Show(TXT As TextBox, Optional X As Long = 0, Optional Y As Long = 0)
    Dim PT As POINTAPI
    Dim ptTip As Long
    Dim ret As Long

    With PT
        .X = X
        .Y = Y
    End With
    ret = ClientToScreen(TXT.hwnd, PT)
    ptTip = PT.Y * &H10000
    ptTip = ptTip + PT.X
    ret = SendMessage(m_lTTHwnd, TTM_TRACKPOSITION, 0, ByVal ptTip)
    ret = SendMessage(m_lTTHwnd, TTM_TRACKACTIVATE, True, Ti)
End Sub
Public Function CreateToolTip(ByVal ParentHwnd As Long) As Boolean
    Dim lWinStyle As Long
    If m_lTTHwnd <> 0 Then
        DestroyWindow m_lTTHwnd
    End If
    m_lParentHwnd = ParentHwnd
    ''create baloon style if desired
    If mvarStyle = TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON

    m_lTTHwnd = CreateWindowEx(0&, TOOLTIPS_CLASSA, vbNullString, lWinStyle, 0&, 0&, 0&, 0&, m_lParentHwnd, 0&, 0&, 0&)
    With Ti
        If mvarCentered Then
            If mvarPopupOnDemand = False Then
                .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
            Else
                .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_CENTERTIP
            End If
        Else
            If mvarPopupOnDemand = False Then
                .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
            Else
                .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_TRANSPARENT
            End If
        End If

        'set the hwnd prop to our parent control's hwnd
        .hwnd = m_lParentHwnd
        .lId = m_lParentHwnd                              '0
        .hInstance = App.hInstance
        .lSize = Len(Ti)
    End With
    'add the tooltip structure
    SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, Ti

    'if we want a title or we want an icon
    If mvarTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETTITLE, CLng(mvarIcon), ByVal mvarTitle
    End If
    If mvarForeColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If
    If mvarBackColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If
    SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mvarVisibleTime
    SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, mvarDelayTime
End Function

Public Sub Destroy()
    If m_lTTHwnd <> 0 Then
        DestroyWindow m_lTTHwnd
    End If
End Sub

