VERSION 5.00
Begin VB.UserControl CandyButton 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "CandyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
'download by http://www.codefans.net
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Event Status(ByVal sStatus As String)

Private Const WM_MOUSEMOVE    As Long = &H200
Private Const WM_MOUSELEAVE   As Long = &H2A3
Private Const WM_MOVING       As Long = &H216
Private Const WM_SIZING       As Long = &H214
Private Const WM_EXITSIZEMOVE As Long = &H232

Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                      As Long
  dwFlags                     As TRACKMOUSEEVENT_FLAGS
  hwndTrack                   As Long
  dwHoverTime                 As Long
End Type

Private bTrack                As Boolean
Private bTrackUser32          As Boolean
Private IsHover               As Boolean
Private bMoving               As Boolean

Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'-Candy Button declarations----------------------------------------------------------------------------
Public Enum eAlignment
    PIC_TOP
    PIC_BOTTOM
    PIC_LEFT
    PIC_RIGHT
End Enum

Public Enum eStyle
    XP_Button
    XP_ToolBarButton
    Crystal
    Mac
    Mac_Variation
    WMP
    Plastic
End Enum

Public Enum eColorScheme
    Custom
    Aqua
    WMP10
    DeepBlue
    DeepRed
    DeepGreen
    DeepYellow
End Enum

Public Enum eState
    eNormal
    ePressed
    eFocus
    eHover
    eChecked
End Enum

Private Type tCrystalParam
    Ref_MixColorFrom As Long
    Ref_Intensity As Long
    Ref_Left As Long
    Ref_Top As Long
    Ref_Radius As Long
    Ref_Height As Long
    Ref_Width As Long
    RadialGXPercent As Long
    RadialGYPercent As Long
End Type

Private m_PictureAlignment As eAlignment
Private m_Style As eStyle
Private m_Checked As Boolean
Private m_hasFocus As Boolean
Private m_Caption As String
Private m_StdPicture As StdPicture
Private m_Font As StdFont
Private m_ColorButtonHover As OLE_COLOR
Private m_ColorButtonUp As OLE_COLOR
Private m_ColorButtonDown As OLE_COLOR
Private m_ColorBright As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_DisplayHand As Boolean
Private CornerRadius As Long
Private m_BorderBrightness As Long
Private m_ColorScheme As eColorScheme

Private Const m_def_ForeColor = vbBlack
Private Const m_def_PictureAlignment = 0

Private Const RGN_XOR = 3

Private Const MK_LBUTTON = &H1

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Property Let DisplayHand(newValue As Boolean)
    m_DisplayHand = newValue
End Property

Public Property Get DisplayHand() As Boolean
    DisplayHand = m_DisplayHand
End Property

Public Property Let ColorScheme(newValue As eColorScheme)
    Select Case newValue
        Case Aqua
            ColorButtonUp = &HD06720
            ColorButtonHover = &HE99950
            ColorButtonDown = &HA06710
            ColorBright = &HFFEDB0
        Case WMP10
            ColorButtonUp = &HD09060
            ColorButtonHover = &HE06000
            ColorButtonDown = &HA98050
            ColorBright = &HFFFAFA
        Case DeepBlue
            ColorButtonUp = &H800000
            ColorButtonHover = &HA00000
            ColorButtonDown = &HF00000
            ColorBright = &HFF0000
        Case DeepRed
            ColorButtonUp = &H80&
            ColorButtonHover = &HA0&
            ColorButtonDown = &HF0&
            ColorBright = &HFF&
        Case DeepGreen
            ColorButtonUp = &H8000&
            ColorButtonHover = &HA000&
            ColorButtonDown = &HC000&
            ColorBright = &HFF00&
        Case DeepYellow
            ColorButtonUp = &H8080&
            ColorButtonHover = &HA0A0&
            ColorButtonDown = &HC0C0&
            ColorBright = &HFFFF&
    End Select
    m_ColorScheme = newValue
    PropertyChanged "m_ColorScheme"
    DrawButton (eNormal)
End Property

Public Property Get ColorScheme() As eColorScheme
    ColorScheme = m_ColorScheme
End Property

Public Property Let BorderBrightness(newValue As Long)
    m_BorderBrightness = SetBound(newValue, -100, 100)
    PropertyChanged "m_BorderBrightness"
    DrawButton (eNormal)
End Property

Public Property Get BorderBrightness() As Long
    BorderBrightness = m_BorderBrightness
End Property

Public Property Let ColorBright(newValue As OLE_COLOR)
    m_ColorBright = newValue
    If m_ColorScheme <> Custom Then m_ColorScheme = Custom:  PropertyChanged "m_ColorScheme"
    PropertyChanged "m_ColorBright"
    DrawButton (eNormal)
End Property

Public Property Get ColorBright() As OLE_COLOR
    ColorBright = m_ColorBright
End Property

Public Property Let ColorButtonDown(newValue As OLE_COLOR)
    m_ColorButtonDown = newValue
    If m_ColorScheme <> Custom Then m_ColorScheme = Custom:  PropertyChanged "m_ColorScheme"
    PropertyChanged "m_ColorButtonDown"
    DrawButton (eNormal)
End Property

Public Property Get ColorButtonDown() As OLE_COLOR
    ColorButtonDown = m_ColorButtonDown
End Property

Public Property Let ColorButtonUp(newValue As OLE_COLOR)
    m_ColorButtonUp = newValue
    If m_ColorScheme <> Custom Then m_ColorScheme = Custom:  PropertyChanged "m_ColorScheme"
    PropertyChanged "m_ColorButtonUp"
    DrawButton (eNormal)
End Property

Public Property Get ColorButtonUp() As OLE_COLOR
    ColorButtonUp = m_ColorButtonUp
End Property

Public Property Let ColorButtonHover(newValue As OLE_COLOR)
    m_ColorButtonHover = newValue
    If m_ColorScheme <> Custom Then m_ColorScheme = Custom:  PropertyChanged "m_ColorScheme"
    PropertyChanged "m_ColorButtonHover"
    DrawButton (eNormal)
End Property

Public Property Get ColorButtonHover() As OLE_COLOR
    ColorButtonHover = m_ColorButtonHover
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
     m_ForeColor = NewForeColor
     Picture1.ForeColor = m_ForeColor
     PropertyChanged "ForeColor"
     DrawButton (eNormal)
End Property

Public Property Get ForeColor() As OLE_COLOR
     ForeColor = m_ForeColor
End Property

Public Property Set Picture(Value As StdPicture)
    Set m_StdPicture = Value
    PropertyChanged "Picture"
    DrawButton (eNormal)
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_StdPicture
End Property

Public Property Let Checked(Value As Boolean)
    m_Checked = Value
    If Value Then
        DrawButton (eChecked)
    Else
        If IsHover Then
            DrawButton (eHover)
        Else
            DrawButton (eNormal)
        End If
    End If
    PropertyChanged "Checked"
End Property

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let Style(eVal As eStyle)
    If eVal <> m_Style Then
        m_Style = eVal
        PropertyChanged "Style"
        Init_Style
    End If
End Property

Public Property Get Style() As eStyle
    Style = m_Style
End Property

Public Property Let PictureAlignment(eVal As eAlignment)
    If eVal <> m_PictureAlignment Then
        m_PictureAlignment = eVal
        PropertyChanged "PictureAlignment"
        DrawButton (eNormal)
    End If
End Property

Public Property Get PictureAlignment() As eAlignment
    PictureAlignment = m_PictureAlignment
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawButton (eNormal)
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Set Font(ByVal NewFont As StdFont)
     Set Picture1.Font = NewFont
     PropertyChanged "Font"
     DrawButton (eNormal)
End Property

Public Property Get Font() As StdFont
     Set Font = Picture1.Font
End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Caption = Extender.Name
End Sub

Private Sub UserControl_Initialize()
    m_Style = Style
End Sub

Private Sub UserControl_InitProperties()
    If Not Ambient.UserMode Then
        m_ColorButtonHover = &HFFC090
        m_ColorButtonUp = &HE99950
        m_ColorBright = &HFFEDB0
        m_ColorButtonDown = &HE99950
        m_Caption = UserControl.Name
        Picture1.Picture = LoadPicture("")
    End If
    m_Caption = Extender.Name
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then UserControl_MouseDown 1, 0, 0, 0
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then UserControl_MouseUp 1, 0, 0, 0
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_hasFocus = True
    DrawButton (ePressed)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
    If Button = 1 And (x < 0 Or x > UserControl.ScaleWidth Or y < 0 Or y > UserControl.ScaleHeight) Then
        IsHover = False
        DrawButton (eNormal)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If m_Checked = False Then If IsHover Then DrawButton (eHover) Else If m_hasFocus Then DrawButton (eFocus)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_DblClick()
    DrawButton (ePressed)
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    m_hasFocus = True
    If m_Checked = False And Not IsHover Then DrawButton (eFocus)
End Sub

Private Sub UserControl_ExitFocus()
    m_hasFocus = False
    If m_Checked = False Then DrawButton (eNormal)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", Picture1.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Name)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Picture", m_StdPicture, Nothing)
    Call PropBag.WriteProperty("PictureAlignment", m_PictureAlignment, m_def_PictureAlignment)
    Call PropBag.WriteProperty("Style", m_Style, 0)
    Call PropBag.WriteProperty("Checked", m_Checked)
    Call PropBag.WriteProperty("ColorButtonHover", m_ColorButtonHover)
    Call PropBag.WriteProperty("ColorButtonUp", m_ColorButtonUp)
    Call PropBag.WriteProperty("ColorButtonDown", m_ColorButtonDown)
    Call PropBag.WriteProperty("BorderBrightness", m_BorderBrightness)
    Call PropBag.WriteProperty("ColorBright", m_ColorBright)
    Call PropBag.WriteProperty("DisplayHand", m_DisplayHand)
    Call PropBag.WriteProperty("ColorScheme", m_ColorScheme)
End Sub

Private Sub UserControl_Resize()
    Init_Style
End Sub

Private Sub UserControl_Show()
    DrawButton (eNormal)
End Sub

Private Sub Init_Style()
    Select Case m_Style
        Case Crystal, WMP, Mac_Variation
            CreateRoundedRegion UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, SetBound(UserControl.ScaleHeight \ 2, 0, UserControl.ScaleWidth \ 2)
        Case Mac
            CreateRoundedRegion UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, 12
        Case Plastic
            CreateRoundedRegion UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, SetBound(UserControl.ScaleHeight \ 3, 0, UserControl.ScaleWidth \ 3)
        Case Else
            Call SetWindowRgn(UserControl.hwnd, 0, True)
    End Select
    Picture1.Picture = LoadPicture("")
    Picture1.Width = UserControl.ScaleWidth
    Picture1.Height = UserControl.ScaleHeight
    DrawButton (eNormal)
End Sub

Private Sub DrawButton(Optional vState As eState = eNormal)
    If m_Checked Then vState = eChecked
    Select Case m_Style
        Case XP_Button
            DrawXPButton vState
        Case Crystal, Mac, WMP, Mac_Variation
            DrawCrystalButton vState
        Case Plastic
            DrawPlasticButton vState
        Case XP_ToolBarButton
            DrawXPToolbarButton vState
    End Select
    DrawIconWCaption vState
    Picture1.Picture = LoadPicture("")
End Sub

Public Sub DrawIconWCaption(vState As eState)
    Dim pW As Long, pH As Long, lW As Long, lH As Long, uW As Long, uH As Long
    Dim StartX As Long, StartY As Long
    
    Picture1.ForeColor = m_ForeColor
    
    If Not m_StdPicture Is Nothing Then
        pW = ScaleX(m_StdPicture.Width, vbHimetric, vbPixels)
        pH = ScaleY(m_StdPicture.Height, vbHimetric, vbPixels)
    End If
    
    If Len(m_Caption) <> 0 Then
        lW = Picture1.TextWidth(m_Caption)
        lH = Picture1.TextHeight(m_Caption)
    End If
    
    uW = UserControl.ScaleWidth
    uH = UserControl.ScaleHeight
    
    Select Case m_PictureAlignment
        Case Is = PIC_TOP
            StartX = ((uW - pW) \ 2) + 1
            StartY = (uH - (pH + lH)) \ 2
            Picture1.CurrentX = Abs(uW \ 2 - lW \ 2)
            Picture1.CurrentY = Abs(uH \ 2 + pH \ 2 - lH \ 2)
        Case Is = PIC_BOTTOM
            StartX = (uW - pW) \ 2
            StartY = (uH - (pH - lH)) \ 2
            Picture1.CurrentX = Abs(uW \ 2 - lW \ 2)
            Picture1.CurrentY = Abs(uH \ 2 - (pH + lH) \ 2)
        Case Is = PIC_LEFT
            If CornerRadius Then StartX = CornerRadius Else StartX = 8
            StartY = (uH - pH) \ 2
            Picture1.CurrentX = Abs(uW \ 2 - lW \ 2)
            Picture1.CurrentY = Abs(uH \ 2 - lH \ 2)
        Case Is = PIC_RIGHT
            If CornerRadius Then StartX = uW - CornerRadius - pW Else StartX = uW - 8 - pW
            StartY = (uH - pH) \ 2
            Picture1.CurrentX = Abs(uW \ 2 - lW \ 2)
            Picture1.CurrentY = Abs(uH \ 2 - lH \ 2)
    End Select
    If vState = ePressed Then
        StartX = StartX + 1: Picture1.CurrentX = Picture1.CurrentX + 1
        StartY = StartY + 1: Picture1.CurrentY = Picture1.CurrentY + 1
    End If
    Picture1.Print m_Caption
    If Not m_StdPicture Is Nothing Then m_StdPicture.Render Picture1.hdc, CLng(StartX), CLng(StartY), CLng(pW), CLng(pH), _
        0, m_StdPicture.Height, m_StdPicture.Width, -m_StdPicture.Height, ByVal 0&
        
    Set UserControl.Picture = Picture1.Image
End Sub

Private Function DrawXPToolbarButton(Optional vState As eState)
Dim i As Long
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim uH As Long, uW As Long
    uH = UserControl.ScaleHeight - 1
    uW = UserControl.ScaleWidth - 1
    On Error Resume Next
    Picture1.Line (0, 0)-(uW, uH), UserControl.Parent.BackColor, BF
    On Error GoTo 0
    If vState = ePressed Then
        r1 = 220: g1 = 218: b1 = 209
        r2 = 231: g2 = 230: b2 = 224
        For i = 0 To 3
            Picture1.Line (0, 1 + i)-(uW, 1 + i), RGB(r2 * (i / 3) + r1 - (r1 * (i / 3)), g2 * (i / 3) + g1 - (g1 * (i / 3)), b2 * (i / 3) + b1 - (b1 * (i / 3)))
        Next
        r1 = 231: g1 = 230: b1 = 224
        r2 = 225: g2 = 224: b2 = 216
        For i = 4 To uH - 4
            Picture1.Line (0, i)-(uW, i), RGB(r2 * (i / (uH - 6)) + r1 - (r1 * (i / (uH - 6))), g2 * (i / (uH - 6)) + g1 - (g1 * (i / (uH - 6))), b2 * (i / (uH - 6)) + b1 - (b1 * (i / (uH - 6))))
        Next
        r1 = 225: g1 = 224: b1 = 216
        r2 = 235: g2 = 234: b2 = 229
        For i = 0 To 3
            Picture1.Line (0, uH - 4 + i)-(uW, uH - 4 + i), RGB(r2 * (i / 3) + r1 - (r1 * (i / 3)), g2 * (i / 3) + g1 - (g1 * (i / 3)), b2 * (i / 3) + b1 - (b1 * (i / 3)))
        Next
        Picture1.PSet (1, 0), RGB(215, 215, 204): Picture1.PSet (0, 1), RGB(215, 215, 204)
        Picture1.Line (0, 2)-(2, 0), RGB(179, 179, 168) '7617536
        Picture1.Line (2, 0)-(uW - 2, 0), RGB(157, 157, 146)
        Picture1.PSet (uW - 1, 0), RGB(215, 215, 204): Picture1.PSet (uW, 1), RGB(215, 215, 204)
        Picture1.Line (uW - 2, 0)-(uW, 2), RGB(179, 179, 168) '7617536
        Picture1.Line (uW, 2)-(uW, uH - 2), RGB(157, 157, 146)
        Picture1.PSet (uW, uH - 1), RGB(215, 215, 204): Picture1.PSet (uW - 1, uH), RGB(215, 215, 204)
        Picture1.Line (uW, uH - 2)-(uW - 2, uH), RGB(179, 179, 168) ' 7617536
        Picture1.Line (uW - 2, uH)-(2, uH), RGB(157, 157, 146)
        Picture1.PSet (1, uH), RGB(215, 215, 204): Picture1.PSet (0, uH - 1), RGB(215, 215, 204)
        Picture1.Line (2, uH)-(0, uH - 2), RGB(179, 179, 168) '7617536
        Picture1.Line (0, uH - 2)-(0, 2), RGB(157, 157, 146)
    ElseIf vState = eHover Then
        r1 = 254: g1 = 254: b1 = 253
        r2 = 252: g2 = 252: b2 = 249
        For i = 0 To 3
            Picture1.Line (0, 1 + i)-(uW, 1 + i), RGB(r2 * (i / 3) + r1 - (r1 * (i / 3)), g2 * (i / 3) + g1 - (g1 * (i / 3)), b2 * (i / 3) + b1 - (b1 * (i / 3)))
        Next
        r1 = 252: g1 = 252: b1 = 249
        r2 = 238: g2 = 237: b2 = 229
        For i = 4 To uH - 4
            Picture1.Line (0, i)-(uW, i), RGB(r2 * (i / (uH - 6)) + r1 - (r1 * (i / (uH - 6))), g2 * (i / (uH - 6)) + g1 - (g1 * (i / (uH - 6))), b2 * (i / (uH - 6)) + b1 - (b1 * (i / (uH - 6))))
        Next
        r1 = 238: g1 = 237: b1 = 229
        r2 = 215: g2 = 210: b2 = 198
        For i = 0 To 3
            Picture1.Line (0, uH - 4 + i)-(uW, uH - 4 + i), RGB(r2 * (i / 3) + r1 - (r1 * (i / 3)), g2 * (i / 3) + g1 - (g1 * (i / 3)), b2 * (i / 3) + b1 - (b1 * (i / 3)))
        Next
        
        Picture1.PSet (1, 0), RGB(232, 232, 221): Picture1.PSet (0, 1), RGB(232, 232, 221)
        Picture1.Line (0, 2)-(2, 0), RGB(216, 216, 205) '7617536
        Picture1.Line (2, 0)-(uW - 2, 0), RGB(206, 206, 195)
        Picture1.PSet (uW - 1, 0), RGB(232, 232, 221): Picture1.PSet (uW, 1), RGB(232, 232, 221)
        Picture1.Line (uW - 2, 0)-(uW, 2), RGB(216, 216, 205) '7617536
        Picture1.Line (uW, 2)-(uW, uH - 2), RGB(206, 206, 195)
        Picture1.PSet (uW, uH - 1), RGB(232, 232, 221): Picture1.PSet (uW - 1, uH), RGB(232, 232, 221)
        Picture1.Line (uW, uH - 2)-(uW - 2, uH), RGB(216, 216, 205) ' 7617536
        Picture1.Line (uW - 2, uH)-(2, uH), RGB(206, 206, 195)
        Picture1.PSet (1, uH), RGB(232, 232, 221): Picture1.PSet (0, uH - 1), RGB(232, 232, 221)
        Picture1.Line (2, uH)-(0, uH - 2), RGB(216, 216, 205) '7617536
        Picture1.Line (0, uH - 2)-(0, 2), RGB(206, 206, 195)
    ElseIf vState = eChecked Then
        Picture1.Line (1, 1)-(uW - 1, uH - 1), vbWhite, BF
        Picture1.PSet (1, 0), RGB(203, 213, 214): Picture1.PSet (0, 1), RGB(203, 213, 214)
        Picture1.Line (0, 2)-(2, 0), RGB(152, 175, 190) '7617536
        Picture1.Line (2, 0)-(uW - 2, 0), RGB(122, 152, 175)
        Picture1.PSet (uW - 1, 0), RGB(203, 213, 214): Picture1.PSet (uW, 1), RGB(203, 213, 214)
        Picture1.Line (uW - 2, 0)-(uW, 2), RGB(152, 175, 190) '7617536
        Picture1.Line (uW, 2)-(uW, uH - 2), RGB(122, 152, 175)
        Picture1.PSet (uW, uH - 1), RGB(203, 213, 214): Picture1.PSet (uW - 1, uH), RGB(203, 213, 214)
        Picture1.Line (uW, uH - 2)-(uW - 2, uH), RGB(152, 175, 190) ' 7617536
        Picture1.Line (uW - 2, uH)-(2, uH), RGB(122, 152, 175)
        Picture1.PSet (1, uH), RGB(203, 213, 214): Picture1.PSet (0, uH - 1), RGB(203, 213, 214)
        Picture1.Line (2, uH)-(0, uH - 2), RGB(152, 175, 190) '7617536
        Picture1.Line (0, uH - 2)-(0, 2), RGB(122, 152, 175)
    End If
End Function

Private Function DrawXPButton(Optional vState As eState)
Dim i As Long
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim uH As Long, uW As Long
    uH = UserControl.ScaleHeight - 1
    uW = UserControl.ScaleWidth - 1
    On Error Resume Next
    Picture1.Line (0, 0)-(uW, uH), UserControl.Parent.BackColor, BF
    On Error GoTo 0
    If vState = ePressed Then
        r1 = 209: g1 = 204: b1 = 193
        r2 = 229: g2 = 228: b2 = 221
        For i = 0 To 3
            Picture1.Line (0, 1 + i)-(uW, 1 + i), RGB(r2 * (i / 3) + r1 - (r1 * (i / 3)), g2 * (i / 3) + g1 - (g1 * (i / 3)), b2 * (i / 3) + b1 - (b1 * (i / 3)))
        Next
        r1 = 229: g1 = 228: b1 = 221
        r2 = 226: g2 = 226: b2 = 218
        For i = 4 To uH - 4
            Picture1.Line (0, i)-(uW, i), RGB(r2 * (i / (uH - 6)) + r1 - (r1 * (i / (uH - 6))), g2 * (i / (uH - 6)) + g1 - (g1 * (i / (uH - 6))), b2 * (i / (uH - 6)) + b1 - (b1 * (i / (uH - 6))))
        Next
        r1 = 226: g1 = 226: b1 = 218
        r2 = 242: g2 = 241: b2 = 238
        For i = 0 To 4
            Picture1.Line (0, uH - 4 + i)-(uW, uH - 4 + i), RGB(r2 * (i / 3) + r1 - (r1 * (i / 3)), g2 * (i / 3) + g1 - (g1 * (i / 3)), b2 * (i / 3) + b1 - (b1 * (i / 3)))
        Next
    Else
        r1 = 236: g1 = 235: b1 = 230
        r2 = 214: g2 = 208: b2 = 197
        For i = 0 To uH - 3
            Picture1.Line (1, i)-(uW, i), RGB(r1 * (i / (uH - 3)) + 255 - (255 * (i / (uH - 3))), g1 * (i / (uH - 3)) + 255 - (255 * (i / (uH - 3))), b1 * (i / (uH - 3)) + 255 - (255 * (i / (uH - 3))))
        Next
    
        For i = 0 To 3
            Picture1.Line (0, uH - 4 + i)-(uW, uH - 4 + i), RGB(r2 * (i / 3) + r1 - (r1 * (i / 3)), g2 * (i / 3) + g1 - (g1 * (i / 3)), b2 * (i / 3) + b1 - (b1 * (i / 3)))
        Next
    End If
    
    Select Case vState
        Case Is = eFocus
            Picture1.Line (0, 1)-(uW, 1), RGB(206, 231, 255)
            Picture1.Line (0, 2)-(uW, 2), RGB(188, 212, 246)
            r1 = 188: g1 = 212: b1 = 246
            r2 = 137: g2 = 173: b2 = 228
            For i = 3 To uH - 3
                Picture1.Line (0, i)-(3, i), RGB(r2 * (i / uH) + r1 - (r1 * (i / uH)), g2 * (i / uH) + g1 - (g1 * (i / uH)), b2 * (i / uH) + b1 - (b1 * (i / uH)))
                Picture1.Line (uW - 2, i)-(uW, i), RGB(r2 * (i / uH) + r1 - (r1 * (i / uH)), g2 * (i / uH) + g1 - (g1 * (i / uH)), b2 * (i / uH) + b1 - (b1 * (i / uH)))
            Next
            Picture1.Line (0, uH - 2)-(uW, uH - 2), RGB(137, 173, 228)
            Picture1.Line (0, uH - 1)-(uW, uH - 1), RGB(105, 130, 238)
        Case Is = eHover
            Picture1.Line (0, 1)-(uW, 1), RGB(255, 240, 202)
            Picture1.Line (0, 2)-(uW, 2), RGB(253, 216, 137)
            r1 = 253: g1 = 216: b1 = 137
            r2 = 248: g2 = 178: b2 = 48
            For i = 3 To uH - 3
                Picture1.Line (0, i)-(3, i), RGB(r2 * (i / uH) + r1 - (r1 * (i / uH)), g2 * (i / uH) + g1 - (g1 * (i / uH)), b2 * (i / uH) + b1 - (b1 * (i / uH)))
                Picture1.Line (uW - 2, i)-(uW, i), RGB(r2 * (i / uH) + r1 - (r1 * (i / uH)), g2 * (i / uH) + g1 - (g1 * (i / uH)), b2 * (i / uH) + b1 - (b1 * (i / uH)))
            Next
            Picture1.Line (0, uH - 2)-(uW, uH - 2), RGB(248, 178, 48)
            Picture1.Line (0, uH - 1)-(uW, uH - 1), RGB(229, 151, 0)
    End Select
    
    Picture1.PSet (0, 1), RGB(122, 149, 168): Picture1.PSet (1, 0), RGB(122, 149, 168)
    Picture1.Line (0, 2)-(2, 0), RGB(37, 87, 131) '7617536
    Picture1.Line (2, 0)-(uW - 2, 0), 7617536
    Picture1.PSet (uW - 1, 0), RGB(122, 149, 168): Picture1.PSet (uW, 1), RGB(122, 149, 168)
    Picture1.Line (uW - 2, 0)-(uW, 2), RGB(37, 87, 131)  '7617536
    Picture1.Line (uW, 2)-(uW, uH - 2), 7617536
    Picture1.PSet (uW, uH - 1), RGB(122, 149, 168): Picture1.PSet (uW - 1, uH), RGB(122, 149, 168)
    Picture1.Line (uW, uH - 2)-(uW - 2, uH), RGB(37, 87, 131) ' 7617536
    Picture1.Line (uW - 2, uH)-(2, uH), 7617536
    Picture1.PSet (1, uH), RGB(122, 149, 168): Picture1.PSet (0, uH - 1), RGB(122, 149, 168)
    Picture1.Line (2, uH)-(0, uH - 2), RGB(37, 87, 131)  '7617536
    Picture1.Line (0, uH - 2)-(0, 2), 7617536
End Function

Private Function DrawCrystalButton(vState As eState)
    Dim CrystalParam As tCrystalParam
    If m_Style = Mac Then 'Mac
        CrystalParam.Ref_MixColorFrom = 0 '20
        CrystalParam.Ref_Intensity = 70 '50
        CrystalParam.Ref_Left = (CornerRadius \ 3)
        CrystalParam.Ref_Top = 0
        CrystalParam.Ref_Height = 10 'CornerRadius - 2
        CrystalParam.Ref_Width = Picture1.ScaleWidth + 2 * CornerRadius
        CrystalParam.Ref_Radius = 5 'CornerRadius \ 2
        CrystalParam.RadialGXPercent = 100
        CrystalParam.RadialGYPercent = 100 - (7 * 100 / UserControl.ScaleHeight)
        If CrystalParam.RadialGYPercent > 80 Then CrystalParam.RadialGYPercent = 80
    ElseIf m_Style = WMP Then 'WMP
        CrystalParam.Ref_Intensity = 50
        CrystalParam.Ref_Left = -CornerRadius / 2
        CrystalParam.Ref_Top = -CornerRadius - 2
        CrystalParam.Ref_Height = (2 * CornerRadius) + 1
        CrystalParam.Ref_Width = Picture1.ScaleWidth + 2 * CornerRadius
        CrystalParam.Ref_Radius = CornerRadius
        CrystalParam.RadialGXPercent = 80
        CrystalParam.RadialGYPercent = 70
    ElseIf m_Style = Mac_Variation Then
        CrystalParam.Ref_Intensity = 70
        CrystalParam.Ref_Left = (CornerRadius \ 3) - 1
        CrystalParam.Ref_Height = CornerRadius - 2
        CrystalParam.Ref_Width = Picture1.ScaleWidth + 2 * CornerRadius
        CrystalParam.Ref_Top = 0
        CrystalParam.Ref_Radius = (CornerRadius \ 2)
        CrystalParam.RadialGXPercent = 100
        CrystalParam.RadialGYPercent = 70
    ElseIf m_Style = Crystal Then
        CrystalParam.Ref_Intensity = 50
        CrystalParam.Ref_Left = CornerRadius / 2
        CrystalParam.Ref_Height = CornerRadius
        CrystalParam.Ref_Width = Picture1.ScaleWidth + 2 * CornerRadius
        CrystalParam.Ref_Top = 1
        CrystalParam.Ref_Radius = CornerRadius \ 2
        CrystalParam.RadialGXPercent = 100
        CrystalParam.RadialGYPercent = 60
    End If
    Select Case vState
        Case eHover
            DrawCrystal 0, 0, Picture1.ScaleWidth - 1, UserControl.ScaleHeight - 1, m_ColorButtonHover, CrystalParam
        Case ePressed, eChecked
            DrawCrystal 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, ColorButtonDown, CrystalParam
        Case eNormal, eFocus
            DrawCrystal 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, m_ColorButtonUp, CrystalParam
    End Select
End Function

Private Sub DrawCrystal(x As Long, y As Long, Width As Long, Height As Long, Color As Long, CrystalParam As tCrystalParam)
Dim i As Long, j As Long, HighlightColor As Long, BorderColor As Long
Dim ptColor As Long, RadialGRadius As Long, LinearGPercent As Long
Dim RGXPercent As Single, RGYPercent As Single

    RGYPercent = (100 - CrystalParam.RadialGYPercent) / (Height * 2)
    RGXPercent = (100 - CrystalParam.RadialGXPercent) / Width
    If m_BorderBrightness >= 0 Then
        BorderColor = BlendColors(Color, vbWhite, m_BorderBrightness)
    Else
        BorderColor = BlendColors(Color, vbBlack, -m_BorderBrightness)
    End If
    For j = 0 To Height
        For i = 0 To Width \ 2
            If IsInRoundRect(i, j, 1, 1, Width - 2, Height - 2, CornerRadius) Then
                If j > CrystalParam.Ref_Top And j < CrystalParam.Ref_Top + CrystalParam.Ref_Height And CornerRadius Then
                    HighlightColor = BlendColors(vbWhite, Color, CrystalParam.Ref_MixColorFrom + j * CrystalParam.Ref_Intensity \ CornerRadius)
                End If
                'Drawing the button properly
                If IsInRoundRect(i, j, CrystalParam.Ref_Left, CrystalParam.Ref_Top, CrystalParam.Ref_Width, CrystalParam.Ref_Height, CrystalParam.Ref_Radius) Then
                    ptColor = HighlightColor 'draw reflected highlight
                Else
                    RadialGRadius = ((j - y - Height) * RGYPercent) ^ 2 + _
                                    ((i - x - Width \ 2) * RGXPercent) ^ 2
                    ptColor = BlendColors(m_ColorBright, Color, RadialGRadius)
                    If i < CornerRadius Then
                        ptColor = BlendColors(vbBlack, ptColor, (j * 3 \ 2 + i) * 70 \ CornerRadius)
                    End If
                End If
                SetPixelV Picture1.hdc, i + x, j + y, ptColor
                SetPixelV Picture1.hdc, x + Width - i, j + y, ptColor
            ElseIf IsInRoundRect(i, j, 0, 0, Width, Height, CornerRadius) Then
                'this draw a thin border
                SetPixelV Picture1.hdc, i + x, j + y, BorderColor
                SetPixelV Picture1.hdc, x + Width - i, j + y, BorderColor
            End If
        Next i
    Next j
End Sub

Private Function DrawPlasticButton(vState As eState)
    Select Case vState
        Case eHover
            DrawPlastic 0, 0, Picture1.ScaleWidth - 1, UserControl.ScaleHeight - 1, m_ColorButtonHover
        Case ePressed, eChecked
            DrawPlastic 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, ColorButtonDown
        Case eNormal, eFocus
            DrawPlastic 0, 0, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, m_ColorButtonUp
    End Select
End Function

Private Sub DrawPlastic(x As Long, y As Long, Width As Long, Height As Long, Color As Long)
Dim i As Long, j As Long, HighlightColor As Long, ShadowColor As Long
Dim ptColor As Long, LinearGPercent As Long
    ShadowColor = BlendColors(vbBlack, Color, 50)
    
    For j = 0 To Height
        If j < CornerRadius Then
            HighlightColor = BlendColors(vbWhite, Color, j * 30 \ CornerRadius)
        End If
        LinearGPercent = Abs((2 * j - Height) * 100 \ Height)
        For i = 0 To Width \ 2
            If IsInRoundRect(i, j, 1, 1, Width - 2, Height - 2, CornerRadius) Then
                'Drawing the button properly
                If IsInRoundRect(i, j, 4, 2, Width - CornerRadius, 2 * CornerRadius - 1, 2 * CornerRadius \ 3) _
                And Not IsInRoundRect(i, j, 4, CornerRadius \ 2, Width - CornerRadius, 2 * CornerRadius - 1, 2 * CornerRadius \ 3) Then
                    ptColor = HighlightColor 'draw reflected highlight
                Else
                    ptColor = BlendColors(Color, m_ColorBright, LinearGPercent)
                End If
                SetPixelV Picture1.hdc, i + x, j + y, ptColor
                SetPixelV Picture1.hdc, x + Width - i, j + y, ptColor
            ElseIf IsInRoundRect(i, j, 0, 0, Width, Height, CornerRadius) Then
                'this draw a thin border
                SetPixelV Picture1.hdc, i + x, j + y, ShadowColor
                SetPixelV Picture1.hdc, x + Width - i, j + y, ShadowColor
            End If
        Next i
    Next j
End Sub

Private Sub CreateRoundedRegion(Width As Long, Height As Long, Radius As Long)
Dim i As Long, j As Long, i2 As Long, j2 As Long
Dim hRgn As Long
    CornerRadius = Radius
    'Create initial region
    hRgn = CreateRectRgn(0, 0, Width, Height)
    For j = 0 To Height
        For i = 0 To Width \ 2
            If IsInRoundRect(i, j, 0, 0, Width, Height, CornerRadius) = False Then
                'This will substract the pixels outside the rounded rectangle to make the
                'button transparent.
                If j <> j2 Then
                    'If 2 * i2 <> Width Then i2 = i2 + 1
                    ExcludePixelsFromRegion hRgn, Width - i2, j2, Width - i, j
                    If 2 * i2 <> Width Then i2 = i2 + 1
                    ExcludePixelsFromRegion hRgn, i, j, i2, j2
                End If
                i2 = i
                j2 = j
            End If
        Next i
    Next j
    Call SetWindowRgn(UserControl.hwnd, hRgn, True)
    DeleteObject hRgn
End Sub

Private Function IsInRoundRect(i As Long, j As Long, x As Long, y As Long, Width As Long, Height As Long, Radius As Long) As Boolean
Dim offX As Long, offY As Long
    offX = i - x
    offY = j - y
    If offY > Radius And offY + Radius < Height And _
       offX > Radius And offX + Radius < Width Then
       'This is to catch early most cases
        IsInRoundRect = True
    ElseIf offX < Radius And offY <= Radius Then
        If IsInCircle(offX - Radius, offY, Radius) Then IsInRoundRect = True
    ElseIf offX + Radius > Width And offY <= Radius Then
        If IsInCircle(offX - Width + Radius, offY, Radius) Then IsInRoundRect = True
    ElseIf offX < Radius And offY + Radius >= Height Then
        If IsInCircle(offX - Radius, offY - Height + Radius * 2, Radius) Then IsInRoundRect = True
    ElseIf offX + Radius > Width And offY + Radius >= Height Then
        If IsInCircle(offX - Width + Radius, offY - Height + Radius * 2, Radius) Then IsInRoundRect = True
    Else
        If offX > 0 And offX < Width And offY > 0 And offY < Height Then IsInRoundRect = True
    End If
End Function

Private Function IsInCircle(ByRef x As Long, ByRef y As Long, ByRef R As Long) As Boolean
Dim lResult As Long
    'this detect a circunference that has y centered on y=0 and x=0
    lResult = (R ^ 2) - (x ^ 2)
    If lResult >= 0 Then
        lResult = Sqr(lResult)
        If Abs(y - R) < lResult Then IsInCircle = True
    End If
End Function

Public Function BlendColors(ByRef Color1 As Long, ByRef Color2 As Long, ByRef Percentage As Long) As Long
Dim R(2) As Long, G(2) As Long, B(2) As Long
    
    Percentage = SetBound(Percentage, 0, 100)
    
    GetRGB R(0), G(0), B(0), Color1
    GetRGB R(1), G(1), B(1), Color2
    
    R(2) = R(0) + (R(1) - R(0)) * Percentage \ 100
    G(2) = G(0) + (G(1) - G(0)) * Percentage \ 100
    B(2) = B(0) + (B(1) - B(0)) * Percentage \ 100
    
    BlendColors = RGB(R(2), G(2), B(2))
End Function

Private Function SetBound(ByRef Num As Long, ByRef MinNum As Long, ByRef MaxNum As Long) As Long
    If Num < MinNum Then
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        SetBound = MaxNum
    Else
        SetBound = Num
    End If
End Function

Public Sub GetRGB(ByRef R As Long, ByRef G As Long, ByRef B As Long, ByRef Color As Long)
Dim TempValue As Long
    TranslateColor Color, 0, TempValue
    R = TempValue And &HFF&
    G = (TempValue And &HFF00&) \ &H100&
    B = (TempValue And &HFF0000) \ &H10000
End Sub

Private Sub ExcludePixelsFromRegion(hRgn As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    Dim hRgnTemp As Long
    hRgnTemp = CreateRectRgn(x1, y1, x2, y2)
    CombineRgn hRgn, hRgn, hRgnTemp, RGN_XOR
    DeleteObject hRgnTemp
End Sub

Private Function HiWord(lDWord As Long) As Integer
  HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Private Function LoWord(lDWord As Long) As Integer
  If lDWord And &H8000& Then
    LoWord = lDWord Or &HFFFF0000
  Else
    LoWord = lDWord And &HFFFF&
  End If
End Function
'Read the properties from the property bag - also, a good place to start the subclassing (if we're running)
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim w As Long
  Dim h As Long
  Dim s As String
  
    Set Picture1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", UserControl.Name)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_StdPicture = PropBag.ReadProperty("Picture", Nothing)
    m_PictureAlignment = PropBag.ReadProperty("PictureAlignment", m_def_PictureAlignment)
    m_Style = PropBag.ReadProperty("Style", 0)
    m_Checked = PropBag.ReadProperty("Checked", m_Checked)
    m_ColorButtonHover = PropBag.ReadProperty("ColorButtonHover", &HFFC090)
    m_ColorButtonUp = PropBag.ReadProperty("ColorButtonUp", &HE99950)
    m_ColorButtonDown = PropBag.ReadProperty("ColorButtonDown", &HE99950)
    m_ColorBright = PropBag.ReadProperty("ColorBright", &HFFEDB0)
    m_BorderBrightness = PropBag.ReadProperty("BorderBrightness", 0)
    m_DisplayHand = PropBag.ReadProperty("DisplayHand", False)
    m_ColorScheme = PropBag.ReadProperty("ColorScheme", 0)
    If m_DisplayHand Then UserControl.MousePointer = vbCustom Else UserControl.MousePointer = vbArrow
    UserControl.ForeColor = m_ForeColor
    
  If Ambient.UserMode Then                                                              'If we're not in design mode
    bTrack = True
    bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  
    If Not bTrackUser32 Then
      If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
        bTrack = False
      End If
    End If
  
    If bTrack Then
      'OS supports mouse leave, so let's subclass for it
      With UserControl
        'Subclass the UserControl
        sc_Subclass .hwnd
        sc_AddMsg .hwnd, WM_MOUSEMOVE
        sc_AddMsg .hwnd, WM_MOUSELEAVE
      End With
    End If
  End If
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
  sc_Terminate                                                              'Terminate all subclassing
End Sub

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    FreeLibrary hMod
  End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      TrackMouseEvent tme
    Else
      TrackMouseEventComCtl tme
    End If
  End If
End Sub

'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nID           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nID                                    'Get the process ID associated with the window handle
  If nID <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(i)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal newValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = newValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = 0 Then                                                  'If the element is free...
        zData(i) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(i) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim i      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If zData(i) = uMsg Then                                               'If the message is found...
        zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
Dim x As Long, y As Long
  Select Case uMsg
    Case WM_MOUSEMOVE
        If wParam <> MK_LBUTTON And Not IsHover Then
            x = LoWord(lParam)
            y = HiWord(lParam)
            If x > 0 And x < UserControl.ScaleWidth And y > 0 And y < UserControl.ScaleHeight Then
                IsHover = True
                TrackMouseLeave lng_hWnd
                RaiseEvent MouseEnter
                DrawButton (eHover)
            End If
        End If

  Case WM_MOUSELEAVE
        IsHover = False
        RaiseEvent MouseLeave
        If Not m_Checked Then If m_hasFocus Then DrawButton (eFocus) Else DrawButton (eNormal)
  End Select
End Sub
