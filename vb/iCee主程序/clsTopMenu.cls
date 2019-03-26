VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'菜单美化
Private m_MenuID As Long                '菜单项ID
Private m_Caption As String             '菜单项标题
Private m_width As Integer              '菜单项宽度
Private m_height As Integer             '菜单项高度
Private m_Picture As PictureBox         '菜单项图标
Private m_Enabled As Boolean            '菜单项是否可用
Private m_Selected As Boolean           '菜单项是否勾选
Private m_DrawStruct As DRAWITEMSTRUCT  '菜单项的画图信息

Private Const PS_SOLID = 0
Private Const TRANSPARENT = 1
Private Const DSS_DISABLED = &H20
Private Const DST_TEXT = &H1
Private Const ODS_SELECTED = &H1

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Private Type DRAWITEMSTRUCT
     CtlType As Long     '控件类型
     CtlID As Long       '控件id
     itemID As Long      '菜单项.列表框或组合框中某一项的索引值
     itemAction As Long  '控件行为
     itemState As Long   '控件状态
     hwndItem As Long    '父窗口句柄或菜单句柄
     hDC As Long         '控件对应的绘图设备句柄
     rcItem As RECT      '控件所占据的矩形区域
     ItemData As Long    '列表框或组合框中某一项的值
End Type

Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Public Property Get MenuID() As Long
  MenuID = m_MenuID
End Property

Public Property Get Caption() As String
  Caption = m_Caption
End Property

Public Property Get Width() As Integer
  Width = m_width
End Property

Public Property Get Height() As Integer
  Height = m_height
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal flag As Boolean)
  m_Enabled = flag
End Property

Public Property Get Selected() As Boolean
  Selected = m_Selected
End Property

Public Property Let Selected(ByVal flag As Boolean)
  m_Selected = flag
End Property

'初始化菜单项基本信息
Public Sub InitMenu(ByVal nMenuID As Long, ByVal sCaption As String, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal obj As PictureBox)
  m_MenuID = nMenuID
  m_Caption = sCaption
  m_width = nWidth
  m_height = nHeight
  Set m_Picture = obj
End Sub

'初始化菜单项画图信息
Public Sub InitStruct(ByVal hDC As Long, ByVal nAction As Long, ByVal nID As Long, ByVal nState As Long, ByVal nLeft As Long, ByVal nTop As Long, ByVal nBottom As Long, ByVal nRight As Long)
  m_DrawStruct.hDC = hDC
  m_DrawStruct.itemAction = nAction
  m_DrawStruct.itemID = nID
  m_DrawStruct.itemState = nState
  m_DrawStruct.rcItem.Left = nLeft
  m_DrawStruct.rcItem.Top = nTop
  m_DrawStruct.rcItem.Bottom = nBottom
  m_DrawStruct.rcItem.Right = nRight
End Sub

Private Sub SetMenuText(ByVal hDC As Long, ByVal clr As Long, Rc As RECT, ByVal PosX As Long)
  Dim OldColor As Long
  SetBkMode hDC, TRANSPARENT
  OldColor = SetTextColor(hDC, clr)
  If Not (m_Enabled) Then
    DrawState hDC, 0, 0, m_Caption, LenB(StrConv(m_Caption, vbFromUnicode)), _
                PosX, Rc.Top + 2, 0, 0, DSS_DISABLED Or DST_TEXT
  Else
    TextOut hDC, PosX, Rc.Top + 2, m_Caption, LenB(StrConv(m_Caption, vbFromUnicode))
  End If
  SetTextColor hDC, OldColor
End Sub

Public Sub DrawMenu()
  Dim DrawInfo As DRAWITEMSTRUCT
  Dim drawRect As RECT, Rc As RECT
  Dim IsSelected As Boolean, IsPicture As Boolean
  Dim OldColor As Long, TMP As Long
  Dim hPen As Long, hOldPen As Long
  
  DrawInfo = m_DrawStruct
  IsSelected = (DrawInfo.itemState And ODS_SELECTED)        '是否选中菜单项
  IsPicture = Not (m_Picture Is Nothing)                    '是否带图标
  drawRect = DrawInfo.rcItem
  
  If IsSelected Then
  
    If m_Caption <> "" Then    '如果不是分割线
    
      FillRect DrawInfo.hDC, drawRect, CreateSolidBrush(g_clrBkgSelect)
      FrameRect DrawInfo.hDC, drawRect, CreateSolidBrush(g_clrFrame)
      
      If m_Selected Then
        StretchBlt DrawInfo.hDC, 3, drawRect.Top + 1, 16, 16, Frmm.Pic(0).hDC, 0, 0, Frmm.Pic(0).ScaleWidth, Frmm.Pic(0).ScaleHeight, vbSrcCopy
      ElseIf IsPicture Then
        StretchBlt DrawInfo.hDC, 3, drawRect.Top + 1, 16, 16, m_Picture.hDC, 0, 0, m_Picture.ScaleWidth, m_Picture.ScaleHeight, vbSrcCopy
      End If
      
      '设置菜单文本颜色并根据状态画出菜单文本
      SetMenuText DrawInfo.hDC, g_clrTxtSelect, drawRect, 30
      
    End If
    
  Else
    
    Rc = drawRect
    TMP = Rc.Right   '画左半部分(就是左边那个长条,用于显示图标)
    Rc.Right = 24
    FillRect DrawInfo.hDC, Rc, CreateSolidBrush(g_clrLeft)
    
    Rc.Left = 24     '画右半部分(用于显示文字)
    Rc.Right = TMP
    FillRect DrawInfo.hDC, Rc, CreateSolidBrush(g_clrBkgNormal)
    
    If m_Caption = "" Then   '如果是分割线
    
      hPen = CreatePen(PS_SOLID, 1, g_clrSep)
      hOldPen = SelectObject(DrawInfo.hDC, hPen)
      MoveToEx DrawInfo.hDC, Rc.Left, Rc.Top + 2, 0
      LineTo DrawInfo.hDC, Rc.Right, Rc.Top + 2
      SelectObject DrawInfo.hDC, hOldPen
      DeleteObject hPen
      
    Else
    
      If m_Selected Then
        StretchBlt DrawInfo.hDC, 4, drawRect.Top + 2, 16, 16, Frmm.Pic(0).hDC, 0, 0, Frmm.Pic(0).ScaleWidth, Frmm.Pic(0).ScaleHeight, vbSrcCopy
      ElseIf IsPicture Then
        StretchBlt DrawInfo.hDC, 4, drawRect.Top + 2, 16, 16, m_Picture.hDC, 0, 0, m_Picture.ScaleWidth, m_Picture.ScaleHeight, vbSrcCopy
      End If
      
      '设置菜单文本颜色并根据状态画出菜单文本
      SetMenuText DrawInfo.hDC, g_clrTxtNormal, drawRect, 30
      
    End If
    
  End If
  
End Sub

Private Sub Class_Initialize()
  m_Enabled = True
End Sub




