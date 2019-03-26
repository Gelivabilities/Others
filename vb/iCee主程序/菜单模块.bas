Attribute VB_Name = "菜单模块"
Option Explicit
Public g_clrFrame As Long '选中菜单项时所画矩形框的颜色
Public g_clrBkgSelect As Long '选中菜单项时的背景色
Public g_clrBkgNormal As Long '正常情况下菜单项的背景色
Public g_clrTxtSelect As Long '选中菜单项时文本的颜色
Public g_clrTxtNormal As Long '正常情况下菜单项文本的颜色
Public g_clrLeft As Long '正常情况下菜单项左边的颜色
Public g_clrSep As Long '分割线的颜色
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const ODT_MENU = 1
Private Const WM_COMMAND = &H111
Private Const WM_DESTROY = &H2
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Public Enum MenuFlags
  MF_INSERT = &H0
  MF_ENABLED = &H0
  MF_UNCHECKED = &H0
  MF_BYCOMMAND = &H0
  MF_STRING = &H0
  MF_UNHILITE = &H0
  MF_GRAYED = &H1
  MF_DISABLED = &H2
  MF_BITMAP = &H4
  MF_CHECKED = &H8
  MF_POPUP = &H10
  MF_MENUBARBREAK = &H20
  MF_MENUBREAK = &H40
  MF_HILITE = &H80
  MF_CHANGE = &H80
  MF_END = &H80                    ' Obsolete -- only used by old RES files
  MF_APPEND = &H100
  MF_OWNERDRAW = &H100
  MF_DELETE = &H200
  MF_USECHECKBITMAPS = &H200
  MF_BYPOSITION = &H400
  MF_SEPARATOR = &H800
  MF_REMOVE = &H1000
  MF_DEFAULT = &H1000
  MF_SYSMENU = &H2000
  MF_HELP = &H4000
  MF_RIGHTJUSTIFY = &H4000
  MF_MOUSESELECT = &H8000&
End Enum
Public Const CFalse = False
Public Const CTrue = 1
Public Enum MII_Mask
  MIIM_STATE = &H1
  MIIM_ID = &H2
  MIIM_SUBMENU = &H4
  MIIM_CHECKMARKS = &H8
  MIIM_TYPE = &H10
  MIIM_DATA = &H20
End Enum
Private Type DRAWITEMSTRUCT
 CtlType As Long '控件类型
 CtlID As Long '控件id
 itemID As Long '菜单项.列表框或组合框中某一项的索引值
 itemAction As Long '控件行为
 itemState As Long '控件状态
 hwndItem As Long '父窗口句柄或菜单句柄
 hdc As Long '控件对应的绘图设备句柄
 rcItem As RECT '控件所占据的矩形区域
 ItemData As Long '列表框或组合框中某一项的值
End Type
Private Type MEASUREITEMSTRUCT
 CtlType As Long
 CtlID As Long
 itemID As Long
 itemWidth As Long
 ItemHeight As Long
 ItemData As Long
End Type
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public MyTopMenu() As clsTopMenu '该数组保存各个菜单项的信息
Public g_CntOfTopMenu As Integer '保存菜单项的数量(从0算起,实际数目是g_CntOfTopMenu+1)
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Public right1 As Boolean
Public Right2 As Boolean
Public iStop As Boolean
Public MyRightMenu() As clsTopMenu
Public g_CntOfRightMenu As Integer
Public p_ID As Long
Public m_hMenu As Long
Private Const POPUP_LEFTALIGN = &H0&
'Public Type POINTAPI
'X As Long
'Y As Long
'End Type
Private Type MENUITEMINFO
CBSIZE As Long
fMask As Long
fType As Long
fState As Long
wID As Long
hSubMenu As Long
hbmpChecked As Long
hbmpUnchecked As Long
dwItemData As Long
dwTypeData As String
cch As Long
End Type
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, mi As MENUINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Enum MENUINFO_STYLES
MNS_NOCHECK = &H80000000
MNS_MODELESS = &H40000000
MNS_DRAGDROP = &H20000000
MNS_AUTODISMISS = &H10000000
MNS_NOTIFYBYPOS = &H8000000
MNS_CHECKORBMP = &H4000000
End Enum
Private Enum MENUINFO_MASKS
MIM_MAXHEIGHT = &H1
MIM_BACKGROUND = &H2
MIM_HELPID = &H4
MIM_MENUDATA = &H8
MIM_STYLE = &H10
MIM_APPLYTOSUBMENUS = &H80000000
End Enum
Private Type MENUINFO
CBSIZE As Long
fMask As MENUINFO_MASKS
dwStyle As MENUINFO_STYLES
cyMax As Long
hbrBack As Long
dwContextHelpID As Long
dwMenuData As Long
End Type

'泡泡
Public Const WM_RBUTTONUP = &H205
Public Const WM_USER = &H400
Public Const WM_NOTIFYICON = WM_USER + 1               '   自定义消息
Public Const WM_LBUTTONDBLCLK = &H203
'   关于气球提示的自定义消息,   2000下不产生这些消息
Public Const NIN_BALLOONSHOW = (WM_USER + &H2)               '   当   Balloon   Tips   弹出时执行
Public Const NIN_BALLOONHIDE = (WM_USER + &H3)               '   当   Balloon   Tips   消失时执行（如   SysTrayIcon   被删除）,
'   但指定的   TimeOut   时间到或鼠标点击   Balloon   Tips   后的消失不发送此消息
Public Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)               '   当   Balloon   Tips   的   TimeOut   时间到时执行
Public Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)               '   当鼠标点击   Balloon   Tips   时执行.
'   注意:在XP下执行时   Balloon   Tips   上有个关闭按钮,
'   如果鼠标点在按钮上将接收到   NIN_BALLOONTIMEOUT   消息.
Public preWndProc     As Long
Public Sub SetMenuBar(frm As Form, ByVal clr As Long)
Dim MyMenu As MENUINFO
MyMenu.CBSIZE = Len(MyMenu)
MyMenu.fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
MyMenu.hbrBack = CreateSolidBrush(clr)
SetMenuInfo GetMenu(frm.hwnd), MyMenu
End Sub
Private Sub AddMenuItem(ByVal hMenu As Long, ByVal sCaption As String, ByVal nWidth As Integer, ByVal nHeight As Integer, Optional obj As PictureBox = Nothing)
Dim MENUINFO As MENUITEMINFO
With MENUINFO
.CBSIZE = LenB(MENUINFO) '多预留点空间
.fMask = MIIM_TYPE Or MIIM_ID
.fType = MF_OWNERDRAW
.wID = p_ID
End With
InsertMenuItem hMenu, p_ID, False, MENUINFO
ReDim Preserve MyRightMenu(g_CntOfRightMenu)
Set MyRightMenu(g_CntOfRightMenu) = New clsTopMenu
MyRightMenu(g_CntOfRightMenu).InitMenu p_ID, sCaption, nWidth, nHeight, obj
g_CntOfRightMenu = g_CntOfRightMenu + 1
p_ID = p_ID + 1
End Sub
Public Sub ReleaseObj_RightMenu()
Dim i As Integer
If g_CntOfRightMenu Then
For i = 0 To g_CntOfRightMenu - 1
Set MyRightMenu(i) = Nothing
Next
End If
End Sub
Public Sub ReleaseObj_TopMenu()
Dim i As Integer
If g_CntOfTopMenu Then
For i = 0 To g_CntOfTopMenu - 1
Set MyTopMenu(i) = Nothing
Next
End If
End Sub
'顶级菜单部分的内容
'*****************************************************************************************************************
Public Sub RegisterMenu(ByVal hMenu As Long, ByVal nPosition As Integer, ByVal sCaption As String, ByVal nWidth As Integer, ByVal nHeight As Integer, Optional obj As PictureBox = Nothing)
Dim mnuID As Long
mnuID = GetMenuItemID(hMenu, nPosition) '得到菜单项的ID
ModifyMenu hMenu, nPosition, MF_OWNERDRAW Or MF_BYPOSITION, mnuID, 0 '使菜单项具有自画属性
'保存菜单项的信息
ReDim Preserve MyTopMenu(g_CntOfTopMenu)
Set MyTopMenu(g_CntOfTopMenu) = New clsTopMenu
MyTopMenu(g_CntOfTopMenu).InitMenu mnuID, sCaption, nWidth, nHeight, obj
g_CntOfTopMenu = g_CntOfTopMenu + 1
End Sub
Public Sub SubClassWindow(frm As Form)
If GetProp(frm.hwnd, "OrigProcAddr") = 0 Then
SetProp frm.hwnd, "OrigProcAddr", SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
End If
End Sub
Private Function NewWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim origProc As Long
Dim isSubclassed As Long
Dim DrawInfo As DRAWITEMSTRUCT
Dim MeasureInfo As MEASUREITEMSTRUCT
Dim i As Integer
origProc = GetProp(hwnd, "OrigProcAddr")
If origProc <> 0 Then
If uMsg = WM_MEASUREITEM Then '在控件或菜单被创建的时候,向自绘按钮,组合框,列表框,列表视图(list view)
'或菜单项的所有者发送WM_MEASUREITEM消息
CopyMemory MeasureInfo, ByVal lParam, Len(MeasureInfo)
If MeasureInfo.CtlType <> ODT_MENU Then Exit Function
If iStop Then
For i = 0 To g_CntOfTopMenu - 1
If MyTopMenu(i).MenuID = MeasureInfo.itemID Then
MeasureInfo.itemWidth = MyTopMenu(i).Width
MeasureInfo.ItemHeight = MyTopMenu(i).Height
Exit For
End If
Next
Else
For i = 0 To g_CntOfRightMenu - 1
If MyRightMenu(i).MenuID = MeasureInfo.itemID Then
MeasureInfo.itemWidth = MyRightMenu(i).Width
MeasureInfo.ItemHeight = MyRightMenu(i).Height
Exit For
End If
Next
End If
CopyMemory ByVal lParam, MeasureInfo, Len(MeasureInfo)
ElseIf uMsg = WM_COMMAND Then
NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
ElseIf uMsg = WM_DRAWITEM Then '当具有自绘风格的按钮.组合框.列表框或者菜单的可见部分发生改变时，
'就会发送WM_DRAWITEM消息给自绘控件所在的窗体
CopyMemory DrawInfo, ByVal lParam, Len(DrawInfo)
If DrawInfo.CtlType <> ODT_MENU Then Exit Function
If iStop Then
For i = 0 To g_CntOfTopMenu - 1
If MyTopMenu(i).MenuID = DrawInfo.itemID Then
MyTopMenu(i).InitStruct DrawInfo.hdc, DrawInfo.itemAction, DrawInfo.itemID, DrawInfo.itemState, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, DrawInfo.rcItem.Bottom, DrawInfo.rcItem.Right
MyTopMenu(i).DrawMenu
Exit For
End If
Next
Else
For i = 0 To g_CntOfRightMenu - 1
If MyRightMenu(i).MenuID = DrawInfo.itemID Then
MyRightMenu(i).InitStruct DrawInfo.hdc, DrawInfo.itemAction, DrawInfo.itemID, DrawInfo.itemState, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, DrawInfo.rcItem.Bottom, DrawInfo.rcItem.Right
MyRightMenu(i).DrawMenu
Exit For
End If
Next
End If
ElseIf uMsg = WM_DESTROY Then
SetWindowLong hwnd, GWL_WNDPROC, origProc
NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
RemoveProp hwnd, "OrigProcAddr"
ReleaseObj_RightMenu
ReleaseObj_TopMenu
Else
NewWindowProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
End If
Else
'如果有意外发生的话
NewWindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
End If
End Function
