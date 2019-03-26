VERSION 5.00
Begin VB.Form FrmPicInfo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H007A7417&
   BorderStyle     =   0  'None
   Caption         =   "图像信息"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   Icon            =   "FrmPicInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox C2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5445
      Picture         =   "FrmPicInfo.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   21
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5445
      Picture         =   "FrmPicInfo.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   20
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5445
      Picture         =   "FrmPicInfo.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   19
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox PICD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4800
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin ICEE.ICEE_COMMAND OK 
      Height          =   495
      Left            =   4200
      TabIndex        =   17
      Top             =   5400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
   End
   Begin VB.TextBox TxtTit 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "示例文本"
      Top             =   1005
      Width           =   3495
   End
   Begin VB.TextBox TxtPath 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "示例文本"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox TxtDri 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "无法读取"
      Top             =   4080
      Width           =   3495
   End
   Begin VB.PictureBox picSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H007A7417&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   240
   End
   Begin VB.PictureBox picLarge 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H007A7417&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   840
      Width           =   480
   End
   Begin VB.Shape SB 
      BackColor       =   &H00241D0A&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   0  'Transparent
      Height          =   735
      Index           =   3
      Left            =   30
      Top             =   5295
      Width           =   6150
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   855
      Index           =   2
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Index           =   1
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   3735
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   495
      Index           =   0
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label LT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图像位置："
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   16
      Top             =   3435
      Width           =   900
   End
   Begin VB.Label LT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图像尺寸："
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   15
      Top             =   1935
      Width           =   900
   End
   Begin VB.Label LT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件类型："
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   600
      TabIndex        =   14
      Top             =   1575
      Width           =   900
   End
   Begin VB.Label LT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件创建时间："
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   3
      Left            =   600
      TabIndex        =   13
      Top             =   2295
      Width           =   1260
   End
   Begin VB.Label LT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "最后修改时间："
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   4
      Left            =   600
      TabIndex        =   12
      Top             =   2655
      Width           =   1260
   End
   Begin VB.Label LT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件描述："
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   5
      Left            =   600
      TabIndex        =   11
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label LbType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label Lbsize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1920
      TabIndex        =   9
      Top             =   1935
      Width           =   90
   End
   Begin VB.Label LbDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1920
      TabIndex        =   8
      Top             =   2295
      Width           =   90
   End
   Begin VB.Label LbChange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1920
      TabIndex        =   7
      Top             =   2655
      Width           =   90
   End
   Begin VB.Label LT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "最后访问时间："
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   3015
      Width           =   1260
   End
   Begin VB.Label LbLast 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1920
      TabIndex        =   5
      Top             =   3015
      Width           =   90
   End
End
Attribute VB_Name = "FrmPicInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type
Dim XX, YY, O, L, t As Long
Private Const RGN_OR = 2

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject As Long)
Private Declare Function FileTimeToSystemTime& Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME)
Private Declare Function lopen& Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long)
Private Declare Function GetFileTime& Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME)
Private Const READ_CONTROL = &H20000
Private Sub Form_Load()
On Error Resume Next
MakeTransparent Me.hwnd, 250
Call PaintPng(App.Path & "\SKIN\FL_T.PNG", Me.hdc, 8, 8)
oldproc = GetWindowLong(TxtTit.hwnd, GWL_WNDPROC)
oldproc = GetWindowLong(txtPath.hwnd, GWL_WNDPROC)
oldproc = GetWindowLong(TxtDri.hwnd, GWL_WNDPROC)
SetWindowLong TxtTit.hwnd, GWL_WNDPROC, AddressOf TextWndProc
SetWindowLong txtPath.hwnd, GWL_WNDPROC, AddressOf TextWndProc
SetWindowLong TxtDri.hwnd, GWL_WNDPROC, AddressOf TextWndProc
Me.Move frmGraphic.Left + (frmGraphic.Width - Me.Width) / 2, frmGraphic.Top + (frmGraphic.Height - Me.Height) / 2
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
'窗体总在最上
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H231C09, B
Call SHOWINFO
Call Seekico
frmGraphic.Enabled = False
OK.HASLINE = False
OK.SETTXT "确   定"
End Sub
Sub Seekico()
Dim hImgSmall As Long     ' The handle to the system image list
Dim hImgLarge As Long
Dim filename As String    ' The file name to get icon from
Dim r As Long
filename = frmGraphic.Select_Pic
hImgSmall& = SHGetFileInfo(filename, 0, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
hImgLarge& = SHGetFileInfo(filename, 0, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
LbType.Caption = Left(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr(0)) - 1)
picSmall.Cls
picLarge.Cls
hImgSmall& = SHGetFileInfo(filename$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
hImgLarge& = SHGetFileInfo(filename$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
r& = ImageList_Draw(hImgSmall&, shinfo.hIcon, picSmall.hdc, 0, 0, ILD_TRANSPARENT)
r& = ImageList_Draw(hImgLarge&, shinfo.hIcon, picLarge.hdc, 0, 0, ILD_TRANSPARENT)
LbType.Caption = shinfo.szTypeName ' Right(FileName, 3)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = True
C2.Visible = False
C3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong TxtTit.hwnd, GWL_WNDPROC, oldproc
SetWindowLong txtPath.hwnd, GWL_WNDPROC, oldproc
SetWindowLong TxtDri.hwnd, GWL_WNDPROC, oldproc
frmGraphic.Enabled = True
End Sub

Private Sub LbChange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub LbDay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub LbLast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Lbsize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub LbType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub LT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub OK_CLICK()
Unload Me
End Sub

Private Sub picLarge_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub picSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Sub SHOWINFO()
Dim hwnd As Long   '文件句柄
Dim CT As FILETIME  '文件建立时间
Dim AT As FILETIME  '文件访问时间
Dim WT As FILETIME  '最后修改时间
Dim st As SYSTEMTIME
Dim Str1 As String
Dim str2 As String
Str1 = frmGraphic.Select_Pic
Dim RetVal As Long  '接收返回值
    hwnd = lopen(frmGraphic.Select_Pic, READ_CONTROL)
    RetVal = GetFileTime(hwnd, CT, AT, WT)
    RetVal = FileTimeToSystemTime(CT, st)
    If st.wHour < 16 Then
    LbDay.Caption = Trim(str(st.wYear)) + "年" + Trim(str(st.wMonth)) + "月" + Trim(str(st.wDay)) + "日 " + Trim(str(st.wHour + 8)) + ":" + Trim(str(st.wMinute)) + ":" + Trim(str(st.wSecond))
    Else
    LbDay.Caption = Trim(str(st.wYear)) + "年" + Trim(str(st.wMonth)) + "月" + Trim(str(st.wDay)) + "日 " + Trim(str(24 - st.wHour)) + ":" + Trim(str(st.wMinute)) + ":" + Trim(str(st.wSecond))
    End If
    RetVal = FileTimeToSystemTime(AT, st)
    LbLast.Caption = Trim(str(st.wYear)) + "年" + Trim(str(st.wMonth)) + "月" + Trim(str(Day(Date))) + "日 "
    RetVal = FileTimeToSystemTime(WT, st)
    If st.wHour < 16 Then
    Me.LbChange.Caption = Trim(str(st.wYear)) + "年" + Trim(str(st.wMonth)) + "月" + Trim(str(st.wDay)) + "日 " + Trim(str(st.wHour + 8)) + ":" + Trim(str(st.wMinute)) + ":" + Trim(str(st.wSecond))
    Else
    LbChange.Caption = Trim(str(st.wYear)) + "年" + Trim(str(st.wMonth)) + "月" + Trim(str(st.wDay)) + "日 " + Trim(str(24 - st.wHour)) + ":" + Trim(str(st.wMinute)) + ":" + Trim(str(st.wSecond))
    End If
    txtPath.Text = frmGraphic.filHidden.Path & frmGraphic.Select_Pic
    Select Case UCase(Right(frmGraphic.Select_Pic, 3))
    Case "PNG"
    Call OPENISPNG(PICD, frmGraphic.Select_Pic)
    Case "BMP", "JPG", "ICO", "GIF"
    PICD.PICTURE = LoadPicture(frmGraphic.Select_Pic)
    End Select
    Lbsize.Caption = Int(PICD.Width) & "×" & Int(PICD.Height) & "( 宽X高 )"
    Me.TxtTit.Text = frmGraphic.Select_Pic
    
    RetVal = CloseHandle(hwnd)  '关闭文件句柄
Exit Sub
End Sub

Private Sub TxtDri_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub

Private Sub TxtDri_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtDri.MousePointer = 0
End Sub

Private Sub TxtPath_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub

Private Sub TxtPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtPath.MousePointer = 0
End Sub

Private Sub TxtTit_GotFocus()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub

Private Sub TxtTit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TxtTit.MousePointer = 0
End Sub
Private Sub c1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = False
C2.Visible = True
End Sub
Private Sub c2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
C2.Visible = False
C3.Visible = True
End If
End Sub
Private Sub c3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C3.Visible = False
C1.Visible = True
If C3.Visible = False Then
Unload Me
End If
End Sub
