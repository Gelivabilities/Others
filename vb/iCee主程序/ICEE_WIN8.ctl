VERSION 5.00
Begin VB.UserControl ICEE_WIN8 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0094A63E&
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
   ControlContainer=   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   291
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   623
   Begin VB.Timer TMBYE 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.PictureBox PC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5880
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin ICEE.ICEE_KEY ICL 
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin VB.FileListBox FILEPATH 
      Height          =   3690
      Left            =   7440
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer TMPLAY 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   960
      Top             =   0
   End
   Begin VB.PictureBox IMG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00899F1E&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   5040
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer TMIN 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer TMOUT 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   0
   End
   Begin VB.PictureBox PDF 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   0
      Top             =   2040
      Width           =   4335
      Begin VB.Label LBC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   3975
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PUS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00565656&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3255
      Left            =   0
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   457
      TabIndex        =   7
      Top             =   0
      Width           =   6855
      Begin VB.Image K 
         Height          =   720
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Label LBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tittle"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1800
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "ICEE_WIN8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Event Click()
Public Event DBLCLICK()
Public MY_STYLE As Integer, AUTOSIZE As Boolean
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public IS_PIC As Boolean, MYTIT As String, HASLINE As Boolean, HASTIP As Boolean, MYTIP As String, IS_TIME As Boolean, IS_PIC_SHOW As Boolean, IS_MUSIC As Boolean
Dim pos As POINTAPI '定义这个变量是取得鼠标坐标
Dim OUT_COLOR As Long, IN_COLOR As Long, P_I As Long
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public MY_PIC As PictureBox

Private Sub ICL_Click()
On Error Resume Next
Dim BFPATH As String
BFPATH = BrowseFolder("浏览文件夹", frmma)
If BFPATH = "" Then Exit Sub
Call SETPATH(BFPATH)  '将文件列表设为选中的文件夹
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换了图库路径:" & BFPATH
End Sub

Private Sub LBA_Change()
LBA.Left = (UserControl.ScaleWidth - LBA.Width) / 2
MYTIT = LBA.Caption
End Sub

Private Sub LBA_Click()
RaiseEvent Click
End Sub

Private Sub LBA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
TMOUT.Enabled = True
End Sub

Private Sub LBA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBC_DblClick()
RaiseEvent DBLCLICK
End Sub

Private Sub LBC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub PDF_Click()
    RaiseEvent Click
End Sub

Private Sub PDF_DblClick()
RaiseEvent DBLCLICK
End Sub

Private Sub PDF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub PDF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
 Call UserControl_MouseMove(Button, Shift, X, Y)
 TMOUT.Enabled = True
End Sub

Private Sub PDF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub PDF_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub PDF_Resize()
If LBC.Top <> (PDF.ScaleHeight - LBC.Height) / 2 Then LBC.Move (PDF.ScaleWidth - LBC.Width) / 2, (PDF.ScaleHeight - LBC.Height) / 2
End Sub

Private Sub PUS_Resize()
K.Move 10, PUS.ScaleHeight - 10 - K.Height
End Sub

Private Sub TMBYE_Timer()
If HASTIP = False Then Exit Sub
TMIN.Enabled = False
PUS.Top = PUS.Top + 10
PDF.Top = PUS.Top + PUS.Height + 10
If PUS.Top >= 0 Then
TMBYE.Enabled = False
PUS.Top = 0
PDF.Top = PDF.Height
End If
End Sub

Private Sub TMIN_Timer()
If HASTIP = False Then Exit Sub
TMBYE.Enabled = False
PUS.Top = PUS.Top - 10
PDF.Top = PUS.Top + PUS.Height - 10
If PUS.Top <= -PUS.Height Then
TMIN.Enabled = False
PUS.Top = -PUS.Height
PDF.Top = 0
End If
End Sub

Private Sub TMOUT_Timer() '
Dim r As RECT, p As POINTAPI, L As Long, rtn As Long, H As Long, H1 As Long, r1 As Long '鼠标移出/移入透明值得改变
L = GetWindowRect(UserControl.hwnd, r)
L = GetCursorPos(p)
GetCursorPos pos
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then '移出界面
TMOUT.Enabled = False

If ICL.Visible = True Then ICL.Visible = False
If IS_PIC = False Then UserControl.BackColor = OUT_COLOR
If HASLINE = True Then UserControl.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
UserControl.Refresh
PUS.Refresh
If HASTIP = False Then Exit Sub
If IS_AM = False Then PDF.Top = UserControl.ScaleHeight: Exit Sub
If PUS.Top <> 0 Then TMIN.Enabled = False: TMBYE.Enabled = True

Else

If IS_PIC_SHOW = True Then ICL.Visible = True
If IS_PIC = False Then UserControl.BackColor = IN_COLOR
If HASLINE = True Then UserControl.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
If HASTIP = False Then Exit Sub
If IS_AM = False Then PDF.Top = 0: Exit Sub
TMIN.Enabled = True
TMBYE.Enabled = False

End If

End Sub

Private Sub TMPLAY_Timer()
On Error Resume Next
Dim I As Integer, P_PIC As String ', J As Long ', num As Integer
FILEPATH.ListIndex = P_I
P_PIC = FILEPATH.Path & "\" & FILEPATH.filename
P_I = P_I + 1
If P_I > FILEPATH.ListCount - 1 Then P_I = 0
IMG.PICTURE = LoadPicture(P_PIC)
PUS.Cls
PUS.PaintPicture IMG.PICTURE, (PUS.ScaleWidth - IMG.ScaleWidth) / 2, (PUS.ScaleHeight - IMG.ScaleHeight) / 2, IMG.Width, IMG.Height, 0, 0      ', IMG.Width, IMG.Height
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DBLCLICK
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
IS_PIC = False
TMOUT.Enabled = True
HASLINE = False
HASTIP = True
AUTOSIZE = False
MY_STYLE = 1 '1居中,0左上角,2右下角,3左下角,4右下角
PDF.Move 0, UserControl.ScaleHeight
OUT_COLOR = &H5C6105
IN_COLOR = &H94A63E
LBA.Caption = ""
LBC.Caption = "说明"
LBA.FontName = "微软雅黑"
LBC.FontName = "微软雅黑"
FILEPATH.Pattern = "*.BMP;*.JPG"
P_I = 0
FILEPATH.Path = (GetInitEntry("ICEE", "PLAYPIC", App.Path + "\MEDIA\PIC"))
ICL.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICL.SETTXT "修改图册"
End Sub
Sub SETCOLOR(NOMAL_COLOR As Long, HIGH_COLOR As Long)
UserControl.BackColor = NOMAL_COLOR
OUT_COLOR = NOMAL_COLOR
pc.BackColor = OUT_COLOR
PUS.BackColor = NOMAL_COLOR
IN_COLOR = HIGH_COLOR
ICL.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Call PaintPng(App.Path & "\SKIN\WHITE.PNG", pc.hdc, 0, 0)
PDF.BackColor = pc.POINT(1, 1)
IMG.BackColor = NOMAL_COLOR
End Sub

Sub SETPIC(PIC As String)
On Error Resume Next
IMG.PICTURE = LoadPicture(PIC)
PUS.Visible = True
PUS.Move 0, 0, ScaleWidth, ScaleHeight
PUS.Cls
PUS.BackColor = COLOR_NOR
If AUTOSIZE = True Then
Call DrawPicture(PUS.hdc, PIC, 0, 0, PUS.ScaleWidth, PUS.ScaleHeight)
Else
PUS.PaintPicture IMG.PICTURE, (PUS.ScaleWidth - IMG.Width) / 2, (PUS.ScaleHeight - IMG.Height) / 2, IMG.Width, IMG.Height, 0, 0
End If
PUS.Refresh
Set MY_PIC = IMG
End Sub
Sub SETIMG(SPIC As PictureBox)
On Error Resume Next
PUS.Visible = True
PUS.Move 0, 0, ScaleWidth, ScaleHeight
IMG.PICTURE = SPIC.image
PUS.PaintPicture IMG.PICTURE, 0, 0, PUS.ScaleWidth, PUS.ScaleHeight
PUS.Refresh
End Sub
Sub SETTXT(TXT As String)
LBA.Caption = TXT
End Sub
Sub SETTIP(TXT As String)
LBC.Caption = TXT
MYTIP = TXT
PDF_Resize
End Sub
Sub SETTXTCOLOR(Color As Long, M_COLOR As Long)
LBC.FOREColor = Color
LBA.FOREColor = M_COLOR
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
TMOUT.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Select Case MY_STYLE
Case 0
LBA.Move 5, 5
Case 1
LBA.Move (UserControl.ScaleWidth - LBA.Width) / 2, (UserControl.ScaleHeight - LBA.Height) / 2
Case 2
LBA.Move UserControl.ScaleWidth - LBA.Width - 5, UserControl.ScaleHeight - LBA.Height - 5
Case 4
LBA.Move 5, (UserControl.ScaleHeight - LBA.Height) - 5
Case 4
LBA.Move (UserControl.ScaleWidth - LBA.Width) - 5, 5
End Select
If IS_PIC = False Then
UserControl.Cls
LBA.Visible = True
Else
UserControl.Cls
LBA.Visible = False
UserControl.PaintPicture IMG.PICTURE, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End If
PDF.Move 0, UserControl.ScaleHeight, UserControl.ScaleWidth, UserControl.ScaleHeight
ICL.Move ScaleWidth - ICL.Width, 0
PUS.Visible = IS_PIC
End Sub
Public Sub SETPNG(File As String, X As Single, Y As Single)
PUS.Visible = True
On Error Resume Next '这句必须加
IS_PIC = True
'Dim I As Long
PUS.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
If IS_MUSIC = True Then PUS.Cls: PUS.PaintPicture IMG.PICTURE, 0, 0, PUS.ScaleWidth, PUS.ScaleHeight
Call PaintPng(File, PUS.hdc, X, Y)
PUS.Refresh
End Sub
Public Sub SETAUTHOR(PIC As PictureBox)
On Error Resume Next
IS_PIC = True
HASLINE = False
PUS.Cls
Set IMG.PICTURE = PIC.image
PUS.PaintPicture IMG.PICTURE, (PUS.ScaleWidth - IMG.Width) / 2, (PUS.ScaleHeight - IMG.Height) / 2, IMG.Width, IMG.Height, 0, 0      ', IMG.Width, IMG.Height
PUS.Refresh
End Sub
Sub SETFONT(FONT_NAME As String, Size As Long, IS_B As Boolean, SIZE_TIP As Long, B_B As Boolean)
LBA.FontName = FONT_NAME
LBA.FontSize = Size
LBA.FontBold = IS_B
LBC.FontSize = SIZE_TIP
LBC.FontBold = B_B
UserControl_Resize
End Sub
Sub SETPATH(Path As String)
On Error Resume Next
IS_PIC = True
FILEPATH.Refresh
FILEPATH.Path = Path
PUS.Visible = True
PUS.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
If FILEPATH.ListCount <> 0 Then TMPLAY.Enabled = True Else TMPLAY.Enabled = False
End Sub
Sub SET_STYLE(STY As Integer)
MY_STYLE = STY
UserControl_Resize
End Sub
'延时子过程
Private Sub Delay(ByVal t As Long)
Dim tm1 As Long, tm2 As Long  '变量定义
tm1 = timeGetTime  '获取时间
Do
tm2 = timeGetTime  '保存时间
If (tm2 - tm1) / 1000 > t Then Exit Do  '延时处理
DoEvents
Loop
End Sub

Sub SETEDIT(SHOWIT As Boolean)
K.Visible = SHOWIT
If IS_MUSIC = True Then K.PICTURE = Frmm.PIC(PLAYDSB).PICTURE
End Sub
Sub Refresh()
PUS.Refresh
UserControl.Refresh
End Sub
