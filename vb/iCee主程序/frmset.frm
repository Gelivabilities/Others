VERSION 5.00
Begin VB.Form frmset 
   BackColor       =   &H007A7417&
   BorderStyle     =   0  'None
   Caption         =   "功能设置"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox IFRAME 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00261700&
      BorderStyle     =   0  'None
      Height          =   9360
      Left            =   15
      ScaleHeight     =   624
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   0
      Top             =   15
      Width           =   5100
      Begin VB.Timer TMMOVE 
         Interval        =   100
         Left            =   240
         Top             =   360
      End
      Begin ICEE.ICEE_COMMAND ICM 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   8760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
      End
      Begin ICEE.ICEE_COMMAND ICM 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   3
         Top             =   8760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
      End
      Begin VB.PictureBox PD 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00261700&
         BorderStyle     =   0  'None
         Height          =   7935
         Left            =   120
         ScaleHeight     =   529
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   321
         TabIndex        =   1
         Top             =   600
         Width           =   4815
         Begin VB.PictureBox PO 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00261700&
            BorderStyle     =   0  'None
            Height          =   11640
            Left            =   0
            ScaleHeight     =   776
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   321
            TabIndex        =   2
            Top             =   -2880
            Width           =   4815
            Begin ICEE.ICHECK CHK_SOUND 
               Height          =   495
               Left            =   480
               TabIndex        =   6
               Top             =   0
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_ALOGIN 
               Height          =   495
               Left            =   480
               TabIndex        =   7
               Top             =   960
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_ATR 
               Height          =   495
               Left            =   480
               TabIndex        =   8
               Top             =   1440
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_NEW 
               Height          =   495
               Left            =   480
               TabIndex        =   9
               Top             =   1920
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_TRANS 
               Height          =   495
               Left            =   480
               TabIndex        =   10
               Top             =   2880
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_QK 
               Height          =   495
               Left            =   480
               TabIndex        =   11
               Top             =   3840
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_SELF 
               Height          =   495
               Left            =   480
               TabIndex        =   12
               Top             =   4800
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_SCR 
               Height          =   495
               Left            =   480
               TabIndex        =   13
               Top             =   4320
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_AD 
               Height          =   495
               Left            =   480
               TabIndex        =   14
               Top             =   5280
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_AP 
               Height          =   495
               Left            =   480
               TabIndex        =   15
               Top             =   2400
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_MT 
               Height          =   495
               Left            =   480
               TabIndex        =   16
               Top             =   3360
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_PASS 
               Height          =   495
               Left            =   480
               TabIndex        =   21
               Top             =   8160
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_TOP 
               Height          =   495
               Left            =   480
               TabIndex        =   22
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_AUTOSERCH 
               Height          =   495
               Left            =   480
               TabIndex        =   23
               Top             =   5760
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_AM 
               Height          =   495
               Left            =   480
               TabIndex        =   24
               Top             =   6240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_MINI 
               Height          =   495
               Left            =   480
               TabIndex        =   25
               Top             =   6720
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin ICEE.ICHECK CHK_S 
               Height          =   495
               Left            =   480
               TabIndex        =   26
               Top             =   7200
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   873
            End
            Begin VB.PictureBox FRAME1 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00261700&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   240
               ScaleHeight     =   41
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   305
               TabIndex        =   17
               Top             =   9240
               Width           =   4575
               Begin VB.TextBox txtpass1 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   1
                  EndProperty
                  Height          =   200
                  IMEMode         =   3  'DISABLE
                  Left            =   720
                  MaxLength       =   12
                  TabIndex        =   18
                  Text            =   "1111"
                  ToolTipText     =   "密码采用了MD5加密技术，请用户放心使用"
                  Top             =   150
                  Width           =   1215
               End
               Begin VB.TextBox txtpass2 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   1
                  EndProperty
                  Height          =   200
                  IMEMode         =   3  'DISABLE
                  Left            =   2640
                  MaxLength       =   12
                  TabIndex        =   19
                  Text            =   "1111"
                  ToolTipText     =   "密码采用了MD5加密技术，请用户放心使用"
                  Top             =   150
                  Width           =   1335
               End
               Begin VB.Label bt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "输入"
                  ForeColor       =   &H00000000&
                  Height          =   180
                  Index           =   0
                  Left            =   240
                  TabIndex        =   29
                  Top             =   150
                  Width           =   360
               End
               Begin VB.Label bt 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "确认"
                  ForeColor       =   &H00000000&
                  Height          =   180
                  Index           =   4
                  Left            =   2160
                  TabIndex        =   20
                  Top             =   150
                  Width           =   360
               End
               Begin VB.Shape SB 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H00808080&
                  Height          =   495
                  Index           =   0
                  Left            =   120
                  Top             =   0
                  Width           =   4095
               End
            End
            Begin ICEE.ICHECK CHK_H 
               Height          =   495
               Left            =   480
               TabIndex        =   27
               Top             =   7680
               Width           =   3255
               _extentx        =   5741
               _extenty        =   873
            End
            Begin ICEE.ICHECK CHK_SUPER 
               Height          =   495
               Left            =   480
               TabIndex        =   28
               Top             =   8640
               Width           =   3255
               _extentx        =   5741
               _extenty        =   873
            End
         End
         Begin ICEE.ucScrollbar SCRO 
            Height          =   7215
            Left            =   4560
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _extentx        =   450
            _extenty        =   12726
         End
      End
      Begin ICEE.ICEE_COMMAND ICM 
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   30
         Top             =   8760
         Width           =   1575
         _extentx        =   2778
         _extenty        =   661
      End
      Begin VB.Shape SB 
         BackColor       =   &H00241D0A&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00D16C29&
         BorderStyle     =   0  'Transparent
         Height          =   735
         Index           =   2
         Left            =   0
         Top             =   8640
         Width           =   5175
      End
      Begin VB.Image RunME 
         Height          =   300
         Left            =   4200
         ToolTipText     =   "返回主界面"
         Top             =   240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Image XME 
         Height          =   300
         Left            =   4680
         ToolTipText     =   "关闭"
         Top             =   240
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim passyn As Integer
Dim tR As Integer
Dim super As Integer
Dim AutoL As Integer
Private WithEvents PSubClass As cSubclass
Attribute PSubClass.VB_VarHelpID = -1
Const pwdChar = "·"
Const pwdChar1 = "·"
Dim Pwd As String, PwdLen As Long
Dim PwQd As String, PwQDLen As Long
Dim SelPos As Long, SELPOS1 As Long, Insert As Integer, INSERT1 As Integer
Sub ChangeValue1(TXT As TextBox)
On Error Resume Next
    Dim I As Long, s As String, L As Long
    L = TXT.SelStart
    For I = 1 To Len(PwQd)
        s = s & pwdChar1
    Next
    TXT.Text = s
    TXT.SelStart = L + INSERT1
End Sub
Sub ChangeValue(TXT As TextBox)
    Dim I As Long, s As String, L As Long
    L = TXT.SelStart
    For I = 1 To Len(Pwd)
        s = s & pwdChar
    Next
    TXT.Text = s
    TXT.SelStart = L + Insert
End Sub
Private Sub 完成了()
'这个不能删
Sound = CHK_SOUND.Value
Call SaveSetting("ICEE", "Main", "Tr", tR)
Call SaveSetting("ICEE", "Main", "passyn", passyn) '保存是否启动密码
Call SaveSetting("ICEE", "Main", "news", NEWS)
Call SaveSetting("ICEE", "Main", "SUPER", super)
Call SaveSetting("ICEE", "Main", "SOUND", Sound)
lRet = SetInitEntry("Player", "AutoPlay", CHK_AP.Value)
lRet = SetInitEntry("System", "weather", CHK_AD.Value)
lRet = SetInitEntry("ScreenSaver", "Opened", CHK_SCR.Value)
lRet = SetInitEntry("LocalSafe", "SelfHelp", CHK_SELF.Value)
lRet = SetInitEntry("System", "Quickkey", CHK_QK.Value)
lRet = SetInitEntry("SYSTEM", "transparent", CHK_MT.Value)
If CHK_S.Value = 1 Then AUTO_SINGER = True Else AUTO_SINGER = False
lRet = SetInitEntry("SYSTEM", "AUTO_S", AUTO_SINGER)
If CHK_AM.Value = 1 Then lRet = SetInitEntry("SYSTEM", "AM", True) Else lRet = SetInitEntry("SYSTEM", "AM", False)
lRet = SetInitEntry("PLAYER", "AUTOSERCHLRC", CHK_AUTOSERCH.Value)
If CHK_H.Value = 1 Then HAS_HEAD = True Else HAS_HEAD = False
lRet = SetInitEntry("SYSTEM", "HEAD_VIS", HAS_HEAD)
If CHK_TOP.Value = 0 Then
lRet = SetInitEntry("SYSTEM", "ONTOP", False)
ALWAYSONTOP = False
RESL = SetWindowPos(frmma.hwnd, 1, 0, 0, 0, 0, flags)
RESL = SetWindowPos(frmmabk.hwnd, 1, 0, 0, 0, 0, flags)
Else
lRet = SetInitEntry("SYSTEM", "ONTOP", True)
ALWAYSONTOP = True
RESL = SetWindowPos(frmma.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
RESL = SetWindowPos(frmmabk.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End If
QuickKey = CHK_QK.Value
If CHK_PASS.Value = 1 And Len(Trim(Pwd)) > 0 And Pwd = PwQd Then
OL = Col.DigestStrToHexStr(Pwd)
Call SaveSetting("ICEE", "Main", "password", OL) '保存密码的值
Call SaveSetting("ICEE", "Main", "pw", Pwd)
Unload Me
ElseIf CHK_PASS.Value = 0 Then
Call SaveSetting("ICEE", "Main", "password", "") '保存密码的值
Call SaveSetting("ICEE", "Main", "pw", "")
Call SaveSetting("ICEE", "Main", "passyn", 0)
Unload Me
Else
Me.Show
Me.txtpass2.SetFocus
Me.txtpass2.Text = ""
Call SHOWWRONG("对不起,两次输入的密码不同，请重新输入", 0)
Call SaveSetting("ICEE", "Main", "password", "") '保存密码的值
Call SaveSetting("ICEE", "Main", "pw", "")
Call SaveSetting("ICEE", "Main", "passyn", 0)
Exit Sub
End If
Call frmma.LoadSettings
End Sub
Private Sub bt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub CHK_ALOGIN_CLICK()
lRet = SetInitEntry("IM", "AutoLogin", CHK_ALOGIN.Value)
End Sub

Private Sub CHK_ATR_CLICK()
If CHK_ATR.Value = 1 Then SetAutoRun (True) Else SetAutoRun (False)
End Sub

Private Sub CHK_MINI_Click()
If CHK_MINI.Value = 0 Then CAN_MINI = False Else CAN_MINI = True
lRet = SetInitEntry("SYSTEM", "MINI_PLAYER", CAN_MINI)
End Sub

Private Sub CHK_NEW_CLICK()
NEWS = CHK_NEW.Value
End Sub

Private Sub CHK_PASS_CLICK()
passyn = CHK_PASS.Value
If passyn = 1 Then FRAME1.Visible = True Else FRAME1.Visible = False
End Sub

Private Sub CHK_SUPER_CLICK()
super = CHK_SUPER.Value
End Sub

Private Sub CHK_TRANS_CLICK()
tR = CHK_TRANS.Value
End Sub

Private Sub Form_Activate()
Call UnHook
H_DOS = 4
gHW = Me.hwnd '鼠标控件
Call Hook '唤醒鼠标滑轮API
PSubClass.AddWindowMsgs Me.hwnd
Call MoveWindow(FrmSetBK.hwnd, Me.Left / Screen.TwipsPerPixelX - 20, Me.Top / Screen.TwipsPerPixelY - 10, 380, 650, True)
End Sub
Private Sub Form_Load()
Set PSubClass = New cSubclass '无拖影模块事件
Dim Ro As Long, regV$ '检测附加启动项
Dim passw, PW As String  '密码的值
If frmma.Left > Me.Width Then
Me.Move frmma.Left - Me.Width, frmma.Top
Else
Me.Move frmma.Width + frmma.Left, frmma.Top
End If
PO.Height = 750
SCRO.Max = PO.Height - PD.ScaleHeight + 100
SCRO.Max = 150
SCRO.Value = 0
PO.Move 0, 0
PD.BackColor = &H261700
Call SeekMe(Me)
iFrame.Cls
IS_SET = True
Call PaintPng(App.Path & "\SKIN\TM.PNG", iFrame.hdc, 136, 12)

ICM(0).SETTXT "确定"
ICM(1).SETTXT "取消"
ICM(2).SETTXT "进入模式"

CHK_PASS.SETTXT "开启密码保护"
CHK_ALOGIN.SETTXT "登陆后自动登录"
CHK_SUPER.SETTXT "开启强制模式"
CHK_TRANS.SETTXT "为一部分窗体开启透明效果"
CHK_ATR.SETTXT "跟随Windows启动"
CHK_NEW.SETTXT "启动后显示资讯榜"
CHK_SOUND.SETTXT "开启音效提示"
CHK_TOP.SETTXT "窗体始终置顶"
CHK_AD.SETTXT "剪切板有可下载内容时自动下载"
CHK_SELF.SETTXT "自动修复桌面快捷方式"
CHK_QK.SETTXT "开启快捷键操作"
CHK_AP.SETTXT "启动后自动播放歌曲"
CHK_SCR.SETTXT "运行时禁止屏保"
CHK_MT.SETTXT "鼠标移出主窗体时透明"
CHK_AM.SETTXT "按钮开启动画效果"
CHK_AUTOSERCH.SETTXT "自动搜索歌词"
CHK_MINI.SETTXT "最小化后显示迷你播放器"
CHK_S.SETTXT "自动搜索歌手封面"
CHK_H.SETTXT "显示头像"
CHK_AUTOSERCH.Value = AUTOSERCH
If AUTO_SINGER = True Then CHK_S.Value = 1 Else CHK_S.Value = 0
If IS_AM = True Then CHK_AM.Value = 1 Else CHK_AM.Value = 0
CHK_MT.Value = GetInitEntry("SYSTEM", "transparent", 0)
CHK_AP.Value = GetInitEntry("Player", "AutoPlay", 0)
CHK_SCR.Value = GetInitEntry("ScreenSaver", "Opened", 1)
CHK_SELF.Value = GetInitEntry("LocalSafe", "SelfHelp", 1)
CHK_QK.Value = GetInitEntry("System", "Quickkey", 0)
CHK_AD.Value = GetInitEntry("System", "Weather", 0)
If HAS_HEAD = True Then CHK_H.Value = 1 Else CHK_H.Value = 0
If CAN_MINI = True Then CHK_MINI.Value = 1 Else CHK_MINI.Value = 0
If ALWAYSONTOP = True Then CHK_TOP.Value = 1 Else CHK_TOP.Value = 0
CHK_PASS.Value = GetSetting("ICEE", "Main", "passyn", 0) '检测是否使用了密码
PW = GetSetting("ICEE", "Main", "pw", "") '密码的值
AutoL = GetInitEntry("IM", "AutoLogin", 0) '是否自动登录
Sound = GetSetting("ICEE", "Main", "SOUND", 1)  '是否打开声音
tR = GetSetting("ICEE", "Main", "Tr", 0) '是否开启透明
NEWS = GetSetting("ICEE", "Main", "news", 1) '是否打开每日资讯
super = GetSetting("ICEE", "Main", "SUPER", 0) '是否开启超级模式
oldproc = GetWindowLong(txtpass1.hwnd, GWL_WNDPROC) '屏蔽文本框鼠标右键
SetWindowLong txtpass1.hwnd, GWL_WNDPROC, AddressOf TextWndProc
oldproc = GetWindowLong(txtpass2.hwnd, GWL_WNDPROC)
SetWindowLong txtpass2.hwnd, GWL_WNDPROC, AddressOf TextWndProc
CHK_ALOGIN.Value = AutoL
CHK_SOUND.Value = Sound
CHK_TRANS.Value = tR
If CHK_PASS.Value = 0 Then  'f=0时，则是没有使用密码
FRAME1.Visible = False
txtpass1.Text = ""
passyn = 0
Else
FRAME1.Visible = True
passyn = 1
End If
regV = GetStringValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "ICEE") '检测是否开机启动
If Trim(regV) = "" Then CHK_ATR.Value = 0 Else CHK_ATR.Value = 1
CHK_NEW.Value = NEWS
CHK_SUPER.Value = super
txtpass1.Text = PW
txtpass2.Text = PW
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call MoveWindow(FrmSetBK.hwnd, Me.Left / Screen.TwipsPerPixelX - 20, Me.Top / Screen.TwipsPerPixelY - 10, 380, 650, True)
Me.Show '我出现了
FrmSetBK.Show '使阴影展示
If ALWAYSONTOP = True Then RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(Me.hwnd, 1, 0, 0, 0, 0, flags)
MakeTransparent Me.hwnd, 254
End Sub
Private Sub Form_Unload(Cancel As Integer)
IS_SET = False
Call UnHook
Unload FrmSetBK
TMMOVE.Enabled = False
SetWindowLong txtpass1.hwnd, GWL_WNDPROC, oldproc
SetWindowLong txtpass2.hwnd, GWL_WNDPROC, oldproc
Call frmma.LoadSettings
Call frmma.DRAWFACE
End Sub
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub ICM_Click(Index As Integer)
Select Case Index
Case 0
Call 完成了
Case 1
Unload Me
Case 2
FRMRUN.Show
End Select
End Sub

Private Sub IFRAME_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub IFRAME_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Sub MOVENOW()
If XME.PICTURE <> Frmm.PIC(177).PICTURE Then XME.PICTURE = Frmm.PIC(177).PICTURE
If RunME.PICTURE <> Frmm.PIC(175).PICTURE Then RunME.PICTURE = Frmm.PIC(175).PICTURE
End Sub
Private Sub IFRAME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub

Private Sub RunME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If XME.PICTURE <> Frmm.PIC(177).PICTURE Then XME.PICTURE = Frmm.PIC(177).PICTURE
If RunME.PICTURE <> Frmm.PIC(176).PICTURE Then RunME.PICTURE = Frmm.PIC(176).PICTURE
End Sub

Private Sub RunME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Unload Me
End Sub

Private Sub SCRO_Change()
PO.Top = -SCRO.Value
End Sub

Private Sub SCRO_Scroll()
SCRO_Change
End Sub
Private Sub TMMOVE_Timer()
Dim r As RECT, p As POINTAPI, L As Long
Dim rtn As Long
L = GetWindowRect(Me.hwnd, r)
L = GetCursorPos(p)
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then Call MOVENOW
End Sub

Private Sub txtpass1_Change()
On Error Resume Next
Call ChangeValue(txtpass1)
End Sub

Private Sub txtpass1_DblClick()
txtpass1.SelLength = 0
txtpass1.SelStart = Len(txtpass1.Text)
End Sub

Private Sub txtpass1_GotFocus()
txtpass1.SelLength = 0
txtpass1.SelLength = Len(txtpass1)
End Sub

Private Sub txtpass1_KeyDown(KeyCode As Integer, Shift As Integer)
    SelPos = txtpass1.SelStart
    PwdLen = Len(Pwd)
    Insert = 0
    If KeyCode = 46 Then
        If SelPos < PwdLen Then
            Pwd = Left(Pwd, SelPos) & Mid(Pwd, SelPos + 2)
            Call ChangeValue(txtpass1)
            KeyCode = 0
        End If
    End If

End Sub

Private Sub txtpass1_KeyPress(KeyAscii As Integer)
If KeyAscii = 22 Then KeyAscii = 0
    SelPos = txtpass1.SelStart
    PwdLen = Len(Pwd)
    Insert = 0

    Select Case KeyAscii
    Case 8
        If SelPos > 0 Then
            Pwd = Left(Pwd, SelPos - 1) & Mid(Pwd, SelPos + 1)
            Insert = -1
        End If
    Case 32 To 126
        If (txtpass1.MaxLength > 0 And PwdLen < txtpass1.MaxLength) Or (txtpass1.MaxLength = 0) Then
            Pwd = Left(Pwd, SelPos) & Chr(KeyAscii) & Mid(Pwd, SelPos + 1)
            Insert = 1
        End If
    End Select

    Call ChangeValue(txtpass1)
    KeyAscii = 0
End Sub

Private Sub txtpass1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpass1.SelLength = 0
txtpass1.SelStart = Len(txtpass1.Text)
End Sub

Private Sub txtpass1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpass1.SelLength = 0
txtpass1.SelStart = Len(txtpass1.Text)
End Sub

Private Sub txtpass2_Change()
On Error Resume Next
Call ChangeValue1(txtpass2)
End Sub

Private Sub txtpass2_DblClick()
txtpass2.SelLength = 0
txtpass2.SelStart = Len(txtpass2.Text)
End Sub

Private Sub txtpass2_GotFocus()
txtpass2.SelLength = 0
txtpass2.SelStart = Len(txtpass2.Text)
End Sub

Private Sub txtpass2_KeyDown(KeyCode As Integer, Shift As Integer)
    SELPOS1 = txtpass2.SelStart
    PwdQLen = Len(PwQd)
    INSERT1 = 0
    If KeyCode = 46 Then
        If SELPOS1 < PwdQLen Then
            PwQd = Left(PwQd, SELPOS1) & Mid(PwQd, SELPOS1 + 2)
            Call ChangeValue1(txtpass2)
            KeyCode = 0
        End If
    End If
End Sub

Private Sub txtpass2_KeyPress(KeyAscii As Integer)
If KeyAscii = 22 Then KeyAscii = 0
    SELPOS1 = txtpass2.SelStart
    PwdQLen = Len(PwQd)
    INSERT1 = 0

    Select Case KeyAscii
    Case 8
        If SELPOS1 > 0 Then
            PwQd = Left(PwQd, SELPOS1 - 1) & Mid(PwQd, SELPOS1 + 1)
            INSERT1 = -1
        End If
    Case 32 To 126
        If (txtpass2.MaxLength > 0 And PwdQLen < txtpass2.MaxLength) Or (txtpass2.MaxLength = 0) Then
            PwQd = Left(PwQd, SELPOS1) & Chr(KeyAscii) & Mid(PwQd, SELPOS1 + 1)
            INSERT1 = 1
        End If
    End Select

    Call ChangeValue1(txtpass2)
    KeyAscii = 0

End Sub

Private Sub txtpass2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpass2.SelLength = 0
txtpass2.SelStart = Len(txtpass2.Text)
End Sub

Private Sub txtpass2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpass2.SelLength = 0
txtpass2.SelStart = Len(txtpass2.Text)
End Sub
Private Sub XME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If XME.PICTURE <> Frmm.PIC(178).PICTURE Then XME.PICTURE = Frmm.PIC(178).PICTURE
If RunME.PICTURE <> Frmm.PIC(175).PICTURE Then RunME.PICTURE = Frmm.PIC(175).PICTURE
End Sub

Private Sub XME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Unload Me
End Sub
Private Sub PSubClass_MsgCome(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, lng_hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
If bBefore Then
If uMsg = WM_MOVE Then MoveWindow FrmSetBK.hwnd, Me.Left / Screen.TwipsPerPixelX - 20, Me.Top / Screen.TwipsPerPixelY - 10, 380, 650, True
End If
End Sub
