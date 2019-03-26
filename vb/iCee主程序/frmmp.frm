VERSION 5.00
Begin VB.Form frmmp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDC034&
   BorderStyle     =   0  'None
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   Icon            =   "frmmp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "菜单"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   45
      TabIndex        =   13
      Top             =   45
      Width           =   480
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "百度贴吧"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   12
      Left            =   480
      TabIndex        =   12
      Top             =   4830
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "我的云账户"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   11
      Left            =   480
      TabIndex        =   11
      Top             =   3720
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "屏幕键盘"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "涂鸦画板"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   9
      Left            =   480
      TabIndex        =   9
      Top             =   1200
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "屏幕放大镜"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "每日资讯"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "文件管理"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "图像处理"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   2760
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ICEE新鲜事"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   4080
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "百度空间"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   4
      Left            =   480
      TabIndex        =   3
      Top             =   4440
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "新建便签"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   5
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "退出"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   6
      Left            =   480
      TabIndex        =   1
      Top             =   5190
      Width           =   1680
   End
   Begin VB.Label LT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "检查更新"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   7
      Left            =   480
      TabIndex        =   0
      Top             =   3120
      Width           =   1680
   End
   Begin VB.Shape SB 
      BackColor       =   &H006623D6&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H005C6105&
      Height          =   450
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   32
      X2              =   128
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   32
      X2              =   128
      Y1              =   148
      Y2              =   148
   End
End
Attribute VB_Name = "frmmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private Sub Form_Activate()
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
Me.Cls
Me.BackColor = COLOR_NOR
SB.BackColor = COLOR_HIGH
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
End Sub

Private Sub Form_Load()
MakeTransparent Me.hwnd, 230
'RPC.ROUND_FORM Me, 12, 1, 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call frmma.MOVENOW
If SB.Visible = True Then SB.Visible = False
End Sub

Private Sub Form_Terminate()
Set frmmp = Nothing
End Sub
Private Sub LT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Hide
End Sub

Private Sub LT_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
SB.Top = LT(Index).Top - 8
SB.Visible = True
End Sub

Private Sub LT_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
Call Frmm.CHECKNET
If status.RasConnState <> &H2000 Then Exit Sub '没联网不许看
If Left(IEver, 1) >= 7 Then FRMNEWS.Show
Case 1
If frmma.Left > FRMEX.Width Then
FRMEX.Move frmma.Left - FRMEX.Width, frmma.Top
Else
FRMEX.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMEX.Show
Case 2
frmGraphic.Show
Case 3
Call Frmm.CHECKNET
If status.RasConnState <> &H2000 Then Call SHOWWRONG("木有检测到活动网络,请检查网络状况!", 2): Exit Sub '没联网不许看
If IS_FIRST_LOAD_ACT = True Then
Load FRMWEBACT
Else
Dim I As Integer
For I = 0 To frmma.IWG.Count - 1
frmma.IWG(I).SETCOLOR COLOR_NOR, COLOR_HIGH
frmma.IWG(I).SETIMG FRMWEBACT.IMD(I)
Next
frmma.PF(11).Visible = True
frmma.PF(11).ZOrder 0
Call frmma.RUNSAFE
End If
Case 4
ShellExecute 0&, vbNullString, "http://hi.baidu.com/new/iceeorgan", vbNullString, vbNullString, 0 '调用ie
Me.Hide
Case 5
frmma.LoadNote
Case 6
Unload frmma
Case 7
FRMUP.Show
Case 8
Call frmma.SHOWZOOM
Case 9
FRMBOARD.Show
Case 10
FRMKEYBOARD.Show
Case 11
If frmma.Winsock1.State <> 7 Then
Call SHOWWRONG("请先登录服务器!", 0)
Exit Sub
Else
If frmma.Left > FRMMYID.Width Then
FRMMYID.Move frmma.Left - FRMMYID.Width, frmma.Top
Else
FRMMYID.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMMYID.Show
End If
Case 12
ShellExecute 0&, vbNullString, "http://tieba.baidu.com/f?ie=utf-8&kw=icee", vbNullString, vbNullString, 0 '调用ie
End Select
End Sub
