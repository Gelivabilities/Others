VERSION 5.00
Begin VB.Form FRMTASK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H002DA6FF&
   BorderStyle     =   0  'None
   Caption         =   "TASK"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   Icon            =   "FRMTASK.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   Begin ICEE.ICEE_KEY ISU 
      Height          =   1575
      Left            =   1470
      TabIndex        =   5
      Top             =   15
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2778
   End
   Begin ICEE.ICEE_KEY ICC 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin VB.Timer TMS 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   15
   End
   Begin VB.PictureBox PTT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   15
      Width           =   1575
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在播放"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   960
      End
      Begin VB.Label LT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   960
         TabIndex        =   3
         Top             =   1320
         Width           =   450
      End
   End
   Begin VB.PictureBox PTASK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C28700&
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   1680
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   0
         ScaleHeight     =   79
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   84
         TabIndex        =   1
         Top             =   120
         Width           =   1260
         Begin VB.Image IP 
            Height          =   270
            Index           =   0
            Left            =   525
            Picture         =   "FRMTASK.frx":038A
            Top             =   450
            Width           =   255
         End
         Begin VB.Image IP 
            Height          =   270
            Index           =   1
            Left            =   495
            Picture         =   "FRMTASK.frx":06FA
            Top             =   450
            Width           =   255
         End
      End
      Begin VB.Image IP 
         Height          =   360
         Index           =   2
         Left            =   1320
         Picture         =   "FRMTASK.frx":0A74
         Top             =   555
         Width           =   360
      End
   End
   Begin VB.Image IA 
      Height          =   270
      Index           =   5
      Left            =   960
      Picture         =   "FRMTASK.frx":11DE
      Top             =   15
      Width           =   255
   End
   Begin VB.Image IA 
      Height          =   270
      Index           =   4
      Left            =   720
      Picture         =   "FRMTASK.frx":1558
      Top             =   15
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image IA 
      Height          =   360
      Index           =   3
      Left            =   720
      Picture         =   "FRMTASK.frx":18D2
      Top             =   15
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image IA 
      Height          =   360
      Index           =   2
      Left            =   240
      Picture         =   "FRMTASK.frx":203C
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image IA 
      Height          =   270
      Index           =   1
      Left            =   720
      Picture         =   "FRMTASK.frx":27A6
      Top             =   15
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image IA 
      Height          =   270
      Index           =   0
      Left            =   240
      Picture         =   "FRMTASK.frx":2B16
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FRMTASK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim pos As POINTAPI '定义这个变量是取得鼠标坐标
Private Const PI As Single = 3.14159265358979
Dim j As Double

Private Sub Form_Activate()
Call DRAWSINGER
If Me.BackColor = COLOR_NOR Then Exit Sub
Me.BackColor = COLOR_NOR
PTASK.BackColor = COLOR_NOR
Picture1.BackColor = COLOR_NOR
ICC.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ISU.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
End Sub

Private Sub Form_DblClick()
Call frmma.iCan
End Sub

Private Sub Form_Load()
'Me.BackColor = COLOR_NOR
'PTASK.BackColor = COLOR_NOR
'Call AttachForm(FRMTASK, 150, 30, True)
IS_MINI_MINI = GetInitEntry("SYSTEM", "MINI", False)
If IS_MINI_MINI = True Then
Me.Width = 1600
Else
Me.Width = 5100
End If
Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - GetTaskbarHeight
IS_MINI = True
TMS.Enabled = True
If frmma.Wm.playState = wmppsPlaying Then
IP(1).Visible = True
IP(0).Visible = False
Else
IP(1).Visible = False
IP(0).Visible = True
End If
Call oMagneticWnd.AddWindow(Me.hwnd)
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
'窗体总在最上
Picture1.AutoRedraw = True
Picture1.Scale (-10, 10)-(10, -10)
j = 1
ICC.SETTXT "···"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISU.Visible = True Then ISU.Visible = False

End Sub

Private Sub Form_Resize()
Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - GetTaskbarHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
IS_MINI = False
Unload FRMLIST
'Call DetachForm
End Sub

Private Sub ICC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If FRMLIST.Visible = False Then
Call FRMLIST.Show(vbModeless, Me)
If FRMTASK.Top > FRMLIST.Height Then FRMLIST.Move FRMTASK.Left, FRMTASK.Top - FRMLIST.Height, FRMTASK.Width Else FRMLIST.Move FRMTASK.Left, FRMTASK.Top + FRMTASK.Height, FRMTASK.Width
Else
Unload FRMLIST
End If
End Sub
Private Sub IP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
If IP(0).PICTURE <> IA(0).PICTURE Then IP(0).PICTURE = IA(0).PICTURE
Case 1
If IP(1).PICTURE <> IA(4).PICTURE Then IP(1).PICTURE = IA(4).PICTURE
Case 2
If IP(2).PICTURE <> IA(2).PICTURE Then IP(2).PICTURE = IA(2).PICTURE
End Select
End Sub

Private Sub IP_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
If IP(0).PICTURE <> IA(1).PICTURE Then IP(0).PICTURE = IA(1).PICTURE
Case 1
If IP(1).PICTURE <> IA(5).PICTURE Then IP(1).PICTURE = IA(5).PICTURE
Case 2
If IP(2).PICTURE <> IA(3).PICTURE Then IP(2).PICTURE = IA(3).PICTURE
End Select
End Sub

Private Sub IP_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
frmma.Wm.Controls.Play
IP(1).Visible = True
IP(0).Visible = False
Sleep 500
frmma.TMP.Enabled = True
Case 1
frmma.Wm.Controls.pause
IP(1).Visible = False
IP(0).Visible = True
Sleep 500
frmma.TMP.Enabled = False
Case 2
If LOLIPOP = 3 Or LOLIPOP = 1 Or LOLIPOP = 2 Then
Call frmma.NT(2)
ElseIf LOLIPOP = 0 Then
Call frmma.NT(3)
End If
End Select
End Sub

Private Sub ISU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IS_MINI_MINI = False Then
Me.Width = 1600
IS_MINI_MINI = True
Unload FRMLIST
Else
IS_MINI_MINI = False
Me.Width = 5100
End If
lRet = SetInitEntry("SYSTEM", "MINI", IS_MINI_MINI)

End Sub

Private Sub LA_DblClick()
Call frmma.iCan
End Sub

Private Sub LA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LB_DblClick()
Call frmma.iCan
End Sub

Private Sub LB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LT_DblClick()
Call frmma.iCan
End Sub

Private Sub LT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If frmma.Wm.playState = wmppsPlaying Then frmma.Wm.Controls.pause Else frmma.Wm.Controls.Play
End Sub

Private Sub PTASK_DblClick()
Call frmma.iCan
End Sub

Private Sub PTASK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PTASK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISU.Visible = True Then ISU.Visible = False
If IP(0).PICTURE <> IA(0).PICTURE Then IP(0).PICTURE = IA(0).PICTURE
If IP(1).PICTURE <> IA(4).PICTURE Then IP(1).PICTURE = IA(4).PICTURE
If IP(2).PICTURE <> IA(2).PICTURE Then IP(2).PICTURE = IA(2).PICTURE
End Sub

Sub DRAWSINGER()

If MMAIN.PathFileExists(frmma.SINGERLOGO) = 0 Then  '看看歌手头像文件是否存在
Frmm.PSINGER.PICTURE = Frmm.PIC(152).image  '不存在则使用默认头像
PTT.PaintPicture Frmm.PSINGER.image, 0, 0, 105, 105
Else
Call DrawPicture(PTT.hdc, frmma.SINGERLOGO, 0, 0, 105, 105)
End If
Debug.Print PTT.AutoRedraw
Call PaintPng(App.Path & "\SKIN\TIP.PNG", PTT.hdc, 0, 83)
PTT.Refresh
End Sub

Private Sub PTT_DblClick()
Call frmma.iCan
End Sub

Private Sub PTT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PTT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISU.Visible = False Then ISU.Visible = True
End Sub

Private Sub TMS_Timer()
On Error Resume Next
j = Picture1.ScaleWidth / frmma.Wm.currentMedia.duration * frmma.Wm.Controls.currentPosition
Picture1.Cls
Picture1.FillStyle = 0
Picture1.FillColor = vbWhite
Picture1.FOREColor = vbWhite
Picture1.Circle (0, 0), 8.5, , -PI / 500, -PI * j / 10
Call PaintPng(App.Path & "\SKIN\MINI_P.PNG", Picture1.hdc, 1, 0)
LT.Caption = frmma.Wm.currentMedia.durationString
Dim r As RECT, p As POINTAPI, L As Long, rtn As Long, H As Long, H1 As Long, r1 As Long '鼠标移出/移入透明值得改变
L = GetWindowRect(Me.hwnd, r)
L = GetCursorPos(p)
GetCursorPos pos
SX = IIf(pos.X < 0 Or pos.X > Screen.Width / 15, IIf(pos.X < 0, 0, Screen.Width / 15), pos.X)
SY = IIf(pos.Y < 0 Or pos.Y > Screen.Height / 15, IIf(pos.Y < 0, 0, Screen.Height / 15), pos.Y)
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then '移出界面
If IP(0).PICTURE <> IA(0).PICTURE Then IP(0).PICTURE = IA(0).PICTURE
If IP(1).PICTURE <> IA(4).PICTURE Then IP(1).PICTURE = IA(4).PICTURE
If IP(2).PICTURE <> IA(2).PICTURE Then IP(2).PICTURE = IA(2).PICTURE
If ISU.Visible = True Then ISU.Visible = False
If LA.Visible = True Then LA.Visible = False
Else
If LA.Visible = False Then LA.Visible = True
End If

Select Case frmma.Wm.playState
Case wmppsPaused
IP(0).Visible = True
IP(1).Visible = False
LA.Caption = "暂停播放"
Case wmppsPlaying
LA.Caption = "正在播放"
IP(1).Visible = True
IP(0).Visible = False
End Select

End Sub
