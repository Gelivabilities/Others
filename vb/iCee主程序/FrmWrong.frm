VERSION 5.00
Object = "{95C4D06B-0E76-491A-99C9-7BD3D4D1E34F}#1.0#0"; "Shadow.OCX"
Begin VB.Form FrmWrong 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "提示信息"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   ForeColor       =   &H00DBA13F&
   Icon            =   "FrmWrong.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   Begin VB.Timer TMFUN 
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin VB.PictureBox C1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4320
      Picture         =   "FrmWrong.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   15
      Width           =   750
   End
   Begin prjShadowCtl.ucShadow ucShadow1 
      Left            =   2790
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      Depth           =   20
      FadeTime        =   0
   End
   Begin VB.PictureBox C2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4320
      Picture         =   "FrmWrong.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4320
      Picture         =   "FrmWrong.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label ts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Let's Have Fun"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmWrong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ITIME As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Unload Me
End Sub
Sub DRAWINFOICO(WHICH As Integer)
Me.Cls
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H201400, B
Select Case WHICH
Case 0
Call PaintPng(App.Path & "\SKIN\MSG_WRN.PNG", Me.hdc, 8, 40)
Case 1
Call PaintPng(App.Path & "\SKIN\MSG_DONE.PNG", Me.hdc, 8, 40)
Case 2
Call PaintPng(App.Path & "\SKIN\MSG_INFO.PNG", Me.hdc, 8, 40)
End Select
Call PaintPng(App.Path & "\SKIN\W_T.PNG", Me.hdc, 8, 8)
Me.Refresh
End Sub
Private Sub Form_Load()
ICM.SETTXT "确定"
ITIME = 10
Me.BackColor = COLOR_NOR
ICM.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
MakeTransparent Me.hwnd, 240
Dim MYTOP, MYLEFT
MYLEFT = GetInitEntry("MsgBOX", "LEFT", (Screen.Width - Me.Width) / 2)
MYTOP = GetInitEntry("MsgBOX", "TOP", (Screen.Height - Me.Height) / 2)
Me.Move MYLEFT + 100, MYTOP + 100
lRet = SetInitEntry("MsgBOX", "LEFT", Me.Left)
lRet = SetInitEntry("MsgBOX", "TOP", Me.Top)
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
ts.Top = (Me.ScaleHeight - ts.Height) / 2 - 10
If Sound = 1 Then sndPlaySound App.Path + "\Sound\popo.wav", 1
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub Form_Unload(Cancel As Integer)
lRet = SetInitEntry("MsgBOX", "LEFT", Me.Left)
lRet = SetInitEntry("MsgBOX", "TOP", Me.Top)
End Sub

Private Sub ICM_Click()
Unload Me
End Sub

Private Sub LA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

C1.Visible = True
C2.Visible = False
C3.Visible = False

End Sub

Private Sub TMFUN_Timer()
ITIME = ITIME - 1
ICM.SETTXT "确定  (" & ITIME & ")"
If ITIME <= 0 Then ITIME = 0: TMFUN.Enabled = False: Unload Me
End Sub

Private Sub ts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
