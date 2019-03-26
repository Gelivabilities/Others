VERSION 5.00
Object = "{95C4D06B-0E76-491A-99C9-7BD3D4D1E34F}#1.0#0"; "Shadow.OCX"
Begin VB.Form FRMLL 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00241D0A&
   BorderStyle     =   0  'None
   Caption         =   "启动中"
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FRMLL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   1095
      Left            =   4245
      TabIndex        =   0
      Top             =   15
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   1931
   End
   Begin VB.Timer TMRRE 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   840
      Top             =   480
   End
   Begin prjShadowCtl.ucShadow ucShadow1 
      Left            =   240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Depth           =   5
      FadeTime        =   0
   End
   Begin VB.Image USELOGO 
      Height          =   1695
      Left            =   1800
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "FRMLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim LOADTIME As Long
Private SU As Integer
Private Sub Form_Initialize()
LOADTIME = 0
ALWAYSONTOP = GetInitEntry("SYSTEM", "ONTOP", False)
End Sub

Private Sub Form_Load()
On Error Resume Next
Call SeekMe(Me)
If App.PrevInstance Then End
LOGO = GetSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
If Sound = 1 Then sndPlaySound App.Path + "\Sound\Load.wav", 1        '播放登陆声音
If Len(Trim(LOGO)) > 0 And PathFileExists(LOGO) <> 0 Then
USELOGO.PICTURE = LoadPicture(LOGO)

Else
Call SaveSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
USELOGO.PICTURE = LoadPicture(App.Path + "\Skin\DefaultHead.Bmp")
End If

RESL = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags) '置顶
Dim Buffer As String, LST As String 'declare the needed variables
Buffer = Space(MAX_PATH)
Rtn = GetSystemDirectory(Buffer, Len(Buffer)) 'get the path
Rtn = GetWindowsDirectory(Buffer, Len(Buffer)) 'get the path
WinSysPath = Left(Buffer, Rtn)                'parse the path into the global string
WinPath = Left(Buffer, Rtn)                    'parse the path to the global string
lRet = SetInitEntry("OS", "WINPATH", WinPath)
lRet = SetInitEntry("OS", "WINSYSPATH", WinSysPath)
LST = GetInitEntry("Time", "Last End", Now)
Me.PaintPicture USELOGO.PICTURE, 8, 5, 65, 65
ICM.HASLINE = False
ICM.SETTXT "取消启动"
Call PaintPng(App.Path & "\SKIN\LOGO_65.png", Me.hdc, 8, 5)
Call PaintPng(App.Path & "\SKIN\URL.png", Me.hdc, 48, 0)
Me.CurrentY = 20
Me.CurrentX = 90
Me.Print "上次运行:" & LST
Me.CurrentY = 40
Me.CurrentX = 110
Me.Print "Powered By Mirror"
TMRRE.Enabled = True
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H808080, B
Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - GetTaskbarHeight

Me.Show
Load FRMWEATHER
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call CMV(Me)
End Sub
Private Sub Form_Terminate()
Set FRMLL = Nothing
End Sub

Private Sub LBC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Unload Me
End
End Sub

Private Sub ICM_Click()
Unload Me
End
End Sub

Private Sub TMRRE_Timer()
On Error Resume Next
LOADTIME = LOADTIME + 4
If LOADTIME = 100 Then
TMRRE.Enabled = False
Me.Hide
Frmm.WB.Navigate "http://hi.baidu.com/iceeorgan/item/96d45007a86c1acbff240dfa"
Load frmma
End If
End Sub

