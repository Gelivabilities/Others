VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.UserControl IMUSIC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00606015&
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ToolboxBitmap   =   "IMUSIC.ctx":0000
   Begin VB.PictureBox PP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
      Begin VB.Shape SB 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   75
         Left            =   0
         Top             =   720
         Width           =   615
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未知歌手"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   405
         Width           =   720
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未知歌名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   840
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   1200
   End
   Begin WMPLibCtl.WindowsMediaPlayer WM 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   450
   End
End
Attribute VB_Name = "IMUSIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MUSIC_URL As String
Public IS_PLAY As Boolean
Public HASPIC As Boolean, D_COLOR As Long, N_COLOR As Long
Event MouseDown()


Private Sub LA_Change()
If LA.Caption = "" Then PP.Height = 5 Else PP.Height = 55
UserControl_Resize
End Sub

Private Sub PP_Resize()
SB.Move 0, PP.ScaleHeight - SB.Height, 0, 5
End Sub

Private Sub tmrTimer_Timer()
On Error Resume Next
SB.Width = UserControl.ScaleWidth / WM.currentMedia.duration * WM.Controls.currentPosition
'If IS_PLAY = True Then WM.Controls.Play Else WM.Controls.pause
'If IS_PLAY = False Then
'WM.URL = MUSIC_URL
'WM.Controls.Play
'End If
Select Case WM.playState
Case wmppsPaused, wmppsStopped
IS_PLAY = False
Case wmppsPlaying
IS_PLAY = True
End Select
End Sub

Private Sub UserControl_Initialize()
LA.Caption = ""
LB.Caption = ""
WM.settings.volume = 100
End Sub
Sub PLAY_IT()
    On Error Resume Next
    If ERR.Number <> 0 Then
        Exit Sub
    End If
    WM.URL = MUSIC_URL
    tmrTimer.Enabled = True
    WM.Controls.Play
    IS_PLAY = True
UserControl.Cls
Call PaintPng(App.Path & "\SKIN\PA_N.PNG", UserControl.hdc, (UserControl.ScaleWidth - 84) / 2, (UserControl.ScaleHeight - 84) / 2)
End Sub
Sub STOP_IT()
WM.Controls.Stop
IS_PLAY = False
UserControl_Resize
End Sub
Sub PAUSE_IT()
WM.Controls.pause
UserControl.Cls
IS_PLAY = False
UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown
If Button <> 1 Then Exit Sub
If MUSIC_URL = "" Then Exit Sub
If IS_PLAY = True Then
PAUSE_IT
Else
PLAY_IT
End If
SB.Width = 0
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
PP.Move 0, UserControl.ScaleHeight - PP.Height, UserControl.ScaleWidth
UserControl.Cls
If HASPIC = True Then UserControl.PaintPicture PIC.image, 0, 0, ScaleWidth, ScaleHeight
If IS_PLAY = False Then
Call PaintPng(App.Path & "\SKIN\P_N.PNG", UserControl.hdc, (UserControl.ScaleWidth - 84) / 2, (UserControl.ScaleHeight - 84 - PP.Height) / 2)
Else
Call PaintPng(App.Path & "\SKIN\PA_N.PNG", UserControl.hdc, (UserControl.ScaleWidth - 84) / 2, (UserControl.ScaleHeight - 84 - PP.Height) / 2)
End If
End Sub

Sub SETPIC(PICTURE As String)
HASPIC = True
PIC.PICTURE = LoadPicture(PICTURE)
UserControl_Resize
End Sub
Sub SETIMG(IMG As PictureBox)
HASPIC = True
PIC.PICTURE = IMG.PICTURE
UserControl_Resize
End Sub
Sub SETTXT(ltit As String, AUTH As String)
LA.Caption = ltit
LB.Caption = AUTH
End Sub
Sub SETCOLOR(Color As Long, COLOR_B As Long)
UserControl.BackColor = Color
N_COLOR = Color
PP.BackColor = COLOR_B
PIC.BackColor = Color
UserControl_Resize
D_COLOR = COLOR_B
End Sub
Sub Cls()
UserControl.Cls
UserControl.BackColor = N_COLOR
PIC.BackColor = D_COLOR
Set PIC.PICTURE = Nothing
UserControl_Resize
End Sub

