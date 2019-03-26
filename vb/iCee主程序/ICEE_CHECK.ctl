VERSION 5.00
Begin VB.UserControl ICEE_CHECK 
   BackColor       =   &H0094A63E&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   Begin VB.PictureBox SH_check 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   0
      Width           =   495
      Begin VB.Label LA 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   105
         TabIndex        =   2
         Top             =   120
         Width           =   270
      End
   End
   Begin VB.Timer TMIN 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3360
      Top             =   120
   End
   Begin VB.Timer TMOUT 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   120
   End
   Begin VB.Label ltit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
   Begin VB.Shape SBK 
      BackColor       =   &H007A7417&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H007A7417&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "ICEE_CHECK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'私有变量
Dim POS As POINTAPI '定义这个变量是取得鼠标坐标
Dim I As Long
'公有变量
Public Value As Long
Public Event Click()

Private Sub LA_Change()
SH_check_Resize
End Sub

Private Sub ltit_Change()
UserControl_Resize
End Sub

Private Sub ltit_Click()
UserControl_Click
RaiseEvent Click
End Sub
Private Sub SH_check_Resize()
LA.Move (SH_check.ScaleWidth - LA.Width) / 2, (SH_check.ScaleHeight - LA.Height) / 2
End Sub

Private Sub TMIN_Timer()
'移动时 RGB(23,116,122)
'初始时 RGB(9,28,35)
'相差值 RGB(14，88，87)
I = I + 1
SBK.BackColor = RGB(9 + I, 28 + 6 * I, 35 + 6 * I)
If I >= 14 Then
 TMIN.Enabled = False
 I = 0
 SBK.BackColor = &H7A7417
 TMOUT.Enabled = True
End If
End Sub

Private Sub TMOUT_Timer()
Dim R As RECT, P As POINTAPI, L As Long
Dim Rtn As Long
L = GetWindowRect(UserControl.hWnd, R)
L = GetCursorPos(P)
GetCursorPos POS
If P.X < R.Left Or P.X > R.Right Or P.Y < R.Top Or P.Y > R.Bottom Then '移出界面
TMIN.Enabled = False
SBK.BackColor = &H231C09
I = 0
Else
If IS_AM = True Then
If SBK.BackColor <> &H7A7417 Then TMIN.Enabled = True
Else
SBK.BackColor = &H7A7417
End If
End If
REFRESH_ME
End Sub

Private Sub UserControl_Click()
If Value = 0 Then Value = 1 Else Value = 0
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
I = 0
Value = 0
TMOUT.Enabled = True
TMIN.Enabled = False
On Error Resume Next
SBK.BackColor = &H231C09
ltit.FontName = "微软雅黑"
End Sub

Private Sub UserControl_Resize()
SBK.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
SH_check.Move 0, 0, 33, UserControl.ScaleHeight
ltit.Move (UserControl.ScaleWidth - ltit.Width + SH_check.Width) / 2, (UserControl.ScaleHeight - ltit.Height) / 2
End Sub

Public Sub SETTXT(Tit As String)
ltit.Caption = Tit
End Sub

Public Sub REFRESH_ME()
If Value = 1 Then
SH_check.BackColor = &H28D985
LA.Caption = "ON"
Else
SH_check.BackColor = &H1F1FE2
LA.Caption = "OFF"
End If
End Sub
