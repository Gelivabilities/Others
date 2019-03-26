VERSION 5.00
Begin VB.UserControl ICHECK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00261700&
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   ToolboxBitmap   =   "ICHECK.ctx":0000
   Begin VB.Timer TMOUT 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3000
      Top             =   480
   End
   Begin VB.Image O2 
      Height          =   720
      Left            =   5040
      Picture         =   "ICHECK.ctx":0312
      Top             =   1080
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image O1 
      Height          =   720
      Left            =   3720
      Picture         =   "ICHECK.ctx":2F96
      Top             =   1080
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image C4 
      Height          =   240
      Left            =   7920
      Picture         =   "ICHECK.ctx":5C1A
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image C3 
      Height          =   240
      Left            =   7680
      Picture         =   "ICHECK.ctx":5FA4
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image C2 
      Height          =   240
      Left            =   7440
      Picture         =   "ICHECK.ctx":632E
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image C1 
      Height          =   240
      Left            =   7200
      Picture         =   "ICHECK.ctx":66B8
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image IA4 
      Height          =   450
      Left            =   5640
      Picture         =   "ICHECK.ctx":6A42
      Top             =   360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image IA3 
      Height          =   450
      Left            =   5040
      Picture         =   "ICHECK.ctx":6E92
      Top             =   360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image IA2 
      Height          =   450
      Left            =   4440
      Picture         =   "ICHECK.ctx":730D
      Top             =   360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image IA1 
      Height          =   450
      Left            =   3840
      Picture         =   "ICHECK.ctx":76BC
      Top             =   360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image IMV 
      Height          =   390
      Left            =   120
      Top             =   120
      Width           =   405
   End
   Begin VB.Label LB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
   Begin VB.Image sh1 
      Height          =   390
      Left            =   480
      Picture         =   "ICHECK.ctx":7A5D
      Top             =   600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image sh4 
      Height          =   390
      Left            =   1440
      Picture         =   "ICHECK.ctx":A86C
      Top             =   600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image sh3 
      Height          =   390
      Left            =   960
      Picture         =   "ICHECK.ctx":D858
      Top             =   600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image sh2 
      Height          =   390
      Left            =   120
      Picture         =   "ICHECK.ctx":108BC
      Top             =   600
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "ICHECK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Public Value As Integer
Public Event Click()
Public M_STYLE As Integer
Dim POS As POINTAPI '定义这个变量是取得鼠标坐标

Private Sub IMV_Click()
UserControl_Click
RaiseEvent Click

End Sub

Private Sub IMV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TMOUT.Enabled = True
End Sub

Private Sub LB_Click()
UserControl_Click
RaiseEvent Click

End Sub

Private Sub LB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TMOUT.Enabled = True
End Sub

Private Sub TMOUT_Timer()
Dim r As RECT, p As POINTAPI, L As Long
Dim Rtn As Long
L = GetWindowRect(UserControl.hWnd, r)
L = GetCursorPos(p)
GetCursorPos POS
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then '移出界面
'If LB.ForeColor <> &HC0C0C0 Then LB.ForeColor = &HC0C0C0
Select Case M_STYLE
Case 0
If Value = 0 Then
If IMV.PICTURE <> sh1.PICTURE Then IMV.PICTURE = sh1.PICTURE
Else
If IMV.PICTURE <> sh3.PICTURE Then IMV.PICTURE = sh3.PICTURE
End If
Case 1
If Value = 0 Then
If IMV.PICTURE <> IA1.PICTURE Then IMV.PICTURE = IA1.PICTURE
Else
If IMV.PICTURE <> IA3.PICTURE Then IMV.PICTURE = IA3.PICTURE
End If
Case 2
If Value = 0 Then
If IMV.PICTURE <> C1.PICTURE Then IMV.PICTURE = C1.PICTURE
Else
If IMV.PICTURE <> C3.PICTURE Then IMV.PICTURE = C3.PICTURE
End If

Case 3
If Value = 0 Then
If IMV.PICTURE <> O2.PICTURE Then IMV.PICTURE = O2.PICTURE
Else
If IMV.PICTURE <> O1.PICTURE Then IMV.PICTURE = O1.PICTURE
End If
End Select
TMOUT.Enabled = False
Else
'If LB.ForeColor <> vbWhite Then LB.ForeColor = vbWhite
Select Case M_STYLE
Case 0
If Value = 0 Then
If IMV.PICTURE <> sh2.PICTURE Then IMV.PICTURE = sh2.PICTURE
Else
If IMV.PICTURE <> sh4.PICTURE Then IMV.PICTURE = sh4.PICTURE
End If
Case 1
If Value = 0 Then
If IMV.PICTURE <> IA2.PICTURE Then IMV.PICTURE = IA2.PICTURE
Else
If IMV.PICTURE <> IA4.PICTURE Then IMV.PICTURE = IA4.PICTURE
End If
Case 2
If Value = 0 Then
If IMV.PICTURE <> C2.PICTURE Then IMV.PICTURE = C2.PICTURE
Else
If IMV.PICTURE <> C4.PICTURE Then IMV.PICTURE = C4.PICTURE
End If
Case 3
If Value = 0 Then
If IMV.PICTURE <> O2.PICTURE Then IMV.PICTURE = O2.PICTURE
Else
If IMV.PICTURE <> O1.PICTURE Then IMV.PICTURE = O1.PICTURE
End If
End Select
End If

End Sub

Private Sub UserControl_Click()
If Value = 0 Then Value = 1 Else Value = 0
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
LB.FontName = "微软雅黑"
Value = 0
M_STYLE = 0
TMOUT.Enabled = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TMOUT.Enabled = True
End Sub

Private Sub UserControl_Resize()
Select Case M_STYLE
Case 0
IMV.Move 5, (UserControl.ScaleHeight - sh1.Height) / 2
LB.Move IMV.Left + IMV.Width + 5, (UserControl.ScaleHeight - LB.Height) / 2
Case 1
IMV.Move 5, (UserControl.ScaleHeight - IA1.Height) / 2
LB.Move IMV.Left + IMV.Width + 5, (UserControl.ScaleHeight - LB.Height) / 2
Case 2
IMV.Move 5, (UserControl.ScaleHeight - C1.Height) / 2
LB.Move IMV.Left + IMV.Width + 5, (UserControl.ScaleHeight - LB.Height) / 2
Case 3
If UserControl.BackColor <> vbWhite Then UserControl.BackColor = vbWhite
If LB.ForeColor <> vbBlack Then LB.ForeColor = vbBlack
UserControl.Cls
UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), &H808080, B

LB.Move 10, (UserControl.ScaleHeight - LB.Height) / 2
IMV.Move UserControl.ScaleWidth - O1.Width - 10, (UserControl.ScaleHeight - O1.Height) / 2

Debug.Print LB.Top; LB.Left
End Select

TMOUT.Enabled = True
End Sub
Public Sub SETTXT(TXT As String)
If Len(TXT) > 35 Then LB.Caption = Left(TXT, 20) & "..." Else LB.Caption = TXT
UserControl_Resize
End Sub
Sub SETCOLOR(BKCOLOR As Long, CAP_COLOR As Long)
If MY_STYLE = 3 Then Exit Sub
LB.ForeColor = CAP_COLOR
UserControl.BackColor = BKCOLOR
UserControl_Resize
End Sub
