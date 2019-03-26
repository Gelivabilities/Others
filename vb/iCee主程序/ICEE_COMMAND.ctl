VERSION 5.00
Begin VB.UserControl ICEE_COMMAND 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   ToolboxBitmap   =   "ICEE_COMMAND.ctx":0000
   Begin VB.Timer TMOUT 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer TMIN 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape SB 
      BackColor       =   &H0030F1F1&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0069CDDE&
      Height          =   75
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label ltit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
   Begin VB.Shape SBK 
      BackColor       =   &H00231C09&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "ICEE_COMMAND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'私有变量
Dim POS As POINTAPI '定义这个变量是取得鼠标坐标
Dim I As Long
Public SetSel As Boolean
'公有变量
Public AnyWhere As Integer
Public MYID As String
Public HASLINE As Boolean
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Sub ltit_Change()
UserControl_Resize
End Sub

Private Sub ltit_Click()
RaiseEvent Click
End Sub

Private Sub ltit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub ltit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub ltit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

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
Dim R As RECT, p As POINTAPI, L As Long
Dim Rtn As Long
L = GetWindowRect(UserControl.hwnd, R)
L = GetCursorPos(p)
If SetSel = True Then
If SBK.BackColor <> &H7A7417 Then SBK.BackColor = &H7A7417
'If ltit.ForeColor <> &H0& Then ltit.ForeColor = &H0&
If SB.Visible = False Then SB.Visible = True
Exit Sub
Else
If SB.Visible = True Then SB.Visible = False
'If ltit.ForeColor <> vbWhite Then ltit.ForeColor = vbWhite
End If
If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then '移出界面
TMIN.Enabled = False
SBK.BackColor = &H231C09
TMOUT.Enabled = False
I = 0
Else
If IS_AM = False Then
SBK.BackColor = &H7A7417
Else
If SBK.BackColor <> &H7A7417 Then TMIN.Enabled = True
End If
End If
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
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
Private Sub UserControl_Initialize()
On Error Resume Next
I = 0
AnyWhere = 0 '(0中间对齐，01左边对齐，2右边对齐)
HASLINE = True
TMOUT.Enabled = True
TMIN.Enabled = False
SBK.BackColor = &H231C09
ltit.FontName = "微软雅黑"
SetSel = False
End Sub

Private Sub UserControl_Resize()
SBK.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
SB.Move 1, UserControl.ScaleHeight - SB.Height - 1, UserControl.ScaleWidth - 2
Select Case AnyWhere
Case 0
ltit.Move (UserControl.ScaleWidth - ltit.Width) / 2, (UserControl.ScaleHeight - ltit.Height) / 2
Case 1
ltit.Move 5, (UserControl.ScaleHeight - ltit.Height) / 2
Case 2
ltit.Move UserControl.ScaleWidth - ltit.Width - 5, (UserControl.ScaleHeight - ltit.Height) / 2
End Select
End Sub

Public Sub SETTXT(TIT As String)
ltit.Caption = TIT
MYID = TIT
If HASLINE = True Then SBK.BorderStyle = 1 Else SBK.BorderStyle = 0
End Sub
Public Sub SETMESEL(SEL As Boolean)
SetSel = SEL
End Sub
