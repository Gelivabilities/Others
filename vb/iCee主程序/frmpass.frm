VERSION 5.00
Begin VB.Form frmpass 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00241D0A&
   BorderStyle     =   0  'None
   Caption         =   "安全验证"
   ClientHeight    =   3510
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5700
   ControlBox      =   0   'False
   Icon            =   "frmpass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5700
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PBK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00241D0A&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4850
      Picture         =   "frmpass.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   0
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4850
      Picture         =   "frmpass.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4850
      Picture         =   "frmpass.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin VB.PictureBox pctBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00241D0A&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   5280
      ScaleHeight     =   2160
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   1680
      Width           =   4275
   End
   Begin VB.PictureBox pctShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00241D0A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2370
      Left            =   360
      ScaleHeight     =   2370
      ScaleWidth      =   5160
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   5160
      Begin ICEE.ICEE_COMMAND ICA 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
      End
      Begin ICEE.ICEE_COMMAND ICB 
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin VB.Label LS 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请找出碗下的骰子"
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   180
         Index           =   0
         Left            =   1845
         TabIndex        =   2
         Top             =   75
         Width           =   1440
      End
   End
   Begin VB.PictureBox PICCHOOSE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00241D0A&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   12
      Top             =   720
      Width           =   5295
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   2250
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   120
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   3969
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   2250
         Index           =   1
         Left            =   2640
         TabIndex        =   14
         Top             =   120
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   3969
      End
   End
   Begin VB.PictureBox PICPASS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00241D0A&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   480
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtPASS 
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
         Height          =   210
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   12
         TabIndex        =   9
         Top             =   645
         Visible         =   0   'False
         Width           =   3615
      End
      Begin ICEE.ICEE_COMMAND ICD 
         Height          =   375
         Left            =   960
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1785
         Width           =   1335
         _extentx        =   4471
         _extenty        =   1085
      End
      Begin ICEE.ICEE_COMMAND ICN 
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1785
         Width           =   1455
         _extentx        =   4471
         _extenty        =   1085
      End
      Begin VB.Shape SB 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   495
         Left            =   360
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   7
      Left            =   5700
      Picture         =   "frmpass.frx":0636
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   6
      Left            =   4740
      Picture         =   "frmpass.frx":148C
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   5
      Left            =   3780
      Picture         =   "frmpass.frx":22E2
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   4
      Left            =   2760
      Picture         =   "frmpass.frx":3138
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   3
      Left            =   4260
      Picture         =   "frmpass.frx":3F8E
      Top             =   6525
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   2
      Left            =   3360
      Picture         =   "frmpass.frx":4DE4
      Top             =   6525
      Width           =   855
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   1
      Left            =   2460
      Picture         =   "frmpass.frx":5C3A
      Top             =   6525
      Width           =   855
   End
   Begin VB.Image imgRealObject 
      Height          =   240
      Left            =   3840
      Picture         =   "frmpass.frx":6A90
      Top             =   6105
      Width           =   240
   End
   Begin VB.Image imgObjectMask 
      Height          =   240
      Left            =   2400
      Picture         =   "frmpass.frx":6FD2
      Top             =   5760
      Width           =   240
   End
   Begin VB.Image imgRealCover 
      Height          =   645
      Index           =   0
      Left            =   1440
      Picture         =   "frmpass.frx":70D4
      Top             =   6525
      Width           =   855
   End
   Begin VB.Image imgMask 
      Height          =   645
      Left            =   1440
      Picture         =   "frmpass.frx":7F2A
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "frmpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const pwdChar = "·"
Dim IS_MV As Boolean
Dim Pwd As String, PwdLen As Long
Dim SelPos As Long, Insert As Integer
Dim super As Integer
Private Type CoverObject
  Top As Integer
  Left As Integer
  Possition As Integer
  dX As Integer
  Length As Integer
End Type
Private Cover(2) As CoverObject
Private OnBet As Boolean
Private WithEvents PSubClass As cSubclass
Attribute PSubClass.VB_VarHelpID = -1
Private CMD5 As New clsMD5
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(TXTPASS.Text) > 0 Then 我确定
End Sub

Private Sub Form_Load()
'建立界面
If App.PrevInstance Then End
Timer1.Enabled = False

'Dim WindowRegion As Long '定义抠图
'WindowRegion = getpic(Picture1) '开始抠图
'SetWindowRgn Me.hwnd, WindowRegion, True '抠图完成
'屏蔽密码框右键菜单
Set PSubClass = New cSubclass '无拖影
PSubClass.AddWindowMsgs Me.hwnd
oldproc = GetWindowLong(TXTPASS.hwnd, GWL_WNDPROC)
SetWindowLong TXTPASS.hwnd, GWL_WNDPROC, AddressOf TextWndProc
'建立透明
MakeTransparent Me.hwnd, 254
SetParent Me.hwnd, FindWindow(vbNullString, Me.Caption) '锁定到桌面
Call checkPass
Call SeekMe(Me)
ICN.HASLINE = False
ICD.HASLINE = False
ICA.HASLINE = False
ICB.HASLINE = False

ICN.SETTXT "取消"
ICD.SETTXT "确定"
ICA.SETTXT "开始"
ICB.SETTXT "取消"

Dim i As Integer
For i = 0 To IW.Count - 1
IW(i).HASLINE = False
IW(i).SETCOLOR &H241D0A, vbBlack
IW(i).SETTXTCOLOR vbWhite, vbWhite
Next
IW(0).SETTIP "输入密码"
IW(1).SETTIP "强制模式"
IW(0).SETPNG App.Path & "\SKIN\PIN.PNG", 40, 40
IW(1).SETPNG App.Path & "\SKIN\SIN.PNG", 40, 40

TXTPASS.Text = ""
Pwd = ""
TXTPASS.Visible = True
Timer1.Enabled = True
Call MoveWindow(frmpassbk.hwnd, Me.Left / Screen.TwipsPerPixelX - 9, Me.Top / Screen.TwipsPerPixelY - 9, 500, 360, True): frmpassbk.Show
IS_MV = True
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H201400, B

Me.Width = 5590
Me.Height = 3400

Me.Show
End Sub
Sub checkPass()
On Error Resume Next
Dim PS As String
PS = GetSetting("ICEE", "Main", "password", "")
super = GetSetting("ICEE", "Main", "SUPER", 0)
If super = 0 Then
PICPASS.Visible = True
TXTPASS.SetFocus
PICCHOOSE.Visible = False
Else
PICCHOOSE.Visible = True
PICPASS.Visible = False
Me.ScaleMode = 1
Me.Refresh
End If

If PS <> "" Then
Me.TXTPASS.Enabled = True
Else
Me.TXTPASS.Enabled = False
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpNow
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub Form_Terminate()
Set frmpass = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong TXTPASS.hwnd, GWL_WNDPROC, oldproc
Unload frmpassbk
End Sub
Private Sub ICA_CLICK()
Dim Possition1 As Integer
Dim Possition2 As Integer
Dim Possition3 As Integer
ICA.Visible = False
ICB.Visible = False
  GetRandomPossition Possition1, Possition2, Possition3
  Change Possition1, Possition2, Possition3, 50
  OpenCover 0
  OnBet = True
  DoEvents
  Play 0
  ICA.SETTXT "重新开始"
  ICA.Visible = True
  ICB.Visible = True
End Sub

Private Sub ICB_Click()
pctShow.Visible = False
PICCHOOSE.Visible = True
Me.Refresh
End Sub

Private Sub ICD_Click()
Call 我确定
End Sub

Private Sub ICN_Click()
Unload Me
End Sub

Private Sub IW_Click(Index As Integer)
On Error Resume Next
PICCHOOSE.Visible = False
PBK.Visible = True

Select Case Index
Case 0
PICPASS.Visible = True
TXTPASS.SetFocus
Case 1
pctShow.Visible = True
Call 超级模式
End Select

End Sub

Private Sub IW_MOUSEMOVE(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub PBK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = False Then
IS_MV = True
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub PBK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PICPASS.Visible = False
pctShow.Visible = False
PICCHOOSE.Visible = True
PBK.Visible = False
End Sub

Private Sub pctShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Y >= 0 And Y <= 800 Then UpNow
End Sub

Private Sub PICCHOOSE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PICCHOOSE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICPASS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PICPASS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim r As RECT, p As POINTAPI, L As Long
Dim rtn As Long
L = GetWindowRect(Me.hwnd, r)
L = GetCursorPos(p)
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then '移出界面
MOVENOW
Else '移入界面
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 254, LWA_ALPHA
End If
End Sub
Private Sub txtPASS_DblClick()
TXTPASS.SelLength = 0
TXTPASS.SelStart = Len(TXTPASS)
End Sub

Private Sub txtPass_GotFocus()
TXTPASS.SelLength = 0
TXTPASS.SelStart = Len(TXTPASS)
End Sub

Private Sub txtpass_KeyDown(KeyCode As Integer, Shift As Integer)
    SelPos = TXTPASS.SelStart
    PwdLen = Len(Pwd)
    Insert = 0

    If KeyCode = 46 Then
        If SelPos < PwdLen Then
            Pwd = Left(Pwd, SelPos) & Mid(Pwd, SelPos + 2)
            Call ChangeValue(TXTPASS)
            KeyCode = 0
        End If
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 22 Then KeyAscii = 0
If KeyAscii = 13 And Len(TXTPASS.Text) > 0 Then 我确定
    SelPos = TXTPASS.SelStart
    PwdLen = Len(Pwd)
    Insert = 0

    Select Case KeyAscii
    Case 8
        If SelPos > 0 Then
            Pwd = Left(Pwd, SelPos - 1) & Mid(Pwd, SelPos + 1)
            Insert = -1
        End If
    Case 32 To 126
        If (TXTPASS.MaxLength > 0 And PwdLen < TXTPASS.MaxLength) Or (TXTPASS.MaxLength = 0) Then
            Pwd = Left(Pwd, SelPos) & Chr(KeyAscii) & Mid(Pwd, SelPos + 1)
            Insert = 1
        End If
    End Select

    Call ChangeValue(TXTPASS)
    KeyAscii = 0

End Sub

Private Sub txtPASS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
TXTPASS.SelLength = 0
TXTPASS.SelStart = Len(TXTPASS)
End Sub

Private Sub txtPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TXTPASS.SelLength = 0
TXTPASS.SelStart = Len(TXTPASS)
End Sub
Private Sub x1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then
Unload Me
End If
End Sub
Private Sub 超级模式()
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H201400, B
Dim i As Integer

  Randomize
  For i = 0 To UBound(Cover)
    Cover(i).Left = 1000 + 1200 * i
    Cover(i).Top = 960
    If Rnd > 0.5 Then
      Cover(i).dX = -1
    Else
      Cover(i).dX = 1
    End If
    Cover(i).Length = Int(Rnd * 8) + 2
    Cover(i).Possition = Int(Rnd * 8)
  Next i
  
  OpenCover 0
End Sub
Private Sub 我确定()
Dim SMARTPASS As String
Dim Wrn As New FrmWrong
OL = CMD5.DigestStrToHexStr(Pwd)
SMARTPASS = GetSetting("ICEE", "Main", "password", "")
If OL = SMARTPASS Then
Timer1.Enabled = False
Me.Hide
Unload frmpassbk
frmma.Show
Else
Pwd = ""
TXTPASS.Text = ""
TXTPASS.SetFocus
With Wrn
.ts.Caption = "抱歉,您输入的密码有错误，无权限进入"
.DRAWINFOICO (0)
.Move Me.Left + (Me.Width / 2) - (.Width / 2) + 100, Me.Top + (Me.Height / 2) - (.Height / 2)
.Show vbModal
End With
Exit Sub
End If
End Sub
Sub UpNow()
Call ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub PSubClass_MsgCome(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, lng_hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    If bBefore Then
        If uMsg = WM_MOVE Then MoveWindow frmpassbk.hwnd, Me.Left / Screen.TwipsPerPixelX - 9, Me.Top / Screen.TwipsPerPixelY - 9, 500, 360, True
    End If
End Sub
Sub ChangeValue(TXT As TextBox)
    Dim i As Long, s As String, L As Long
    L = TXT.SelStart
    For i = 1 To Len(Pwd)
        s = s & pwdChar
    Next
    TXT.Text = s
    TXT.SelStart = L + Insert
End Sub


Private Sub Change(Index1 As Integer, Index2 As Integer, Index3 As Integer, ExDelay As Integer)
Dim dX As Integer
Dim dY As Integer
Dim X As Integer
Dim y1 As Integer
Dim y2 As Integer

Dim strPattern As String
Dim Length As Integer
Dim Divider As Integer
Dim MoveMode As Integer

Dim Cheat As Boolean
Dim CheatMode As Integer
Dim Index() As Integer

Dim i As Integer
Dim j As Integer
Dim K As Integer

  ReDim Index(UBound(Cover))
  If Cover(Index1).Left > Cover(Index2).Left Then
    Length = Cover(Index1).Left - Cover(Index2).Left
    dX = -1
  Else
    Length = Cover(Index2).Left - Cover(Index1).Left
    dX = 1
  End If
  
  If Rnd > 0.5 Then
    dY = -1
  Else
    dY = 1
  End If
  
  If Abs(Cover(Index1).Left - Cover(Index2).Left) > 2000 Then
    Divider = 400
    MoveMode = 0
  Else
    Divider = 240
    MoveMode = CInt(Rnd * 3)
  End If
  
  Cheat = (Rnd > 0.8)
  If Cheat Then
    CheatMode = CInt(Rnd * 3)
  End If
  
  For X = 1 To Length \ Divider
    If X <= (Length \ Divider) / 2 Then
      If X < 2 Then
        strPattern = strPattern & X
      Else
        strPattern = strPattern & 3
      End If
    Else
      If (((Length \ Divider)) - X) < 2 Then
        strPattern = strPattern & (((Length \ Divider)) - X)
      Else
        strPattern = strPattern & 3
      End If
    End If
  Next X
  
  y1 = Cover(Index1).Top
  y2 = Cover(Index1).Top
  
  For X = 1 To Length \ Divider
    Cover(Index1).Left = Cover(Index1).Left + dX * Divider
    Select Case MoveMode
      Case 0, 1, 3
        Cover(Index1).Top = y1 + dY * (Mid(strPattern, X, 1) * 150)
    End Select
    
    Cover(Index2).Left = Cover(Index2).Left - dX * Divider
    Select Case MoveMode
      Case 0, 2, 3
        Cover(Index2).Top = y1 - dY * (Mid(strPattern, X, 1) * 180)
    End Select
    
    For i = 0 To UBound(Cover)
      Index(i) = i
    Next i
    For i = 0 To 2
      For j = i + 1 To 2
        If Cover(Index(i)).Top > Cover(Index(j)).Top Then
          K = Index(i)
          Index(i) = Index(j)
          Index(j) = K
        End If
      Next j
    Next i
          
    pctBack.Cls
    For i = 0 To 2
      pctBack.PaintPicture imgMask.PICTURE, Cover(Index(i)).Left, Cover(Index(i)).Top, opcode:=vbSrcAnd
      If Index(i) <> Index3 Then
        If Cover(Index(i)).dX = 0 Then
          If Rnd > 0.5 Then
            Cover(Index(i)).dX = 1
          Else
            Cover(Index(i)).dX = -1
          End If
          Cover(Index(i)).Length = Int(Rnd * 8) + 2
        End If
        
        If Cover(Index(i)).dX = 1 Then
          Cover(Index(i)).Possition = (Cover(Index(i)).Possition + 1) Mod 8
        Else
          Cover(Index(i)).Possition = (Cover(Index(i)).Possition - 1 + 8) Mod 8
        End If
        Cover(Index(i)).Length = Cover(Index(i)).Length - 1
        
        If Cover(Index(i)).Length = 0 Then
          If Rnd > 0.5 Then
            Cover(Index(i)).dX = 1
          Else
            Cover(Index(i)).dX = -1
          End If
          Cover(Index(i)).Length = Int(Rnd * 8) + 2
        End If
      End If
      pctBack.PaintPicture imgRealCover(Cover(Index(i)).Possition).PICTURE, Cover(Index(i)).Left, Cover(Index(i)).Top, opcode:=vbSrcPaint
    Next i
    pctShow.PaintPicture pctBack.image, 0, 0, opcode:=vbSrcCopy
    Sleep ExDelay
    DoEvents
  Next X
End Sub

Private Sub OpenCover(Index As Integer)
Dim i As Integer
Dim Y As Integer
Dim y1 As Integer
Dim X As Integer
Dim strPattern As String

Dim Index1 As Integer
Dim Index2 As Integer

  Index1 = 0
  If Index1 = Index Then
    Index1 = 1
    Index2 = 2
  ElseIf Index = 1 Then
    Index2 = 2
  Else
    Index2 = 1
  End If
    
  strPattern = "12222210"
  Y = Cover(Index).Top
  
  For i = 1 To Len(strPattern)
    Cover(Index).Top = Y - (Mid(strPattern, i, 1) * 240)
    DoEvents
    
    pctBack.Cls
    pctBack.Circle (Cover(Index).Left + Me.imgMask.Width / 2 + Mid(strPattern, i, 1) * 30, Y + Me.imgMask.Height - 200 + Mid(strPattern, i, 1) * 30), imgMask.Width \ 2 + Mid(strPattern, i, 1) * 30, , , , 0.4
    If Index = 0 Then
      pctBack.PaintPicture imgObjectMask.PICTURE, Cover(Index).Left + 100, Y + 360, opcode:=vbSrcAnd
      pctBack.PaintPicture imgRealObject.PICTURE, Cover(Index).Left + 100, Y + 360, opcode:=vbSrcPaint
    End If
    
    pctBack.PaintPicture imgMask.PICTURE, Cover(Index1).Left, Cover(Index1).Top, opcode:=vbSrcAnd
    pctBack.PaintPicture imgRealCover(Cover(Index1).Possition).PICTURE, Cover(Index1).Left, Cover(Index1).Top, opcode:=vbSrcPaint
    
    pctBack.PaintPicture imgMask.PICTURE, Cover(Index2).Left, Cover(Index2).Top, opcode:=vbSrcAnd
    pctBack.PaintPicture imgRealCover(Cover(Index2).Possition).PICTURE, Cover(Index2).Left, Cover(Index2).Top, opcode:=vbSrcPaint
    
    pctBack.PaintPicture imgMask.PICTURE, Cover(Index).Left, Cover(Index).Top, opcode:=vbSrcAnd
    pctBack.PaintPicture imgRealCover(Cover(Index).Possition).PICTURE, Cover(Index).Left, Cover(Index).Top, opcode:=vbSrcPaint
    
    pctShow.PaintPicture pctBack.image, 0, 0, opcode:=vbSrcCopy
    Sleep 40
  Next i
End Sub

Private Sub Play(ByVal Level As Integer)
Dim Possition1 As Integer
Dim Possition2 As Integer
Dim i As Integer
Dim Possition3 As Integer
Dim Pattern As Single
Dim Delay As Integer
Dim strTemp As String
Dim ExtDelay As Integer

  DoEvents
  
    ExtDelay = 40
    Delay = 800 - ((Level + 1) * 150)

  For i = 1 To Level + 3
    GetRandomPossition Possition1, Possition2, Possition3
    
    Pattern = Rnd
    Change Possition1, Possition2, Possition3, ExtDelay
    If Pattern > 0.2 Then Change Possition3, Possition1, Possition2, ExtDelay
    If Pattern > 0.4 Then Change Possition2, Possition3, Possition1, ExtDelay
    If Pattern > 0.6 Then Change Possition1, Possition3, Possition2, ExtDelay
    If Pattern > 0.8 Then Change Possition3, Possition2, Possition1, ExtDelay
    If Pattern > 0.95 Then Change Possition2, Possition1, Possition3, ExtDelay
    Sleep Delay
  Next

End Sub
Private Sub pctShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim Found As Boolean
Dim ClickOnBowl As Boolean

 If Not OnBet Then Exit Sub
  
  OnBet = False
  ClickOnBowl = False
  Found = False
  
  For i = 0 To UBound(Cover)
    If X >= Cover(i).Left And X <= Cover(i).Left + 855 And Y >= Cover(i).Top And Y <= Cover(i).Top + 645 Then
      ClickOnBowl = True
      Found = (i = 0)
      OpenCover i
      
      Exit For
    End If
  Next i
  
  If ClickOnBowl Then
    If Not Found Then
      OpenCover 0
    Else
      Unload Me
      frmma.Show
    End If
  Else
    OnBet = True
  End If
  
End Sub

Private Sub GetRandomPossition(Possition1 As Integer, _
                               Possition2 As Integer, _
                               Possition3 As Integer)
Dim strTemp As String

  strTemp = "012"
  Possition1 = Int(Rnd * 3)
  Do
    Possition2 = Int(Rnd * 3)
  Loop Until Possition2 <> Possition1
  
  Mid(strTemp, InStr(strTemp, Possition1), 1) = " "
  Mid(strTemp, InStr(strTemp, Possition2), 1) = " "
  Possition3 = Val(strTemp)
End Sub


Sub MOVENOW()
X1.Visible = True
X2.Visible = False
X3.Visible = False
If IS_MV = True Then
IS_MV = False
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
End If
End Sub

