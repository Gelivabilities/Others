VERSION 5.00
Begin VB.Form FRMTOOL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "桌面歌词"
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   47
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   450
      Index           =   0
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   794
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer TMOUT 
      Interval        =   5000
      Left            =   -240
      Top             =   480
   End
   Begin VB.PictureBox PC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   1560
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox PC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   1200
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox PC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00DB59D8&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   840
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox PC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H002EBC7C&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   480
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox PC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AA7402&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   450
      Index           =   1
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   794
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   450
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   794
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   450
      Index           =   3
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   794
   End
   Begin VB.Shape SB 
      BackColor       =   &H0030F1F1&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0030F1F1&
      Height          =   45
      Left            =   1560
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "FRMTOOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim POS As POINTAPI '定义这个变量是取得鼠标坐标
Public cDeskLrc As New clsDeskLrc
Private Sub Form_Load()
On Error Resume Next
Dim SIT As Integer
SIT = GetInitEntry("PLAYER", "LRCSHOW_COLOR", 0)
Call SETCOLOR(SIT + 1)
cDeskLrc.ShowText "ICEE音乐,音乐您的生活"
cDeskLrc.FontName = "微软雅黑"
cDeskLrc.FontBold = True
cDeskLrc.Karaoke = True
cDeskLrc.FontSize = 32
cDeskLrc.ReDraw
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H808080, B
RESL = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
ICM(0).HASLINE = False
ICM(1).HASLINE = False
ICM(2).HASLINE = False
ICM(3).HASLINE = False
ICM(1).SETTXT "×"
ICM(0).SETTXT "编辑"
ICM(3).SETTXT "←"
ICM(2).SETTXT "→"
SB.Move PC((SIT) + 1).Left
Call MakeTransparent(Me.hWnd, 254)
End Sub

Private Sub ICM_Click(INDEX As Integer)
Select Case INDEX
Case 0
Call FRMLRC.L_EDIT
Case 1
Set cDeskLrc = Nothing
Unload FRMSHOW
Unload Me
Case 3
Call frmma.NT(1)
Case 2
If LOLIPOP = 3 Or LOLIPOP = 1 Or LOLIPOP = 2 Then
Call frmma.NT(2)
ElseIf LOLIPOP = 0 Then
Call frmma.NT(3)
End If
End Select
End Sub

Private Sub PC_MouseDown(INDEX As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
SETCOLOR INDEX + 1
cDeskLrc.ReDraw
SB.Move PC(INDEX).Left
lRet = SetInitEntry("PLAYER", "LRCSHOW_COLOR", INDEX + 1)
End Sub

Private Sub Timer1_Timer()
If frmma.Wm.playState = wmppsPlaying Then cDeskLrc.SeekLrc frmma.Wm.Controls.currentPosition, False
End Sub

Private Sub TMOUT_Timer()
Dim R As RECT, p As POINTAPI, L As Long, Rtn As Long, H As Long, H1 As Long, r1 As Long '鼠标移出/移入透明值得改变
L = GetWindowRect(Me.hWnd, R)
L = GetCursorPos(p)
GetCursorPos POS
If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then Me.Hide
End Sub
Sub SETCOLOR(ByVal Mode As Integer)
    Select Case Mode
        Case 1          '蓝色
            cDeskLrc.BackColor1 = &HFF013C8F
            cDeskLrc.BackColor2 = &HFF0198D4
            cDeskLrc.ForeColor1 = &HFFBCF9FC
            cDeskLrc.ForeColor2 = &HFF67F0FC
            cDeskLrc.LineColor = &H30000000
        Case 2          '绿色
            cDeskLrc.BackColor1 = &HFF87F321
            cDeskLrc.BackColor2 = &HFF0E6700
            cDeskLrc.ForeColor1 = &HFFDCFEAE
            cDeskLrc.ForeColor2 = &HFFE4FE04
            cDeskLrc.LineColor = &H30000000
        Case 3          '红色
            cDeskLrc.BackColor1 = &HFFFECEFC
            cDeskLrc.BackColor2 = &HFFE144CD
            cDeskLrc.ForeColor1 = &HFFFEFE65
            cDeskLrc.LineColor = &H30000000
        Case 4          '白色
            cDeskLrc.BackColor1 = &HFFFBFBFA
            cDeskLrc.BackColor2 = &HFFCBCBCB
            cDeskLrc.ForeColor1 = &HFF62DDFF
            cDeskLrc.ForeColor2 = &HFF229CFE
            cDeskLrc.LineColor = &H30000000
        Case 5          '黄色
            cDeskLrc.BackColor1 = &HFFFE7A00
            cDeskLrc.BackColor2 = &HFFFF0000
            cDeskLrc.ForeColor1 = &HFFFFFF6E
            cDeskLrc.ForeColor2 = &HFFFEA10F
            cDeskLrc.LineColor = &H30000000
    End Select
End Sub

