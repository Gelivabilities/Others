VERSION 5.00
Begin VB.UserControl ICEE_PIC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   ToolboxBitmap   =   "ICEE_PIC.ctx":0000
   Begin VB.PictureBox PT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   5280
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Timer TMIN 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6000
      Top             =   240
   End
   Begin VB.Timer TMOUT 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5520
      Top             =   240
   End
   Begin VB.PictureBox PBS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H007A7417&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.PictureBox PTOOL 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H005BB645&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   3
         Top             =   2640
         Visible         =   0   'False
         Width           =   3015
         Begin VB.Image ITYPE 
            Height          =   240
            Left            =   2400
            Top             =   75
            Width           =   480
         End
         Begin VB.Label LBSIZE 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "300×300"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Width           =   720
         End
      End
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   7
      Left            =   6120
      Picture         =   "ICEE_PIC.ctx":0312
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   6
      Left            =   6120
      Picture         =   "ICEE_PIC.ctx":099C
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   5
      Left            =   5520
      Picture         =   "ICEE_PIC.ctx":1026
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   4
      Left            =   4920
      Picture         =   "ICEE_PIC.ctx":16B0
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   3
      Left            =   4320
      Picture         =   "ICEE_PIC.ctx":1D3A
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   2
      Left            =   5520
      Picture         =   "ICEE_PIC.ctx":23C4
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   1
      Left            =   4920
      Picture         =   "ICEE_PIC.ctx":2A4E
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   0
      Left            =   4320
      Picture         =   "ICEE_PIC.ctx":30D8
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LBLNAME 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   630
   End
   Begin VB.Shape SBK 
      BackColor       =   &H007A7417&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "ICEE_PIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'私有变量
Dim POS As POINTAPI '定义这个变量是取得鼠标坐标
Dim I As Long
Dim cx As Long, cy As Long
'公有变量
Public AnyWhere As Integer
Public Auto_Size As Boolean, HAS_TXT As Boolean
Public BOARDSTYLE As Integer
Public MYID As String, P_SIZE As String
Public SHOWTOOL As Boolean
Public Event Click()
Public Event DBLCLICK()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Sub LBLNAME_Change()
Call SETICO
End Sub

Private Sub LBLNAME_Click()
RaiseEvent Click
End Sub

Private Sub LBLNAME_DblClick()
RaiseEvent DBLCLICK
End Sub

Private Sub LBLNAME_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBLNAME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
TMOUT.Enabled = True
End Sub

Private Sub LBLNAME_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBSIZE_Change()
P_SIZE = LBSIZE.Caption
End Sub

Private Sub PBS_Click()
RaiseEvent Click
End Sub

Private Sub PBS_DblClick()
RaiseEvent DBLCLICK
End Sub


Private Sub PBS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub PBS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
    If SHOWTOOL = False Then Exit Sub
    If PTOOL.Visible = False Then PTOOL.Visible = True
    TMOUT.Enabled = True
End Sub

Private Sub PBS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub PBS_Resize()
PTOOL.Move 0, PBS.ScaleHeight - PTOOL.Height, PBS.ScaleWidth
End Sub

Private Sub PTOOL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TMOUT.Enabled = True
End Sub

Private Sub PTOOL_Resize()
ITYPE.Move PTOOL.ScaleWidth - ITYPE.Width - 5, 5
End Sub

Private Sub TMIN_Timer()
'移动时 RGB(23,116,122)
'初始时 RGB(9,28,35)
'相差值 RGB(14，88，87)
I = I + 5
SBK.BackColor = RGB(9 + I, 28 + 6 * I, 35 + 6 * I)
If I >= 14 Then
 TMIN.Enabled = False
 I = 0
 SBK.BackColor = &H7A7417
 TMOUT.Enabled = True
End If
End Sub

Private Sub TMOUT_Timer()
Dim r As RECT, p As POINTAPI, L As Long
Dim Rtn As Long
L = GetWindowRect(UserControl.hWnd, r)
L = GetCursorPos(p)
GetCursorPos POS
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then '移出界面
TMIN.Enabled = False
If PTOOL.Visible = True Then PTOOL.Visible = False
If LBLNAME = "" Then Exit Sub
SBK.BackColor = &H231C09
I = 0
TMOUT.Enabled = False
Else
If LBLNAME = "" Then Exit Sub
If IS_AM = False Then
SBK.BackColor = &H7A7417
Else
If SBK.BackColor <> &H7A7417 Then TMIN.Enabled = True
End If
End If
If SBK.BorderStyle <> BOARDSTYLE Then SBK.BorderStyle = BOARDSTYLE
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DBLCLICK
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
TMOUT.Enabled = True
TMIN.Enabled = False
LBLNAME.Caption = ""
AnyWhere = 0
BOARDSTYLE = 0
Auto_Size = True
SHOWTOOL = False
SBK.BackColor = &H231C09
LBLNAME.FontName = "微软雅黑"
LBSIZE.Caption = "·" & "×" & "·"
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

Private Sub UserControl_Resize()
SBK.Move 1, 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
If LBLNAME.Caption = "" Or HAS_TXT = False Then
PBS.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
Else
PBS.Move 5, 5, UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 25
If AnyWhere = 0 Then
LBLNAME.Move 5, UserControl.ScaleHeight - LBLNAME.Height - 3
ElseIf AnyWhere = 1 Then
LBLNAME.Move (UserControl.ScaleWidth - LBLNAME.Width) / 2, UserControl.ScaleHeight - LBLNAME.Height - 3
ElseIf AnyWhere = 2 Then
LBLNAME.Move UserControl.ScaleWidth - LBLNAME.Width - 5, UserControl.ScaleHeight - LBLNAME.Height - 3
End If
End If
End Sub

Public Sub SETTXT(TIT As String)
LBLNAME.Caption = TIT
UserControl_Resize
End Sub

Public Sub SETPIC(pic As String)
On Error Resume Next
Select Case UCase(Right(pic, 3))
Case "PNG"
'Call OPENISPNG(PT, Pic)
Case "BMP", "GIF", "JPG", "ICO", "CUR"
PT.PICTURE = LoadPicture(pic)
End Select
PBS.Visible = True
cx = PT.ScaleWidth
cy = PT.ScaleHeight
cp = cx / PBS.ScaleHeight
If Auto_Size = False Then
PBS.PaintPicture PT.image, (PBS.ScaleWidth - cx) / 2, (PBS.ScaleHeight - cy) / 2, cx, cy, 0, 0, cx, cy
Else
PBS.PaintPicture PT.image, 0, 0, PBS.ScaleWidth, PBS.ScaleHeight, 0, 0, cx, cy
End If
LBSIZE.Caption = PT.ScaleWidth & "×" & PT.ScaleHeight
Call SETICO
End Sub
Public Sub SETIMG(IMG As PictureBox)
On Error Resume Next
PBS.Visible = True
PT.PICTURE = IMG.image
cx = PT.ScaleWidth
cy = PT.ScaleHeight
PBS.Cls
If Auto_Size = False Then
PBS.PaintPicture PT.image, (PBS.ScaleWidth - cx) / 2, (PBS.ScaleHeight - cy) / 2, cx, cy, 0, 0, cx, cy
Else
PBS.PaintPicture PT.image, 0, 0, PBS.ScaleWidth, PBS.ScaleHeight, 0, 0, cx, cy
End If
LBSIZE.Caption = PT.ScaleWidth & "×" & PT.ScaleHeight
Call SETICO
End Sub
Sub SETICO()
On Error Resume Next
Select Case UCase(Right(LBLNAME.Caption, 3))
Case "BMP"
ITYPE.PICTURE = IA(0).PICTURE
Case "PNG"
ITYPE.PICTURE = IA(4).PICTURE
Case "JPG"
ITYPE.PICTURE = IA(3).PICTURE
Case "GIF"
ITYPE.PICTURE = IA(1).PICTURE
Case "ICO", "CUR"
ITYPE.PICTURE = IA(2).PICTURE
End Select
End Sub
Sub SET_SIGN()
PBS.Visible = False
LBLNAME.Move (UserControl.ScaleWidth - LBLNAME.Width) / 2, (UserControl.ScaleHeight - LBLNAME.Height) / 2
End Sub
Sub REMOVE_SIGN()
PBS.Visible = True
LBLNAME.Move (UserControl.ScaleWidth - LBLNAME.Width) / 2, (UserControl.ScaleHeight - LBLNAME.Height) - 3
End Sub
