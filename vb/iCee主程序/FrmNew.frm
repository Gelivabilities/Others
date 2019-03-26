VERSION 5.00
Begin VB.Form FrmNew 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D6AB16&
   BorderStyle     =   0  'None
   Caption         =   "便签"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   ForeColor       =   &H00404040&
   Icon            =   "FrmNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PICSET 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6180
      Left            =   120
      ScaleHeight     =   412
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   239
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   3585
      Begin VB.PictureBox PCO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   600
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   16
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox PIW 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   480
         ScaleHeight     =   201
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   4
         Top             =   120
         Width           =   3375
         Begin ICEE.ICHECK ICK 
            Height          =   855
            Left            =   240
            TabIndex        =   14
            Top             =   2040
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1508
         End
         Begin VB.ComboBox CBTXT 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   2  'OFF
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   2895
         End
         Begin VB.PictureBox PICBK 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00CEF5F3&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   47
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.PictureBox PICFORE 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1680
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   55
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox CMBF 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   2  'OFF
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "背景颜色"
            ForeColor       =   &H00383636&
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "字体颜色"
            ForeColor       =   &H00383636&
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   11
            Top             =   120
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "字体大小"
            ForeColor       =   &H00383636&
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "字体"
            ForeColor       =   &H00383636&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1320
            Width           =   360
         End
      End
      Begin ICEE.ICEE_KEY ICC 
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
      End
   End
   Begin ICEE.ITXT TxtTS 
      Height          =   4335
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      NumbarVisible   =   0   'False
   End
   Begin VB.Label ICM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
      ForeColor       =   &H00404040&
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Shape SB 
      BorderColor     =   &H00E0E0E0&
      Height          =   255
      Left            =   3000
      Top             =   75
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ICM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "＋"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   210
   End
   Begin VB.Label ICM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   210
   End
   Begin VB.Image PJIAO 
      Height          =   165
      Left            =   360
      MousePointer    =   8  'Size NW SE
      Top             =   120
      Width           =   150
   End
   Begin VB.Shape SHA 
      BackColor       =   &H007E5502&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "FrmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gg, gg2

Private Sub CBTXT_Click()
lRet = SetInitEntry("NOTEBOOK", "FONTSIZE", CBTXT.ListIndex)
TXTTS.Font.Size = CBTXT.Text
End Sub

Private Sub CMBF_Click()
TXTTS.Font.name = CMBF.Text
lRet = SetInitEntry("NOTEBOOK", "FONTNAME", CMBF.Text)
End Sub

Private Sub Form_Load()
Dim MYLEFT, MYTOP, MYWIDTH, MYHEIGHT
NOTECOUND = NOTECOUND + 1
MYLEFT = GetInitEntry("NOTEBOOK", "LEFT", frmma.Left - 100)
MYTOP = GetInitEntry("NOTEBOOK", "TOP", frmma.Top - 100)
MYWIDTH = GetInitEntry("NOTEBOOK", "WIDTH", frmma.Width)
MYHEIGHT = GetInitEntry("NOTEBOOK", "HEIGHT", frmma.Height)
Me.Move MYLEFT + 100, MYTOP + 100, MYWIDTH, MYHEIGHT
Call SeekMe(Me)
lRet = SetInitEntry("NOTEBOOK", "LEFT", Me.Left) + 100
lRet = SetInitEntry("NOTEBOOK", "TOP", Me.Top) + 100
Form_Paint
PJIAO.Tag = ""
ICC.SETCOLOR vbWhite, &HD6AB16, vbBlack
ICC.SETTXT "确定"
Dim I As Integer
For I = 8 To 50
CBTXT.AddItem I
Next
ICK.M_STYLE = 3
ICK.SETTXT "显示行数"
ICK.Value = GetInitEntry("NOTE", "LINE_COUNT", 0)
If ICK.Value = 1 Then
TXTTS.Numbar_Visible = True
Else
TXTTS.Numbar_Visible = False
End If
CBTXT.ListIndex = GetInitEntry("NOTEBOOK", "FONTSIZE", 10)
TXTTS.Font.Size = CBTXT.Text
Call FillComboWithFonts(CMBF)
TXTTS.Font.name = GetInitEntry("NOTEBOOK", "FONTNAME", "微软雅黑")
CMBF.Text = TXTTS.Font.name
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
TXTTS.Text = "请输入需要记录的文本"
TXTTS.NumBackColor = COLOR_NOR
TXTTS.NumForeColor = vbWhite
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SB.Visible = True Then SB.Visible = False
End Sub

Private Sub Form_Paint()
On Error Resume Next
Dim I As Integer, r As Long, G As Long, b As Long, ColorRGB As Long, colorrgb_n As Long
Me.BackColor = GetInitEntry("NOTEBOOK", "BACKCOLOR", &HCEF5F3)
PCO.BackColor = Me.BackColor
Call PaintPng(App.Path & "\SKIN\WHITE.PNG", PCO.hdc, 0, 0) '画一次,区分开来
SHA.BackColor = PCO.POINT(0, 0)
TXTTS.BackColor = Me.BackColor
PICBK.BackColor = Me.BackColor
Me.FOREColor = TXTTS.FOREColor
For I = 0 To ICM.Count - 1
ICM(I).FOREColor = TXTTS.FOREColor
Next
ColorRGB = Me.POINT(2, 2)
   r = ColorRGB Mod 256 + 20
   G = ColorRGB \ 256 Mod 256 + 20
   b = ColorRGB \ 256 \ 256 + 20
   colorrgb_n = RGB(r - 5, G - 5, b - 5)
Me.Cls
Me.CurrentX = 5
Me.CurrentY = 8
'Me.Print "便签 " & "[" & NOTECOUND & "]"
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), PCO.POINT(0, 0), B
TXTTS.FOREColor = GetInitEntry("NOTEBOOK", "FORECOLOR", vbBlack)
PICFORE.BackColor = TXTTS.FOREColor

End Sub

Private Sub Form_Resize()
ICM(0).Left = Me.ScaleWidth - ICM(0).Width - 10
ICM(2).Left = ICM(0).Left - ICM(2).Width - 10
TXTTS.Move 3, 40, Me.ScaleWidth - 6, Me.ScaleHeight - TXTTS.Top - 6
PJIAO.Move Me.ScaleWidth - PJIAO.Width, Me.ScaleHeight - PJIAO.Height
PICSET.Move TXTTS.Left, TXTTS.Top, TXTTS.Width, TXTTS.Height
Call Form_Paint
SHA.Move 0, 0, Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
lRet = SetInitEntry("NOTEBOOK", "LEFT", Me.Left) + 100
lRet = SetInitEntry("NOTEBOOK", "TOP", Me.Top) + 100
lRet = SetInitEntry("NOTEBOOK", "WIDTH", Me.Width)
lRet = SetInitEntry("NOTEBOOK", "HEIGHT", Me.Height)
NOTECOUND = NOTECOUND - 1
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub ICC_Click()
Dim I As Integer
For I = 0 To ICM.Count - 1
ICM(I).Enabled = True
Next
PICSET.Visible = False
Form_Paint
End Sub

Private Sub ICK_Click()
If ICK.Value = 1 Then
TXTTS.Numbar_Visible = True
Else
TXTTS.Numbar_Visible = False
End If
lRet = SetInitEntry("NOTE", "LINE_COUNT", ICK.Value)
End Sub

Private Sub ICM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
ICM(Index).Left = ICM(Index).Left - 1
ICM(Index).Top = ICM(Index).Top - 1
End Sub

Private Sub ICM_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SB.Move ICM(Index).Left - 5, 5, ICM(Index).Width + 10
If SB.Visible = False Then SB.Visible = True
End Sub

Private Sub ICM_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
ICM(Index).Left = ICM(Index).Left + 1
ICM(Index).Top = ICM(Index).Top + 1

Select Case Index
Case 0
If InStr(TXTTS.Text, "请输入需要记录的文本") = 0 Then
If Len(TXTTS.Text) <> 0 Then
Open App.Path & "\COFING\NODE.txt" For Binary As #1
Put #1, LOF(1) + 1, Now & vbCrLf & TXTTS.Text & vbCrLf & vbCrLf
Close #1
End If
End If
Unload Me
Case 1
If NOTECOUND >= 10 Then Call SHOWWRONG("对不起，已经有十个桌面便签被打开，为保证程序运行稳定，您不可以继续添加便签啦", 2): Exit Sub
Dim NBQ As New FrmNew
NBQ.Move Me.Left + 200, Me.Top + 200, Me.Width, Me.Height
NBQ.Show
Case 2
Dim I As Integer
For I = 0 To ICM.Count - 1
ICM(I).Enabled = False
Next
PICSET.Visible = True
SB.Visible = False
End Select
End Sub
Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
'Call OpenFile(App.Path & "\COFING\NODE.TXT")
Call CMV(Me)
End Sub

Private Sub LA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 3
Call SetHand
End Select
End Sub

Private Sub PICBK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub PICBK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub PICBK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
On Error GoTo ERR
PICBK.BackColor = frmma.ShowColor(Me)
lRet = SetInitEntry("NOTEBOOK", "BACKCOLOR", PICBK.BackColor)
Call Form_Paint
ERR:
Exit Sub
End Sub

Private Sub PICFORE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub PICFORE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub PICFORE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
On Error GoTo ERR
PICFORE.BackColor = frmma.ShowColor(Me)
lRet = SetInitEntry("NOTEBOOK", "FORECOLOR", PICFORE.BackColor)
Call Form_Paint
ERR:
Exit Sub
End Sub

Private Sub PICSET_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PICSET_Resize()
PIW.Move (PICSET.ScaleWidth - PIW.Width) / 2, (PICSET.ScaleHeight - PIW.Height) / 2
ICC.Move PICSET.ScaleWidth - ICC.Width - 13, PICSET.ScaleHeight - ICC.Height - 13
End Sub

Private Sub Pjiao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PJIAO.Tag = "1"
End Sub
Private Sub Pjiao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If PJIAO.Tag <> "" Then
Dim pos As POINTAPI
GetCursorPos pos
gg = pos.X * 15 - Me.Left
gg2 = pos.Y * 15 - Me.Top
If gg > 4000 Then Me.Width = gg
If gg2 > 5600 Then Me.Height = gg2
End If
End Sub
Private Sub PJIAO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PJIAO.Tag = ""
End Sub

Private Sub TxtTS_GotFocus()
If TXTTS.Text = Left(TXTTS.Text, Len("请输入需要记录的文本")) = "请输入需要记录的文本" Then TXTTS.Text = ""
'TxtTS.SelStart = 0
'TxtTS.SelLength = Len(TxtTS.Text)
End Sub

Private Sub TxtTS_LostFocus()
If Len(Trim(TXTTS.Text)) = 0 Then TXTTS.Text = "请输入需要记录的文本"
End Sub

Private Sub TxtTS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SB.Visible = True Then SB.Visible = False
End Sub
Private Sub TxtTS_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Data.GetFormat(vbCFText) Then TXTTS.Text = Left(TXTTS.Text, TXTTS.SelStart) _
 & Data.GetData(vbCFText) _
 & Right(TXTTS.Text, Len(TXTTS.Text) - TXTTS.SelStart)
End Sub
