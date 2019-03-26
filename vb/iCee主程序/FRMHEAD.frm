VERSION 5.00
Begin VB.Form FRMHEAD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "更换头像"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   Icon            =   "FRMHEAD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox P60 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   1200
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   15
      Top             =   6960
      Width           =   900
      Begin VB.Image IA 
         Enabled         =   0   'False
         Height          =   240
         Index           =   4
         Left            =   600
         Picture         =   "FRMHEAD.frx":038A
         Top             =   600
         Width           =   240
      End
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   0
      Left            =   308
      TabIndex        =   12
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4905
      Picture         =   "FRMHEAD.frx":0714
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   11
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4905
      Picture         =   "FRMHEAD.frx":07F8
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   10
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4905
      Picture         =   "FRMHEAD.frx":08DC
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   9
      Top             =   15
      Width           =   750
   End
   Begin ICEE.ucScrollbar SGRO 
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   423
   End
   Begin ICEE.ucScrollbar SCRO 
      Height          =   4935
      Left            =   5280
      TabIndex        =   7
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   9128
   End
   Begin VB.PictureBox PICCLIP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2760
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PSIT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2040
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox P100 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   2640
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   6480
      Width           =   1500
      Begin VB.Image IA 
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   1200
         Picture         =   "FRMHEAD.frx":09C0
         Top             =   1200
         Width           =   240
      End
   End
   Begin VB.PictureBox PS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4890
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   343
      TabIndex        =   5
      Top             =   855
      Width           =   5145
      Begin VB.PictureBox PPV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5385
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   359
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   6
         Top             =   240
         Width           =   5385
         Begin VB.Shape shRect 
            DrawMode        =   6  'Mask Pen Not
            Height          =   2145
            Left            =   480
            Top             =   240
            Visible         =   0   'False
            Width           =   1725
         End
      End
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   1
      Left            =   1508
      TabIndex        =   13
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   2
      Left            =   2708
      TabIndex        =   14
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   3
      Left            =   4028
      TabIndex        =   16
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100×100"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   3000
      TabIndex        =   1
      Top             =   8040
      Width           =   720
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60×60"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      Top             =   8040
      Width           =   540
   End
End
Attribute VB_Name = "FRMHEAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sFile As String, XHi, XLo, YHi, YLo, StopDraw
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Sub Form_Activate()
Me.BackColor = COLOR_NOR
Dim i As Integer
For i = 0 To ICM.Count - 1
ICM(i).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\HD_T.PNG", Me.hdc, 8, 8)
End Sub

Private Sub Form_Load()
On Error Resume Next
H_CHANGE = True
If frmma.Left > Me.Width Then
Me.Move frmma.Left - Me.Width, frmma.Top
Else
Me.Move frmma.Left + frmma.Width, frmma.Top
End If
RPC.ROUND_PIC P60, 8, 0, 0
RPC.ROUND_PIC P100, 8, 0, 0
SGRO.Orientation = oHorizontal
0
StopDraw = 1
ICM(0).SETTXT "确定修改"
ICM(1).SETTXT "取消"
ICM(2).SETTXT "浏览"
ICM(3).SETTXT "恢复默认"
Call SeekMe(Me)
LOGO = GetSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
Call OpenFile(LOGO)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
H_CHANGE = False
End Sub

Private Sub ICM_Click(Index As Integer)
'On Error Resume Next
Select Case Index
Case 0
Call SavePicture(PICCLIP.image, App.Path & "\COFING\Self_Logo.Bmp")
LOGO = App.Path & "\COFING\Self_Logo.Bmp"
Call SaveSetting("ICEE", "Main", "logo", App.Path & "\COFING\Self_Logo.Bmp")
FRMSETINFO.STLOGO.PICTURE = LoadPicture(App.Path & "\COFING\Self_Logo.Bmp")
FRMSETINFO.STLOGO.ToolTipText = LOGO
Call frmma.LoadParam
Unload Me
Case 1
Unload Me
Case 2
sFile = ShowOpen(Me.hwnd, "头像图像文件" & Chr$(0) & "*.Bmp;*.GIF;*.JPG", "设置头像")
If sFile = "" Or Dir(sFile) = "" Then Exit Sub
Call OpenFile(sFile)
Case 3
Call SaveSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
Call OpenFile(App.Path + "\Skin\DefaultHead.Bmp")
End Select
End Sub
Private Function fncGetInfo(lsPicName As String) As PICINFO '不使用控件获得图片大小
    Dim hBitmap As Long
    Dim res As Long
    Dim Bmp As BITMAP
    res = GetObject(LoadPicture(lsPicName).handle, Len(Bmp), Bmp) '取得BITMAP的结构
    fncGetInfo.PicWidth = Bmp.bmWidth
    fncGetInfo.PicHeight = Bmp.bmHeight
End Function
Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 6 Then Call SetHand
End Sub

Private Sub P100_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub P60_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PPV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StopDraw = 0
XLo = X
YLo = Y
XHi = X
YHi = Y
shRect.Width = Abs(XHi - XLo)
shRect.Height = Abs(YHi - YLo)

End Sub

Private Sub PPV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
XHi = X
YHi = Y
If XHi < 0 Then XHi = 0
If YHi < 0 Then YHi = 0
If XHi > PPV.ScaleWidth Then XHi = PPV.ScaleWidth
If YHi > PPV.ScaleHeight Then YHi = PPV.ScaleHeight
If StopDraw = 0 And Button = 1 Then
shRect.Width = Abs(XHi - XLo)
shRect.Height = Abs(YHi - YLo)
shRect.Visible = True
        If XHi > XLo And YHi > YLo Then
            shRect.Top = YLo
            shRect.Left = XLo
        End If
        If XHi > XLo And YHi < YLo Then
            shRect.Top = YHi
            shRect.Left = XLo
        End If
        If XHi < XLo And YHi < YLo Then
            shRect.Top = YHi
            shRect.Left = XHi
        End If
        If XHi < XLo And YHi > YLo Then
            shRect.Top = YLo
            shRect.Left = XHi
        End If
End If
DoEvents
End Sub

Private Sub PPV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If shRect.Width < 20 Or shRect.Height < 20 Then Exit Sub
    StopDraw = 1
    P60.Cls
    P100.Cls
    PICCLIP.Cls
    PICCLIP.Width = shRect.Width
    PICCLIP.Height = shRect.Height
    PICCLIP.PaintPicture PSIT.PICTURE, 0, 0, shRect.Width, shRect.Height, shRect.Left, shRect.Top, shRect.Width, shRect.Height
    If shRect.Width < 2 Or shRect.Height < 2 Then Exit Sub
    P60.PaintPicture PICCLIP.image, 0, 0, 60, 60
    P100.PaintPicture PICCLIP.image, 0, 0, 100, 100
End Sub

Private Sub PPV_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strpath As String
If Data.files.Count > 0 Then
strpath = Data.files(1)
Select Case LCase$(Right$(Data.files(1), 3))
Case "bmp", "dib", "gif", "jpg"
Call OpenFile(strpath)
Case "png"
Call OPENISPNG(PSIT, strpath)
End Select
End If
End Sub

Private Sub PS_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strpath As String
If Data.files.Count > 0 Then
strpath = Data.files(1)
Select Case LCase$(Right$(Data.files(1), 3))
Case "bmp", "dib", "gif", "jpg"
Call OpenFile(strpath)
Case "png"
Call OPENISPNG(PSIT, strpath)
End Select
End If

End Sub

Private Sub SCRO_Change()
PPV.Top = -SCRO.Value
End Sub

Private Sub SCRO_Scroll()
SCRO_Change
End Sub

Private Sub SGRO_Change()
PPV.Left = -SGRO.Value
End Sub

Private Sub SGRO_Scroll()
SGRO_Change
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
If X3.Visible = False Then Unload Me
End Sub
Sub OpenFile(PIC As String)
On Error Resume Next
'If fncGetInfo(pic).PicHeight > 2000 Or fncGetInfo(pic).PicWidth > 2000 Then Call SHOWWRONG("图像过大,无法使用", 2): Exit Sub
'If fncGetInfo(pic).PicHeight < 50 Or fncGetInfo(pic).PicWidth < 50 Then Call SHOWWRONG("图像过小,无法使用", 2): Exit Sub
P60.Cls
P100.Cls
PPV.Cls
PICCLIP.Cls
PSIT.PICTURE = LoadPicture(PIC)
PPV.Move 0, 0, PSIT.ScaleWidth, PSIT.ScaleHeight
PPV.PaintPicture PSIT.image, 0, 0  ', (PPV.Width - PSIT.Width) / 2, (PPV.Height - PSIT.Height) / 2
shRect.Move 0, 0, PPV.ScaleWidth, PPV.ScaleHeight
PICCLIP.Width = shRect.Width
PICCLIP.Height = shRect.Height
PICCLIP.PaintPicture PSIT.image, 0, 0, shRect.Width, shRect.Height, shRect.Left, shRect.Top, shRect.Width, shRect.Height
If shRect.Width < 2 Or shRect.Height < 2 Then Exit Sub
P60.PaintPicture PICCLIP.image, 0, 0, 60, 60
P100.PaintPicture PICCLIP.image, 0, 0, 100, 100
SGRO.Max = PPV.Width - PS.ScaleWidth + 50
SCRO.Max = PPV.Height - PS.ScaleHeight + 50
SCRO.Value = 0
SGRO.Value = 0

End Sub
