VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FRMRTCHAT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0058B143&
   BorderStyle     =   0  'None
   Caption         =   "与XXX聊天"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9840
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FRMGDICHAT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   656
   Begin VB.PictureBox PSMILE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4020
      Left            =   30
      ScaleHeight     =   268
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   653
      TabIndex        =   9
      Top             =   5340
      Visible         =   0   'False
      Width           =   9795
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   9
         Left            =   7560
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   8
         Left            =   5880
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   7
         Left            =   4200
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   6
         Left            =   2640
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   5
         Left            =   960
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   4
         Left            =   5880
         Top             =   600
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   3
         Left            =   4170
         Top             =   600
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   2
         Left            =   7560
         Top             =   600
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   1
         Left            =   2640
         Top             =   600
         Width           =   1350
      End
      Begin VB.Image IFACE 
         Height          =   1350
         Index           =   0
         Left            =   960
         Top             =   600
         Width           =   1350
      End
   End
   Begin ICEE.ICEE_KEY ICC 
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   17
      Top             =   8640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9045
      Picture         =   "FRMGDICHAT.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   16
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9045
      Picture         =   "FRMGDICHAT.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9045
      Picture         =   "FRMGDICHAT.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   14
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox PHELP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   6120
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   3375
      Begin ICEE.ICEE_KEY IMB 
         Height          =   1455
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY IMB 
         Height          =   1455
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY IMB 
         Height          =   1455
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY IMB 
         Height          =   1455
         Index           =   3
         Left            =   1800
         TabIndex        =   13
         Top             =   3840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·您还可以进行以下快速操作"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3735
         WordWrap        =   -1  'True
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·服务器数据可能会丢失,对您造成的不便请见谅"
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   3255
         WordWrap        =   -1  'True
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·用户是无法和服务器与黑名单中的人进行交互的"
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3135
         WordWrap        =   -1  'True
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·当对方关闭窗口时您也会同时关闭聊天窗口"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox RTFIN 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Text            =   "测试专用"
      Top             =   7560
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.PictureBox PICBK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   7680
      Picture         =   "FRMGDICHAT.frx":0636
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox USELOGO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008D6C18&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   8280
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RTFOUT 
      BackColor       =   &H00835C02&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   0
      Top             =   8640
      Width           =   5895
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin ICEE.ICEE_KEY ICC 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   18
      Top             =   8280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
End
Attribute VB_Name = "FRMRTCHAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ny As Long
Public MYTIT As String
Dim FACE_ID As String
Dim RE_ID As String
Dim X As Integer, xItem As ListItem
Public StrName As String
Dim i As Integer
Public Wnd As Long
'屏幕抓图过程中使用的API函数
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Sub SENDTWO()
RE_ID = Left(LTrim(RTFIN.Text), 3)
'Me.CurrentX = 50
'Me.CurrentY = ny
'Call BackGroundFORM(Me, PICBK)
Call PaintPng(App.Path & "\SKIN\ISAY.png", Me.hdc, 25, ny)

Me.PaintPicture USELOGO.PICTURE, 16, Me.ScaleHeight - 75, 65, 65
Call PaintPng(App.Path & "\SKIN\LOGO_65.png", Me.hdc, 16, Me.ScaleHeight - 75)
Call PaintPng(App.Path & "\SKIN\ISAY.png", Me.hdc, 50, Me.ScaleHeight - 75)
Me.CurrentX = 100
Me.CurrentY = ny + 30
ny = ny + 50
Me.Print RTrim(Replace(RTFIN.Text, "\BG \FK \UD \GN \GM \NW \SN \OM \LV", ""))
Call 过滤接收信息
'Else'
'Call PaintPng(App.Path & "\SKIN\ISAY.png", Me.hdc, 30, ny + 40)
'Me.Print RTrim(Replace(RTFIN.Text, "\BG \FK \UD \GN \GM \NW \SN \OM \LV", ""))

Call 过滤接收信息
'End If
RTFIN.Text = ""
End Sub
Private Sub Command1_Click()
FACE_ID = Left(LTrim(RTFOUT.Text), 3)
Me.CurrentX = 200
Me.CurrentY = ny + 30
If ny > 400 Then
ny = 100
'Me.Cls
Call BackGroundFORM(Me, PICBK)
Me.Line (0, ScaleHeight - 100)-(ScaleWidth, ScaleHeight), vbBlack, BF
Call PaintPng(App.Path & "\SKIN\USAY.png", Me.hdc, 100, ny)
Me.CurrentX = 200
Me.CurrentY = ny + 30
Me.Print RTrim(Replace(RTFOUT.Text, vbCrLf, ""))
Me.PaintPicture USELOGO.PICTURE, 16, Me.ScaleHeight - 75, 65, 65
Me.PaintPicture USELOGO.PICTURE, 140, ny + 20, 30, 30
Call PaintPng(App.Path & "\SKIN\LOGO_65.png", Me.hdc, 16, Me.ScaleHeight - 75)
Call PaintPng(App.Path & "\SKIN\ISAY.png", Me.hdc, 50, Me.ScaleHeight - 75)
Call 过滤发送信息
Me.FOREColor = vbWhite
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Me.CurrentX = 10
Me.CurrentY = 10
Me.Print MYTIT
ny = ny + 50
Else
Call PaintPng(App.Path & "\SKIN\USAY.png", Me.hdc, 100, ny)
Me.Print RTrim(Replace(RTFOUT.Text, vbCrLf, ""))
Me.PaintPicture USELOGO.PICTURE, 140, ny + 20, 30, 30
Call 过滤发送信息
ny = ny + 50
End If
RTFIN.Text = "\FK" & Rnd * 99999
RTFOUT.Text = ""
'Me.Refresh

End Sub
Sub 过滤发送信息()
Select Case FACE_ID
Case "\BG"
Call PaintPng(App.Path & "\SKIN\BORING.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\FK"
Call PaintPng(App.Path & "\SKIN\FUCK.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\GM"
Call PaintPng(App.Path & "\SKIN\GOOD_M.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\GN"
Call PaintPng(App.Path & "\SKIN\GOOD_N.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\NW"
Call PaintPng(App.Path & "\SKIN\NO_WORD.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\OM"
Call PaintPng(App.Path & "\SKIN\OMG.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\SN"
Call PaintPng(App.Path & "\SKIN\SHINE.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\LV"
Call PaintPng(App.Path & "\SKIN\SHOW_LV.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\SH"
Call PaintPng(App.Path & "\SKIN\SHY.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
Case "\UD"
Call PaintPng(App.Path & "\SKIN\U_R_GOD.PNG", Me.hdc, Me.ScaleWidth - 90, ny - 20)
End Select
End Sub
Sub 过滤接收信息()
Select Case RE_ID
Case "\BG"
Call PaintPng(App.Path & "\SKIN\BORING.PNG", Me.hdc, 5, ny - 77)
Case "\FK"
Call PaintPng(App.Path & "\SKIN\FUCK.PNG", Me.hdc, 5, ny - 77)
Case "\GM"
Call PaintPng(App.Path & "\SKIN\GOOD_M.PNG", Me.hdc, 5, ny - 77)
Case "\GN"
Call PaintPng(App.Path & "\SKIN\GOOD_N.PNG", Me.hdc, 5, ny - 77)
Case "\NW"
Call PaintPng(App.Path & "\SKIN\NO_WORD.PNG", Me.hdc, 5, ny - 77)
Case "\OM"
Call PaintPng(App.Path & "\SKIN\OMG.PNG", Me.hdc, 5, ny - 77)
Case "\SN"
Call PaintPng(App.Path & "\SKIN\SHINE.PNG", Me.hdc, 5, ny - 77)
Case "\LV"
Call PaintPng(App.Path & "\SKIN\SHOW_LV.PNG", Me.hdc, 5, ny - 77)
Case "\SH"
Call PaintPng(App.Path & "\SKIN\SHY.PNG", Me.hdc, 5, ny - 77)
Case "\UD"
Call PaintPng(App.Path & "\SKIN\U_R_GOD.PNG", Me.hdc, 5, ny - 77)
End Select
End Sub
Private Sub Form_Load()
Winsock1.Close ' Make sure that Winsock1 (our connection port) is closed on startup - just to be sure.
Call SeekMe(Me)
Randomize
R_P_THU = Int(Rnd * 9)
Select Case R_P_THU
Case 0
PICBK.PICTURE = Frmm.PIC(2).PICTURE
Case 1
PICBK.PICTURE = Frmm.PIC(44).PICTURE
Case 2
PICBK.PICTURE = Frmm.PIC(18).PICTURE
Case 3
PICBK.PICTURE = Frmm.PIC(8).PICTURE
Case 4
PICBK.PICTURE = Frmm.PIC(15).PICTURE
Case 5
PICBK.PICTURE = Frmm.PIC(3).PICTURE
Case 6
PICBK.PICTURE = Frmm.PIC(19).PICTURE
Case 7
PICBK.PICTURE = Frmm.PIC(20).PICTURE
Case 8
PICBK.PICTURE = Frmm.PIC(22).PICTURE
Case 9
PICBK.PICTURE = Frmm.PIC(23).PICTURE
End Select
'
Call BackGroundFORM(Me, PICBK)
Me.Line (0, ScaleHeight - 100)-(ScaleWidth, ScaleHeight), vbBlack, BF
LOGO = GetSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
If Len(Trim(LOGO)) > 0 And Dir(LOGO) <> "" Then
USELOGO.PICTURE = LoadPicture(LOGO)
Else
USELOGO.PICTURE = LoadPicture(App.Path + "\Skin\DefaultHead.Bmp")
End If
ny = 100 '自己发消息
Me.Caption = "与" & MYTIT & "聊天"
Me.PaintPicture USELOGO.PICTURE, 16, Me.ScaleHeight - 75, 65, 65
Call PaintPng(App.Path & "\SKIN\LOGO_65.png", Me.hdc, 16, Me.ScaleHeight - 75)
Call PaintPng(App.Path & "\SKIN\ISAY.png", Me.hdc, 50, Me.ScaleHeight - 75)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Me.CurrentX = 10
Me.CurrentY = 10
Me.FOREColor = vbWhite
Me.Print MYTIT

IMB(0).SETTXT "发送文件"
IMB(1).SETTXT "移至黑名单"
IMB(2).SETTXT "截图并保存"
IMB(3).SETTXT "截获对方屏幕"

For i = 0 To IMB.Count - 1
IMB(i).SETCOLOR vbWhite, &HCCF4F3, vbBlack
Next
ICC(0).SETTXT "表情"
ICC(1).SETTXT "···"

For i = 0 To ICC.Count - 1
ICC(i).M_STYLE = 2
ICC(i).SETCOLOR vbBlack, &H58B143, vbWhite
Next

Call PaintPng(App.Path & "\SKIN\BORING.PNG", PSMILE.hdc, IFACE(0).Left, IFACE(0).Top)
Call PaintPng(App.Path & "\SKIN\FUCK.PNG", PSMILE.hdc, IFACE(1).Left, IFACE(1).Top)
Call PaintPng(App.Path & "\SKIN\GOOD_M.PNG", PSMILE.hdc, IFACE(2).Left, IFACE(2).Top)
Call PaintPng(App.Path & "\SKIN\GOOD_N.PNG", PSMILE.hdc, IFACE(3).Left, IFACE(3).Top)
Call PaintPng(App.Path & "\SKIN\NO_WORD.PNG", PSMILE.hdc, IFACE(4).Left, IFACE(4).Top)
Call PaintPng(App.Path & "\SKIN\OMG.PNG", PSMILE.hdc, IFACE(5).Left, IFACE(5).Top)
Call PaintPng(App.Path & "\SKIN\SHINE.PNG", PSMILE.hdc, IFACE(6).Left, IFACE(6).Top)
Call PaintPng(App.Path & "\SKIN\SHOW_LV.PNG", PSMILE.hdc, IFACE(7).Left, IFACE(7).Top)
Call PaintPng(App.Path & "\SKIN\SHY.PNG", PSMILE.hdc, IFACE(8).Left, IFACE(8).Top)
Call PaintPng(App.Path & "\SKIN\U_R_GOD.PNG", PSMILE.hdc, IFACE(9).Left, IFACE(9).Top)
Call Send
MakeTransparent Me.hwnd, 254
If Sound = 1 Then sndPlaySound App.Path + "\Sound\MSG.wav", 1
End Sub
Private Sub Connect()
    On Error Resume Next ' If there's an error, resume the next command.
    Winsock1.Close ' Close any open ports (just in case).
    Winsock1.RemotePort = "1981"
    Winsock1.Connect RTChatRemoteIP ' Try to connect to the computer IP address specified in the txtRemoteIP text box, on the port specified in the txtPort text box.
End Sub

Private Sub Listen()
    Winsock1.Close
    Winsock1.LocalPort = "1981" ' Set the local port to listen on by getting the value from the txtPort text box.
    Winsock1.Listen ' Listen for the connection request by the other computer.
    DoEvents
    RTChatTemp = Me.Caption
  frmma.Winsock1.SendData ".BeginRTChat " & RTChatTemp
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PHELP.Visible = True Then PHELP.Visible = False
If PSMILE.Visible = True Then PSMILE.Visible = False
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next ' If there's an error, resume next command.
Winsock1.Close ' We want to disconnect or stop listening for a connection request, so close the connected or listening port.
End Sub

Private Sub ICC_Click(Index As Integer)
Select Case Index
Case 0
PSMILE.Visible = True
Case 1
PHELP.Visible = True
End Select
End Sub

Private Sub IFACE_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
RTFOUT.Text = "\BG" & RTFOUT.Text
Case 1
RTFOUT.Text = "\FK" & RTFOUT.Text
Case 2
RTFOUT.Text = "\GM" & RTFOUT.Text
Case 3
RTFOUT.Text = "\GN" & RTFOUT.Text
Case 4
RTFOUT.Text = "\NW" & RTFOUT.Text
Case 5
RTFOUT.Text = "\OM" & RTFOUT.Text
Case 6
RTFOUT.Text = "\SN" & RTFOUT.Text
Case 7
RTFOUT.Text = "\LV" & RTFOUT.Text
Case 8
RTFOUT.Text = "\SH" & RTFOUT.Text
Case 9
RTFOUT.Text = "\UD" & RTFOUT.Text
End Select
PSMILE.Visible = False
RTFOUT.SetFocus
RTFOUT.SelStart = Len(RTFOUT.Text)
End Sub
Private Sub IMB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Call SendFile(Me.MYTIT)
Case 1
frmma.Winsock1.SendData ".AddIgnore " & MYTIT
Case 2
Set Frmm.PICCLIP.PICTURE = CaptureWindow(Form1.hwnd, True, 0, 0, 593, 489)
Call frmma.保存一下(Frmm.PICCLIP)
Case 3
Call 截取他人屏幕
End Select
End Sub
Sub 截取他人屏幕()
  Dim strSend As String
   Dim Str_Copure As String
   StrName = MYTIT
   Str_Copure = Winsock1.LocalHostName   '获取本地主机名
   Call Send    '调用用户自定义的过程
    strSend = "AAA"
    Winsock1.SendData strSend
End Sub
Private Sub LA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 4 Then Call SetHand
End Sub

Private Sub RTFIN_Change()
If Trim(RTFIN.Text) <> "" Then Call SENDTWO
End Sub
Private Sub RTFOUT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
Call SendMessage
End If
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

Private Sub Winsock1_Close()
    Unload Me
End Sub

Private Sub Winsock1_Connect()
    On Error Resume Next ' If there's an error, continue with next command.
    RTFOUT.SetFocus ' Set the focus on the box to enter messages to send to the other computer
    RTFOUT.Enabled = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next ' Just in case there's an error, continue with next command.
    If Winsock1.State <> sckClosed Then Winsock1.Close ' Close any open socket (just in case).
    Winsock1.accept requestID ' Accept the other computer's connection request.
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim currenttext As String ' String to hold contents of RTFIn if needed.
    Dim nData As String ' Declare a variable to hold the incoming data.
    Dim TempBegin As Integer
    On Error Resume Next ' If there's an error, resume next command.
    Winsock1.GetData nData ' Get the incoming data and store it in variable "ndata".
    If InStr(1, nData, Chr(8)) Then
        TempBegin = 0
        Do While InStr(TempBegin + 1, nData, Chr(8)) > 0
            If Len(RTFIN.Text) = 0 Then Exit Sub
            TempBegin = InStr(TempBegin + 1, nData, Chr(8))
            If Len(RTFIN.Text) = 1 Then RTFIN.Text = ""
            If InStr(Len(RTFIN.Text) - 1, RTFIN.Text, Chr(13)) Then
                RTFIN.Text = Mid(RTFIN.Text, 1, Len(RTFIN.Text) - 2)
            Else
                RTFIN.Text = Mid(RTFIN.Text, 1, Len(RTFIN.Text) - 1)
            End If
        Loop
        Exit Sub
    End If
    If InStr(1, nData, Chr(13)) Then
        nData = Replace(nData, Chr(13), vbCrLf & "\par ")
        nData = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman Times New Roman;}}" & vbCrLf & "{\colortbl\red127\green127\blue127;}" & vbCrLf & "\deflang1033\pard\plain\f2\fs20\cf0 " & nData
        nData = nData & vbCrLf & "\plain\f2\fs20\par }"
        Exit Sub
    End If
    
On Error GoTo MyErr
    Dim strdata As String
    Dim strDatas As String
    Winsock1.GetData strdata
    strDatas = StrName
      If strdata = "AAA" Then
      '控制抓图
       ScreenPicture          '屏幕抓图
       SavePictureImage       '保存图象
      End If
    Exit Sub
MyErr:
  Debug.Print "出现意外错误:" & ERR.Description

End Sub
Private Sub Send()
  On Error Resume Next
    With Winsock1
    .RemoteHost = StrName     '要连接的远程计算机
    .RemotePort = 1981         '要连接的端口.
    .Bind 1981                 '绑定到本地的端口上.
    End With
End Sub
Private Sub ScreenPicture()       '屏幕抓图事件
Dim TWidth As Long
Dim THeight As Long
  TWidth = Screen.Width \ Screen.TwipsPerPixelX
  THeight = Screen.Height \ Screen.TwipsPerPixelY
    If TWidth < 1024 Or THeight < 768 Then
        SourceDC = CreateDC("DISPLAY", 0, 0, 0)
        DestDC = CreateCompatibleDC(SourceDC)
        BHandle = CreateCompatibleBitmap(SourceDC, 800, 600)
        SelectObject DestDC, BHandle
        BitBlt DestDC, 0, 0, 800, 600, SourceDC, 0, 0, &HCC0020
        Wnd = Screen.ActiveForm.hwnd
        OpenClipboard Wnd
        EmptyClipboard
        SetClipboardData 2, BHandle
        CloseClipboard
        DeleteDC DestDC
        ReleaseDC DHandle, SourceDC
       Frmm.PICCLIP.PICTURE = Clipboard.GetData()
    Else
        SourceDC = CreateDC("DISPLAY", 0, 0, 0)
        DestDC = CreateCompatibleDC(SourceDC)
        BHandle = CreateCompatibleBitmap(SourceDC, 1024, 768)
        SelectObject DestDC, BHandle
        BitBlt DestDC, 0, 0, 1024, 768, SourceDC, 0, 0, &HCC0020
        Wnd = Screen.ActiveForm.hwnd
        OpenClipboard Wnd
        EmptyClipboard
        SetClipboardData 2, BHandle
        CloseClipboard
        DeleteDC DestDC
    ReleaseDC DHandle, SourceDC
      Frmm.PICCLIP.PICTURE = Clipboard.GetData()
    End If
End Sub

Private Sub SavePictureImage()    '保存图片
   SavePicture Frmm.PICCLIP.PICTURE, App.Path & "\THUMBS\" & Int(Rnd * (1000)) & ".Bmp"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RTFOUT.SetFocus ' Set the focus back on the message box to send another message.
End Sub

Sub SendMessage()
On Error GoTo ErrRTFOKP ' If there is an error in this subroutine, go to "err" code at bottom.
RTFOUT.SelStart = Len(RTFOUT.Text) ' Set cursor to end of outgoing message box. This keeps the last message on the screen.
Winsock1.SendData RTFOUT.Text   'Chr(KeyAscii) ' Send each character (as it is typed to the other) computer.
RTFIN.Text = RTFOUT.Text
RTFOUT.Text = ""
ErrRTFOKP:
Resume Next ' Resume with next command after showing the error.
End Sub
