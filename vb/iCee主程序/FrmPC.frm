VERSION 5.00
Object = "{95C4D06B-0E76-491A-99C9-7BD3D4D1E34F}#1.0#0"; "Shadow.OCX"
Begin VB.Form FrmPC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "是否同意对方的私聊请求？"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   Icon            =   "FrmPC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox C1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4305
      Picture         =   "FrmPC.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   15
      Width           =   750
   End
   Begin prjShadowCtl.ucShadow ucShadow1 
      Left            =   2520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Depth           =   20
   End
   Begin VB.PictureBox C2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4305
      Picture         =   "FrmPC.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4305
      Picture         =   "FrmPC.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin VB.Label ts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "你是否接受来自谁的即时聊天请求?"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   3675
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RTDATA As String
Public QRESULT As Integer
Public QKIND As Integer
Private Sub c1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = False
C2.Visible = True
End Sub
Private Sub c2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
C2.Visible = False
C3.Visible = True
End If
End Sub
Private Sub c3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C3.Visible = False
C1.Visible = True
If C3.Visible = False Then
frmma.Winsock1.SendData ".CancelRTChat " & RTChatTemp
Unload Me
End If
End Sub


Private Sub Form_Load()
Me.BackColor = COLOR_NOR
ICM(0).SETTXT "确定"
ICM(0).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICM(1).SETTXT "取消"
ICM(1).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite

Call PaintPng(App.Path & "\SKIN\W_T.PNG", Me.hdc, 8, 8)
Call PaintPng(App.Path & "\SKIN\MSG_ASK.PNG", Me.hdc, 8, 40)
MakeTransparent Me.hWnd, 250
RESL = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
ts.Top = (Me.ScaleHeight - ts.Height) / 2
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H201400, B
If Sound = 1 Then sndPlaySound App.Path + "\Sound\popo.wav", 1
Me.Move (0.5 * Screen.Width - Me.Width), (Screen.Height - Me.Height) / 2
Select Case QKIND
Case 0

Case 1

Case 3

Case 4

Case 5

End Select
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = True
C2.Visible = False
C3.Visible = False
End Sub

Private Sub ICM_Click(Index As Integer)
Select Case Index
Case 0
Dim NewRTChat As New FRMRTCHAT
NewRTChat.Show
NewRTChat.Caption = RTDATA
NewRTChat.MYTIT = RTDATA
NewRTChat.Winsock1.Close '开始重置监听端口
NewRTChat.Winsock1.LocalPort = "1981" ' Set the local port to listen on by getting the value from the txtPort text box.
NewRTChat.Winsock1.Listen ' Listen for the connection request by the other computer.
DoEvents '挂起程序
frmma.Winsock1.SendData ".BeginRTChat " & RTChatTemp '本地向对方发送即时聊天请求
Unload Me
Case 1
frmma.Winsock1.SendData ".CancelRTChat " & RTChatTemp
Unload Me
End Select
End Sub

Private Sub imgInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub ts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
