VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FRMCHOSE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "选择好友"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   Icon            =   "FRMCHOSE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   1920
      Picture         =   "FRMCHOSE.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   1920
      Picture         =   "FRMCHOSE.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   1920
      Picture         =   "FRMCHOSE.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   8640
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   873
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   7815
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   13785
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   1
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "USEZT"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList USEZT 
      Left            =   360
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHOSE.frx":0636
            Key             =   "ONLINE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHOSE.frx":2310
            Key             =   "BUSY"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHOSE.frx":3FEA
            Key             =   "OFFLINE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHOSE.frx":5CC4
            Key             =   "DNZ"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMCHOSE.frx":799E
            Key             =   "UNKNOW"
         EndProperty
      EndProperty
   End
   Begin VB.Image IU 
      Height          =   705
      Left            =   5010
      Picture         =   "FRMCHOSE.frx":9678
      ToolTipText     =   "关闭"
      Top             =   15
      Width           =   750
   End
End
Attribute VB_Name = "FRMCHOSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Public W_F As String
Private Sub Form_Load()
On Error Resume Next
Dim i As Integer
Me.BackColor = COLOR_NOR
Call PaintPng(App.Path & "\SKIN\CF_T.PNG", Me.hdc, 8, 8)
ICM.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
If ALWAYSONTOP = True Then RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags) Else RESL = SetWindowPos(Me.hwnd, 1, 0, 0, 0, 0, flags)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H5C6105, B
ICM.SETTXT "分享"
For i = 1 To frmma.TreeView1.Nodes.Count
TreeView1.Nodes.Add , frmma.TreeView1.Nodes(i).Text, frmma.TreeView1.Nodes(i).Key, frmma.TreeView1.Nodes(i).Text, 1, 1
Next
TreeView1.Refresh
If frmma.Left > Me.Width Then
Me.Move frmma.Left - Me.Width, frmma.Top
Else
Me.Move frmma.Left + frmma.Width, frmma.Top
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE <> Me.X1.PICTURE Then IU.PICTURE = Me.X1.PICTURE

End Sub

Private Sub ICM_Click()
On Error GoTo ERR
If TreeView1.Nodes.Count = 0 Then Exit Sub
If TreeView1.SelectedItem.Key = "" Then Call SHOWWRONG("请选择一个好友!!" & vbCrLf & "Please choose a friend in this list!", 0): Exit Sub
If TreeView1.SelectedItem.Key = frmma.Text1.Text Then Exit Sub
  ReDim Preserve ftSend(0 To SendCount)
  Select Case DefCOM
  Case 0
  ftSend(SendCount).Comment = "我画了一幅好画,快来看看吧"
  Case 1
  ftSend(SendCount).Comment = "这首歌曲不错嘛,快来听听"
  Case 2
  ftSend(SendCount).Comment = "这张图片很不错,拿去养养眼吧"
  Case 3
  ftSend(SendCount).Comment = "这是我的截屏,想知道我在干嘛吗"
  End Select
  ftSend(SendCount).FileSize = CDbl(FileLen(W_F))
  ftSend(SendCount).To = TreeView1.SelectedItem.Key
  ftSend(SendCount).FileToSend = W_F
  
  ftSend(SendCount).frmSend.InitTransfer SendCount
  SendCount = SendCount + 1
ERR:

  Unload Me
End Sub

Private Sub ICM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ICM.ToolTipText = W_F
End Sub

Private Sub IU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IU.PICTURE = Me.X2.PICTURE Then IU.PICTURE = Me.X3.PICTURE
End Sub
Private Sub IU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE = Me.X1.PICTURE Then IU.PICTURE = Me.X2.PICTURE
End Sub
Private Sub IU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IU.PICTURE = Me.X3.PICTURE Then IU.PICTURE = Me.X1.PICTURE
Unload Me
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

