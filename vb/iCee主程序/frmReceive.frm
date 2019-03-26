VERSION 5.00
Begin VB.Form frmReceiveOpt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "保存文件"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   FillColor       =   &H00383537&
   ForeColor       =   &H00383537&
   Icon            =   "frmReceive.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   Begin ICEE.ICEE_KEY CMDSAVE 
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY CMDCANCEL 
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5835
      Picture         =   "frmReceive.frx":038A
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
      Left            =   5835
      Picture         =   "frmReceive.frx":046E
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
      Left            =   5835
      Picture         =   "frmReceive.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   9
      Top             =   15
      Width           =   750
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4965
   End
   Begin VB.TextBox txtFileSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label LA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "保存或者放弃？"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "我在等你呢"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00CAFDFF&
      BorderStyle     =   4  'Dash-Dot
      X1              =   16
      X2              =   424
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Label LA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件备注:"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   3
      Left            =   525
      TabIndex        =   6
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label LA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件大小:"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   525
      TabIndex        =   5
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label LA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件名称:"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   525
      TabIndex        =   4
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SSS"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   4
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   5850
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmReceiveOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Dim MYID As Long
Public Function Prepare(ByVal id As Long)
  MYID = id
  With ftRcv(id)
    lblFrom = .From & " 正在等待你接收文件."
    txtFileName = .filename
    txtFileSize = Int(.FileSize / 1024) & "Kb"
    txtComments = Trim(.Comment)
    .frmReceive.Caption = "接受文件来自 " & .From
    .frmReceive.lblInfo = .filename & " 来自 " & .From
      FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">文件接收请求,来自" & .From

  End With
  Me.Visible = True
End Function
Sub SAVEMYFILE()
On Error Resume Next
Dim filename As String
filename = App.Path & "\DOWNLOAD\" & LastFileName(ftRcv(MYID).filename)
  ftRcv(MYID).Destination = filename
  ftRcv(MYID).frmReceive.wsReceive.SendData "ACCEPT"
  FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">接受请求"

  ftRcv(MYID).frmReceive.Visible = True
  Unload Me
End Sub

Private Sub CMDCANCEL_CLICK()
On Error Resume Next
  ftRcv(MYID).frmReceive.wsReceive.SendData "DENIED"
  FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">拒绝请求"

  DoEvents
  Unload ftRcv(MYID).frmReceive
  Unload Me
End Sub

Private Sub CMDSAVE_CLICK()
  On Error Resume Next
  Call SAVEMYFILE
End Sub

Private Sub Form_Load()
On Error Resume Next
oldproc = GetWindowLong(txtFileName.hwnd, GWL_WNDPROC)
SetWindowLong txtFileName.hwnd, GWL_WNDPROC, AddressOf TextWndProc
oldproc = GetWindowLong(txtFileSize.hwnd, GWL_WNDPROC)
SetWindowLong txtFileSize.hwnd, GWL_WNDPROC, AddressOf TextWndProc
oldproc = GetWindowLong(txtComments.hwnd, GWL_WNDPROC)
SetWindowLong txtComments.hwnd, GWL_WNDPROC, AddressOf TextWndProc
Me.Move frmma.Left + (frmma.Width - Me.Width) / 2, frmma.Top + (frmma.Height - Me.Height) / 2
Me.BackColor = COLOR_NOR
LA(4).Caption = "默认保存在" & App.Path & "\DOWNLOAD\"
CMDSAVE.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
CMDCANCEL.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
CMDSAVE.SETTXT "保    存"
CMDCANCEL.SETTXT "放    弃"
txtFileName.BackColor = COLOR_NOR
txtFileSize.BackColor = COLOR_NOR
txtComments.BackColor = COLOR_NOR

Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetWindowLong txtFileName.hwnd, GWL_WNDPROC, oldproc
SetWindowLong txtFileSize.hwnd, GWL_WNDPROC, oldproc
SetWindowLong txtComments.hwnd, GWL_WNDPROC, oldproc
Set ftRcv(MYID).frmRcOpt = Nothing
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub lblFrom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
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
If X3.Visible = False Then Unload Me: ftRcv(MYID).frmReceive.wsReceive.SendData "DENIED"
End Sub
