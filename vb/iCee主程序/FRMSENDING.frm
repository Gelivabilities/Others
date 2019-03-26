VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FRMSENDING 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "发送中"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ForeColor       =   &H00231C09&
   Icon            =   "FRMSENDING.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin ICEE.ICEE_KEY cmdCancelClose 
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
   End
   Begin VB.PictureBox pgPercent 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.Label LP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   120
         TabIndex        =   4
         Top             =   75
         Width           =   150
      End
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   1000
      Left            =   5160
      Top             =   720
   End
   Begin MSWinsockLib.Winsock wsSend 
      Left            =   5640
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblSent 
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   3750
      Width           =   90
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   5940
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FRMSENDING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MYID As Long
Dim FileNum As Long
Dim filename As String
Dim RCVAccept As Boolean
Dim Sentbyt As Long
Dim ByteSec As Long, Speed As Long
Dim Complete As Boolean
Public Function InitTransfer(ByVal id As Long)
  MYID = id
  filename = Mid(ftSend(MYID).FileToSend, InStrRev(ftSend(MYID).FileToSend, "\") + 1)
  Caption = "发送文件至:" & ftSend(MYID).To
  lblInfo = filename & " 至 " & ftSend(MYID).To
  wsSend.Connect ftSend(MYID).To, FT_USE_PORT
  Me.Visible = True
End Function
Private Sub cmdCancelClose_CLICK()
On Error Resume Next
  Complete = True
  Close #FileNum
  Unload Me
End Sub

Private Sub Form_Load()
Call SeekMe(Me)
Me.BackColor = COLOR_NOR
cmdCancelClose.SETCOLOR Me.BackColor, COLOR_HIGH, vbWhite
cmdCancelClose.SETTXT "关闭"
Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - GetTaskbarHeight
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set ftSend(MYID).frmSend = Nothing
End Sub
Private Sub lblInfo_Change()
Me.Caption = "发送中 " & lblInfo.Caption
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)

End Sub
Private Sub lblSent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)

End Sub

Private Sub LP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub tmrSpeed_Timer()
  Speed = Format(ByteSec / 1024, "0.0")
  ByteSec = 0
End Sub

Private Sub wsSend_Close()
  On Error Resume Next
  If Not Complete Then
  Call SHOWWRONG("文件发送失败!", 0)
  FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">文件发送失败"

Close #FileNum
Unload Me
  End If
End Sub

Private Sub wsSend_Connect()
  'Send Information regarding the file
  wsSend.SendData "FILE:" & filename & ":" & ftSend(MYID).FileSize & ":" & ftSend(MYID).Comment
End Sub

Private Sub wsSend_DataArrival(ByVal bytesTotal As Long)

    Dim Dat As String
    wsSend.GetData Dat, vbString
    If Trim$(Dat$) = "ACCEPT" Then
      Call SendChunk
      pgPercent.Visible = True
      FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">文件传输被接受"
    ElseIf Trim$(Dat$) = "DENIED" Then
    Call SHOWWRONG("对方拒绝接收文件!", 0)
    FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">文件传输被拒绝"
      wsSend.Close
      Unload Me
    End If
    
End Sub

Private Sub wsSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Select Case Number
    Case sckConnectionRefused, sckHostNotFound, sckHostNotFoundTryAgain
      Call SHOWWRONG("与对方连接失败!" & vbCrLf & "错误代码:" & Number, 0)
      FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">连接失败"

      Unload Me
  End Select
End Sub

Public Function SendChunk()
  'This is where we send the file data
  Dim ChunkSize As Long
  Dim Chunk() As Byte
  Dim arrHash() As Byte
  If wsSend.State <> sckConnected Then Exit Function
  ChunkSize = FT_BUFFER_SIZE
  If FileNum = 0 Then 'No data has been sent yet, open the file
    FileNum = FreeFile
    Open ftSend(MYID).FileToSend For Binary As #FileNum
  End If
  
  'determine chunk size
  If (LOF(FileNum) - Loc(FileNum)) < FT_BUFFER_SIZE Then _
     ChunkSize = (LOF(FileNum) - Loc(FileNum))
  'set array size to fit chunk
  ReDim Chunk(0 To ChunkSize - 1)
  Get #FileNum, , Chunk
  wsSend.SendData Chunk
  Sentbyt = Sentbyt + ChunkSize
  ByteSec = ByteSec + ChunkSize
   Call DrawProc(pgPercent, Int((100 / ftSend(MYID).FileSize) * Sentbyt), COLOR_HIGH)
  lblSent = "总大小: " & ftSend(MYID).FileSize / 1024 & "Kb 速度: " & Speed & " Kb\秒"
    LP.Caption = Int((100 / ftSend(MYID).FileSize) * Sentbyt)
  If Sentbyt = ftSend(MYID).FileSize Then
    Complete = True
    FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">发送文件成功"
    Close #FileNum
    cmdCancelClose.SETTXT "关闭"
    pgPercent.Visible = False
  End If
End Function

Private Sub wsSend_SendComplete()
  DoEvents
  If FileNum > 0 Then
      If Not Complete Then SendChunk
  End If
End Sub

Function Modx(X As Single, Y As Single) As Single
Dim i%
Do While i * Y < X
i = i + 1
Loop
Modx = X - Y * (i - 1)
End Function

