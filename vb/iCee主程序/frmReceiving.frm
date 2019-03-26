VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmReceiving 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "接收中"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Icon            =   "frmReceiving.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin ICEE.ICEE_KEY cmdCancelClose 
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
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
      ScaleWidth      =   201
      TabIndex        =   2
      Top             =   120
      Width           =   3015
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
         TabIndex        =   6
         Top             =   75
         Width           =   150
      End
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   1000
      Left            =   5520
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock wsReceive 
      Left            =   4560
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ICEE.ICEE_KEY cmdFolder 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   1095
      _ExtentX        =   2355
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY CMDOPEN 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   5850
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDownloaded 
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   3750
      Width           =   90
   End
End
Attribute VB_Name = "frmReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MYID As Long
Dim GotHeader As Boolean
Dim FileNum As Long
Dim Receivedbyt As Long
Dim ByteSec As Long, Speed As Long
Dim Complete As Boolean
Private Sub cmdCancelClose_CLICK()
  On Error Resume Next
  Complete = True
  wsReceive.Close
  Close #FileNum
  Unload Me
End Sub

Private Sub cmdFolder_CLICK()
Shell "explorer " & Left(ftRcv(MYID).Destination, Len(ftRcv(MYID).Destination) - Len(ftRcv(MYID).filename)), vbNormalFocus
End Sub

Private Sub CMDOPEN_CLICK()
Shell "Explorer " & ftRcv(MYID).Destination, vbNormalFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set ftRcv(MYID).frmReceive = Nothing
End Sub
Private Sub lblDownloaded_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)

End Sub

Private Sub lblInfo_Change()
Me.Caption = "接收中 " & lblInfo.Caption
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub tmrSpeed_Timer()
  Speed = Format(ByteSec / 1024, "0.0")
  ByteSec = 0
End Sub

Private Sub Form_Load()
ReDim ResendChunk(0)
Call SeekMe(Me)
Me.BackColor = COLOR_NOR
cmdFolder.SETCOLOR Me.BackColor, COLOR_HIGH, vbWhite
CMDOPEN.SETCOLOR Me.BackColor, COLOR_HIGH, vbWhite
cmdCancelClose.SETCOLOR Me.BackColor, COLOR_HIGH, vbWhite
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - GetTaskbarHeight
cmdFolder.SETTXT "打开文件夹"
CMDOPEN.SETTXT " 打开"
cmdCancelClose.SETTXT "关闭"
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
End Sub

Public Function Prepare(ByVal id As Long, ByVal requestID As Long)
  MYID = id
  wsReceive.accept requestID
End Function

Private Sub wsReceive_Close()
On Error Resume Next
If FileNum = 0 Then
wsReceive.Close
Unload Me
Exit Sub
End If
If Not Complete Then
Call SHOWWRONG("文件下载失败!", 0)
FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">文件下载失败"
Close #FileNum
Unload Me
End If
End Sub

Private Sub wsReceive_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
  If Not GotHeader Then
    Dim Dat As String
    wsReceive.GetData Dat$, vbString
    If Left(Dat$, 4) = "FILE" Then
      Dim FirstPos As Long, SecondPos As Long
      FirstPos = InStr(6, Dat, ":")
      SecondPos = InStr(FirstPos + 1, Dat, ":")
      With ftRcv(MYID)
        .filename = Mid(Dat, 6, (FirstPos - 6))
        .FileSize = CDbl(Mid(Dat, FirstPos + 1, (SecondPos - FirstPos) - 1))
        .Comment = Right(Dat, 200)
        .From = wsReceive.RemoteHostIP
        .frmRcOpt.Prepare MYID
      End With
      GotHeader = True
    End If
  Else
    If FileNum = 0 Then
      FileNum = FreeFile
      On Error Resume Next
      If FileLen(ftRcv(MYID).Destination) > 0 Then Kill ftRcv(MYID).Destination
      Open ftRcv(MYID).Destination For Binary As #FileNum
    End If
    Dim GotDat() As Byte
    Dim Hash As String
    ByteSec = ByteSec + bytesTotal
    Receivedbyt = Receivedbyt + bytesTotal
    Call DrawProc(pgPercent, (100 / ftRcv(MYID).FileSize) * Receivedbyt, COLOR_HIGH)
    lblDownloaded = "总大小:" & ftSend(MYID).FileSize / 1024 & "Kb 速度:" & Speed & " Kb\秒"
    LP.Caption = Int((100 / ftRcv(MYID).FileSize) * Receivedbyt)
    ReDim GotDat(0 To bytesTotal - 1)
    wsReceive.GetData GotDat, vbArray + vbByte
    Put #FileNum, , GotDat
    If Receivedbyt = ftRcv(MYID).FileSize Then
      Close #FileNum
      Complete = True
      FRMHIS.TXTNODE.Text = FRMHIS.TXTNODE.Text & vbCrLf & Now & ">文件接收成功"
      cmdCancelClose.SETTXT "关闭"
      pgPercent.Visible = False
    End If
  End If
End Sub
Function Modx(X As Single, Y As Single) As Single
Dim i%
Do While i * Y < X
i = i + 1
Loop
Modx = X - Y * (i - 1)
End Function
