VERSION 5.00
Begin VB.Form FRMLRC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00241D0A&
   BorderStyle     =   0  'None
   Caption         =   "歌词秀"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   Icon            =   "FRMLRC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
   End
   Begin VB.TextBox TXTSINGER 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Text            =   "<歌手>"
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox TXTSONG 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Text            =   "<歌名>"
      Top             =   990
      Width           =   4935
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4830
      Picture         =   "FRMLRC.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4830
      Picture         =   "FRMLRC.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4830
      Picture         =   "FRMLRC.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   15
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_TEXT LBLRC 
      Height          =   1410
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2487
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "搜索歌词"
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
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   960
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label LBL 
      Caption         =   "Label1"
      Height          =   1095
      Left            =   4080
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "FRMLRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim lrc As String, Path As String
Sub L_EDIT()
Call SHOWWRONG("此功能正在策划中", 2)
End Sub
Sub L_VIEW()
On Error Resume Next
Dim Str1 As String
Path = App.Path & "\MEDIA\LRC\" & SONGNAME & ".lrc"
If PathFileExists(Path) = 0 Then Exit Sub
Open Path For Input As #1
While EOF(1) = False
Input #1, Str1
DoEvents
FRMSEE.TXTTS.Text = FRMSEE.TXTTS.Text & vbCrLf & Str1
Wend
Close #1
FRMSEE.Show
End Sub
Sub L_DELETE()
On Error Resume Next
If PathFileExists(Path) = 0 Then Exit Sub
Kill Path

End Sub
Sub L_SAVE()
Path = App.Path + "\Media\Lrc\" & SONGNAME & ".lrc"
Open Path For Output As #1
Print #1, LOF(1) + 1, LBL.Caption
Close #1
If IS_NET = True Then FrmNetMusic.TMLRC.Enabled = True
End Sub

Private Sub Form_Activate()
Me.BackColor = COLOR_NOR
Dim i As Integer
For i = 0 To ICM.Count - 1
ICM(i).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next
ICM(0).SETTXT "搜索"
ICM(1).SETTXT "保存"

LBLRC.SETBACKCOLOR COLOR_NOR
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B

End Sub

Private Sub Form_Load()
On Error Resume Next


LBLRC.SETFORECOLOR vbWhite
LBLRC.HASLINE = False

MakeTransparent Me.hwnd, 254
TXTSINGER.Text = frmma.LBSINGER.Caption
TXTSONG.Text = SONGNAME

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub ICM_Click(Index As Integer)
Dim i As Integer, Str1 As String, Path As String, SIN As String
On Error Resume Next
Select Case Index
Case 0
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then Call SHOWWRONG("木有联网啊亲,请检查网络状况", 2): Exit Sub
SIN = TXTSINGER.Text
If SIN = "<歌手>" Then SIN = ""
Call SERCHLRC(SIN, TXTSONG.Text, App.Path & "\MEDIA\LRC\" & SONGNAME & ".lrc")
Case 1
Call L_SAVE
End Select
End Sub

Sub SERCHLRC(SINGER As String, SONGNAME As String, SaveName As String)
On Error Resume Next
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then LBLRC.SETTXT "木有联网啊亲!": Exit Sub
Dim lrc As String, s As String
s = ""
LBL.Caption = ""
LBLRC.SETTXT ""
lrc = FindLic(SINGER, SONGNAME)
If lrc <> "" Then
URLDownloadToFile 0, lrc, SaveName, 0, 0
Open SaveName For Input As #1
Do While Not EOF(1)
    Input #1, X
    s = s + Chr(13) + X
    Loop
Close #1
LBLRC.SETTXT s
LBL.Caption = s
End If
End Sub

Private Sub LA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub TXTSINGER_GotFocus()
If TXTSINGER.Text = "<歌手>" Then TXTSINGER.Text = ""
TXTSINGER.SelStart = 0
TXTSINGER.SelLength = Len(TXTSINGER.Text)
End Sub

Private Sub TXTSINGER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call SERCHLRC(TXTSINGER.Text, TXTSONG.Text, App.Path & "\MEDIA\Thumb.lrc")

End Sub

Private Sub TXTSINGER_LostFocus()
If Trim(TXTSINGER.Text) = "" Then TXTSINGER.Text = "<歌手>"
End Sub

Private Sub TXTSONG_GotFocus()
If TXTSONG.Text = "<歌名>" Then TXTSONG.Text = ""
TXTSONG.SelStart = 0
TXTSONG.SelLength = Len(TXTSONG.Text)
End Sub

Private Sub TXTSONG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call SERCHLRC(TXTSINGER.Text, TXTSONG.Text, App.Path & "\MEDIA\Thumb.lrc")

End Sub

Private Sub TXTSONG_LostFocus()
If Trim(TXTSONG.Text) = "" Then TXTSONG.Text = "<歌名>"
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
If X3.Visible = False Then Me.Hide
End Sub
Private Sub Timer3_Timer()
Call Frmm.CHECKNET
On Error Resume Next
Path = App.Path & "\MEDIA\LRC\" & SONGNAME & ".lrc"
'思路是线检查本地是否有文件,没有的话如果自动搜索值为1且联网的话搜索，否则设置默认信息
If PathFileExists(Path) = 1 Then

Else
    If Status.RasConnState = &H2000 And AUTOSERCH = 1 Then
        Timer2.Enabled = False
        I_LRC.Visible = False
        
        lrc = FindLic(frmma.LBSINGER.Caption, SONGNAME)
        If lrc = "" Then
        
        I_LRC.Visible = False
        Timer3.Enabled = False
        If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText " ICEE音乐,音乐您的生活"
        Exit Sub
        End If
        URLDownloadToFile 0, lrc, Path, 0, 0
    Else
      
        If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText "ICEE音乐,音乐您的生活"
        Timer3.Enabled = False
        Exit Sub
    End If
End If
Timer3.Enabled = False
End Sub
Sub MOVEME()
Me.Move FrmNetMusic.Left + (FrmNetMusic.Width - Me.Width) / 2, FrmNetMusic.Top + (FrmNetMusic.Height - Me.Height) / 2
End Sub
