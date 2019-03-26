VERSION 5.00
Begin VB.Form Frmadd 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "添加新任务"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   Icon            =   "Frmadd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   259
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4845
      Picture         =   "Frmadd.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4845
      Picture         =   "Frmadd.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4845
      Picture         =   "Frmadd.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   15
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   3735
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   1
      Left            =   3360
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
      _ExtentX        =   1931
      _ExtentY        =   873
   End
   Begin VB.Label LA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下载地址:"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label LA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "存储名称:"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label LabState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   90
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H009D9899&
      FillColor       =   &H00221C13&
      Height          =   495
      Index           =   1
      Left            =   360
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H009D9899&
      FillColor       =   &H00221C13&
      Height          =   495
      Index           =   3
      Left            =   360
      Top             =   1920
      Width           =   4935
   End
End
Attribute VB_Name = "Frmadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intPos As Integer
Dim strTmp As String
Dim URL As String, Dname As String
Private Sub Form_Activate()
Dim mStr1 As String
Me.Move (FRMDOWN.Width - Me.Width) / 2 + FRMDOWN.Left, FRMDOWN.Top + (FRMDOWN.Height - Me.Height) / 2
mStr1 = Clipboard.GetText
If InStr(1, mStr1, "http://") = 1 Or InStr(1, mStr1, "https://") = 1 And Len(mStr1) < 180 Then Text1.Text = mStr1
Me.BackColor = COLOR_NOR
Call PaintPng(App.Path & "\SKIN\ND_T.PNG", Me.hdc, 8, 8)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Dim i As Integer
For i = 0 To ICM.Count - 1
ICM(i).SETCOLOR Me.BackColor, COLOR_HIGH, vbWhite
Next
End Sub

Private Sub Form_Load()
Call FrmTrans(Me)
Call SeekMe(Me)
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
ICM(0).SETTXT "确    定"
ICM(1).SETTXT "取    消"
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
FRMDOWN.Enabled = True
Me.Hide
End Sub
Sub DOWNLOADIT()
If Len(Text1.Text) = 0 Then LabState.Caption = "Url不得为空!": Text1.SetFocus: Exit Sub
If Len(Text3.Text) = 0 Then LabState.Caption = "下载名称不得为空!": Text3.SetFocus: Exit Sub
If YesNoUrl(Text1.Text) = False Then LabState.Caption = "不是个合法的Url!": Text1.Text = "": Text1.SetFocus: Exit Sub
URL = Text1.Text
Dname = Text3.Text
If Right(Dpath, 1) <> "\" Then Dpath = Dpath & "\" '补"\"
Dim ii As Integer
For ii = 1 To FRMDOWN.LVIEW.ListItems.Count
    If FRMDOWN.LVIEW.ListItems(ii).Text = Dname Then
        If FRMDOWN.LVIEW.ListItems(ii).SubItems(4) = Dpath Then
            If Right(FRMDOWN.LVIEW.ListItems(ii).SubItems(2), 1) = "%" Then
                LabState.Caption = "已存在此下载任务!请重新命名!"
                Text3.Text = ""
                Exit Sub
            End If
        End If
    End If
Next ii
''判断结束
FRMDOWN.LstLog.AddItem "[" & TimE & "]:" & Text1
Call BeginDown(URL, Dpath, Dname)
Me.Hide
End Sub
Private Sub ICM_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Call DOWNLOADIT
Case 1
Unload Me
End Select
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LabState_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Text1_Change()
On Error Resume Next
 Dim fullPath$, jj%
 fullPath = Text1
 jj = InStrRev(fullPath, "/")
 Text3 = Mid$(fullPath, jj + 1, Len(fullPath) - jj)
 If Text1 = Text3 Then Text3 = ""
Select Case UCase(Split(Text1.Text, "//")(0))
Case "HTTP:"
  Text1.Text = Text1.Text
Case "QDL:"
  intPos = InStr(Text1.Text, "://")
  strTmp = Base64Decode(Mid$(Text1.Text, intPos + 3))
  Text1.Text = strTmp
  LabState.Caption = "来自QQ旋风下载链接"
Case "THUNDER:"
  intPos = InStr(Text1.Text, "://")
  strTmp = Base64Decode(Mid$(Text1.Text, intPos + 3))
  Text1.Text = Mid$(strTmp, 3, Len(strTmp) - 4)
  LabState.Caption = "来自迅雷下载链接"
Case "FLASHGET:"
  intPos = InStr(Text1.Text, "://")
  strTmp = Base64Decode(Mid$(Text1.Text, intPos + 3))
  Text1.Text = Mid$(strTmp, Len("[FLASHGET]") + 1, Len(strTmp) - Len("[FLASHGET]") * 2)
  LabState.Caption = "来自快车下载链接"
End Select
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
