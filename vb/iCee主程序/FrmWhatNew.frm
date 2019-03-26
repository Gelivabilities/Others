VERSION 5.00
Begin VB.Form FrmWhatNew 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "更新内容"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   Icon            =   "FrmWhatNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PEND 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   4440
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   0
      Top             =   15
      Width           =   690
      Begin VB.PictureBox X2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   0
         Picture         =   "FrmWhatNew.frx":038A
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox X3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   0
         Picture         =   "FrmWhatNew.frx":046E
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox X1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   0
         Picture         =   "FrmWhatNew.frx":0552
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   1
         Top             =   0
         Width           =   750
      End
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   1335
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
      _extentx        =   2355
      _extenty        =   2355
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   1335
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   7920
      Width           =   1335
      _extentx        =   2355
      _extenty        =   2355
   End
   Begin VB.Image IQCODE 
      Height          =   2775
      Left            =   120
      Picture         =   "FrmWhatNew.frx":0636
      Stretch         =   -1  'True
      ToolTipText     =   "扫一扫,联系作者"
      Top             =   6480
      Width           =   2775
   End
End
Attribute VB_Name = "FrmWhatNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
READYLOAD = True
Me.Cls
Me.PaintPicture Frmm.IMGAD.PICTURE, 0, 0, Me.ScaleWidth, Me.ScaleHeight
Call PaintPng(App.Path & "\SKIN\ABOUTME.PNG", Me.hdc, 48, 184)
For I = 0 To IW.Count - 1
IW(I).HASLINE = False
IW(I).SETCOLOR COLOR_HIGH, COLOR_NOR
IW(I).SETTXTCOLOR vbWhite, vbWhite
Next
IW(0).SETTIP "检查更新"
IW(1).SETTIP "帮助与支持"
IW(0).SETPNG App.Path & "\SKIN\UP.PNG", (IW(0).Width - 64) / 2, (IW(0).Height - 64) / 2
IW(1).SETPNG App.Path & "\SKIN\HELP.PNG", (IW(1).Width - 64) / 2, (IW(1).Height - 64) / 2
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Me.Move frmma.Left, frmma.Top
frmma.Enabled = False
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
READYLOAD = False
frmma.Enabled = True
lRet = SetInitEntry("WHATNEW", "LEFT", Me.Left)
lRet = SetInitEntry("WHATNEW", "TOP", Me.Top)
End Sub

Private Sub GIF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub IA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub ICH_CLICK()
FrmHelp.Show
End Sub
Private Sub ICU_Click()
FRMUP.Show
End Sub

Private Sub Image1_Click()

End Sub

Private Sub IQCODE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub IW_Click(Index As Integer)
Select Case Index
Case 0
FRMUP.Show
Case 1
FrmHelp.Show
End Select
End Sub

Private Sub PCODE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub TXTABOUT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
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
If X3.Visible = False Then Unload Me
End Sub
