VERSION 5.00
Begin VB.Form FRMRUN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0008AF66&
   BorderStyle     =   0  'None
   Caption         =   "启动选项卡"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   6720
      Picture         =   "FRMRUN.frx":0000
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   0
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   6720
      Picture         =   "FRMRUN.frx":00E4
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   6720
      Picture         =   "FRMRUN.frx":01C8
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   615
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   615
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   615
      Index           =   3
      Left            =   5640
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择启动后进入的功能"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2100
   End
End
Attribute VB_Name = "FRMRUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
Dim I As Integer
For I = 0 To ICM.Count - 1
ICM(I).SETCOLOR Me.BackColor, &H4000&, vbWhite
Next
ICM(0).SETTXT "主界面(默认)"
ICM(1).SETTXT "绘图模式"
ICM(2).SETTXT "图像浏览"
ICM(3).SETTXT "文件管理"
ICM(GetInitEntry("SYSTEM", "AUTORUN", 0)).IS_SELECT = True
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
Dim I As Integer
For I = 0 To ICM.Count - 1
ICM(I).IS_SELECT = False
Next
ICM(Index).IS_SELECT = True
lRet = SetInitEntry("SYSTEM", "AUTORUN", Index)
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

