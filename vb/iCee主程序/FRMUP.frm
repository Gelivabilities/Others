VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FRMUP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H009AF4FF&
   BorderStyle     =   0  'None
   Caption         =   "在线升级"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   Icon            =   "FRMUP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PUP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H009AF4FF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   2760
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   5175
      Begin ICEE.ICEE_KEY ICM 
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
      End
      Begin VB.TextBox TXTINFO 
         BackColor       =   &H009AF4FF&
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   915
         Width           =   4095
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "更新内容"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   720
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "测试版/正式版"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本号:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   630
      End
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   2655
      Left            =   480
      TabIndex        =   9
      Top             =   4560
      Width           =   7455
      ExtentX         =   13150
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000DECC5&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7440
      Picture         =   "FRMUP.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   0
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7440
      Picture         =   "FRMUP.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7440
      Picture         =   "FRMUP.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox PF 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2100
      Left            =   360
      Picture         =   "FRMUP.frx":0636
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      Top             =   960
      Width           =   2400
   End
   Begin VB.ListBox LSTLINK 
      BackColor       =   &H0047433E&
      Height          =   1320
      Left            =   480
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前已是最新版本"
      Height          =   180
      Index           =   1
      Left            =   3840
      TabIndex        =   4
      Top             =   1680
      Width           =   1440
   End
End
Attribute VB_Name = "FRMUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UP_DL_URL As String
Private Sub Form_Load()
On Error Resume Next
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\UP_T.PNG", Me.hdc, 8, 8)
WB.Navigate "http://hi.baidu.com/iceeorgan/item/d45e07ce97859307ee46651d"
ICM.SETCOLOR Me.BackColor, &HDECC5, vbBlack
ICM.SETTXT "开始更新"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub PF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PUP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Dim i As Integer, s As String, SB As String, CAST As String, NOW_VER As String
LSTLINK.Clear
s = ""
SB = ""
CAST = "[INFO]"
For i = 0 To WB.Document.links.Length - 1
If WB.Document.links.Item(i) <> s Then
SB = WB.Document.links.Item(i).innerText 'SB是页面中所有超链接文字
s = WB.Document.links.Item(i) 'S是页面中所有超链接
If Left(UCase(SB), Len(CAST)) = CAST Then LSTLINK.AddItem SB & "|" & s
End If
Next i
WB.Silent = True
If LSTLINK.ListCount = 0 Then Exit Sub
NOW_VER = Replace(Replace(Replace(Split(LSTLINK.List(0), "|")(0), "[INFO]", ""), ".", ""), "VER", "")
LA(2).Caption = "最新版本号:" & NOW_VER
LA(3).Caption = "版本:" & Replace(Split(LSTLINK.List(2), "|")(0), "[INFO]", "")
TXTINFO.Text = Replace(Replace(Split(LSTLINK.List(3), "|")(0), "[INFO]", ""), " ", vbCrLf)
UP_DL_URL = Replace(Split(LSTLINK.List(4), "|")(0), "[INFO]", "")
If Left(NOW_VER, 2) <= App.Major & App.Minor Then PUP.Visible = False Else PUP.Visible = True
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
