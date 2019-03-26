VERSION 5.00
Begin VB.Form FRMMYID 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00261700&
   BorderStyle     =   0  'None
   Caption         =   "我的账号"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   Icon            =   "FRMMYID.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   2
      Left            =   120
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   576
      TabIndex        =   14
      Top             =   3480
      Width           =   8640
      Begin VB.PictureBox PN 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Index           =   3
         Left            =   120
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   561
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   8415
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   9
            Left            =   3240
            TabIndex        =   49
            Top             =   2280
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
         End
      End
      Begin VB.PictureBox PN 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5055
         Index           =   2
         Left            =   120
         ScaleHeight     =   337
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   561
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   8415
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   4170
            Left            =   120
            TabIndex        =   40
            Top             =   330
            Width           =   3015
         End
         Begin VB.TextBox TTL 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   360
            Width           =   4935
         End
         Begin VB.TextBox TXTNODE 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   3255
            Left            =   3360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   1200
            Width           =   4935
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   4
            Left            =   6480
            TabIndex        =   35
            Top             =   4560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   42
            Top             =   4560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   6
            Left            =   1920
            TabIndex        =   43
            Top             =   4560
            Width           =   1215
            _ExtentX        =   3201
            _ExtentY        =   873
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "已发布的日志"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   1080
         End
         Begin VB.Shape SB 
            BorderColor     =   &H0047491F&
            BorderWidth     =   3
            Height          =   375
            Index           =   2
            Left            =   3360
            Top             =   360
            Width           =   4935
         End
         Begin VB.Shape SB 
            BorderColor     =   &H005C6105&
            BorderWidth     =   3
            Height          =   3255
            Index           =   1
            Left            =   3360
            Top             =   1200
            Width           =   4935
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "正文"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   3360
            TabIndex        =   39
            Top             =   840
            Width           =   360
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标题"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   3360
            TabIndex        =   38
            Top             =   120
            Width           =   360
         End
      End
      Begin VB.PictureBox PN 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Index           =   1
         Left            =   120
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   561
         TabIndex        =   26
         Top             =   120
         Visible         =   0   'False
         Width           =   8415
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   7
            Left            =   6480
            TabIndex        =   44
            Top             =   4560
            Width           =   1935
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   8
            Left            =   4560
            TabIndex        =   48
            Top             =   4560
            Width           =   1935
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin VB.PictureBox P_UP 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4335
            Left            =   120
            ScaleHeight     =   289
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   545
            TabIndex        =   45
            Top             =   120
            Visible         =   0   'False
            Width           =   8175
            Begin ICEE.ucScrollbar SCRO 
               Height          =   4335
               Left            =   7800
               TabIndex        =   46
               Top             =   0
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   6376
            End
         End
         Begin VB.PictureBox P_D 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4335
            Left            =   120
            ScaleHeight     =   289
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   545
            TabIndex        =   47
            Top             =   120
            Width           =   8175
         End
      End
      Begin VB.PictureBox PN 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5175
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   561
         TabIndex        =   25
         Top             =   120
         Width           =   8415
         Begin VB.ListBox PLIST 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   4170
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   8175
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   4560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   1
            Left            =   1920
            TabIndex        =   32
            Top             =   4560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   2
            Left            =   3720
            TabIndex        =   33
            Top             =   4560
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICS 
            Height          =   495
            Index           =   3
            Left            =   5520
            TabIndex        =   34
            Top             =   4560
            Width           =   2775
            _ExtentX        =   3201
            _ExtentY        =   873
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "默认列表(更改后自动同步)"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   2160
         End
      End
   End
   Begin VB.PictureBox USELOGO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   120
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   720
      Width           =   2175
      Begin VB.Image IMUSE 
         Height          =   375
         Left            =   480
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Index           =   0
      Left            =   2280
      ScaleHeight     =   1695
      ScaleWidth      =   6480
      TabIndex        =   1
      Top             =   1200
      Width           =   6480
      Begin VB.Image IA 
         Height          =   240
         Index           =   0
         Left            =   6120
         Picture         =   "FRMMYID.frx":038A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "意见反馈"
         ForeColor       =   &H0094A63E&
         Height          =   780
         Index           =   0
         Left            =   6120
         TabIndex        =   50
         Top             =   720
         Width           =   240
         WordWrap        =   -1  'True
      End
      Begin VB.Label LBPIC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   5280
         TabIndex        =   13
         Top             =   1080
         Width           =   150
      End
      Begin VB.Label LBMUSIC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Top             =   1080
         Width           =   150
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "图像"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   5040
         TabIndex        =   11
         Top             =   480
         Width           =   600
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "音乐"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   3840
         TabIndex        =   10
         Top             =   480
         Width           =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   4
         X1              =   4680
         X2              =   4680
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   3
         X1              =   3480
         X2              =   3480
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Label LBNOTE 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   1080
         Width           =   150
      End
      Begin VB.Label LBNOT 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   150
      End
      Begin VB.Label LBFANS 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   690
         TabIndex        =   5
         Top             =   1080
         Width           =   150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   2
         X1              =   2400
         X2              =   2400
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   1
         X1              =   1200
         X2              =   1200
         Y1              =   240
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   360
         X2              =   6000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日志"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   4
         Top             =   480
         Width           =   600
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收藏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   600
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "好友"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   600
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00828637&
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   2280
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   434
      TabIndex        =   8
      Top             =   720
      Width           =   6510
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "扩容"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   6000
         TabIndex        =   24
         Top             =   120
         Width           =   360
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20Mb"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   5400
         TabIndex        =   23
         Top             =   120
         Width           =   360
      End
      Begin VB.Label LBUSE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "示范用户"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1020
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00828637&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   3
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   578
      TabIndex        =   15
      Top             =   8640
      Width           =   8670
      Begin VB.Label LBPASS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12580"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   840
         TabIndex        =   18
         Top             =   240
         Width           =   450
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "立刻修改"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   7800
         TabIndex        =   17
         Top             =   240
         Width           =   720
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   360
      End
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   615
      Index           =   3
      Left            =   6600
      TabIndex        =   22
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   615
      Index           =   2
      Left            =   4440
      TabIndex        =   21
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   20
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
   End
   Begin VB.Image X1 
      Height          =   300
      Left            =   8040
      Picture         =   "FRMMYID.frx":0714
      Top             =   120
      Width           =   720
   End
   Begin VB.Image X2 
      Height          =   300
      Left            =   8040
      Picture         =   "FRMMYID.frx":1298
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image X3 
      Height          =   300
      Left            =   8040
      Picture         =   "FRMMYID.frx":1E1C
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "FRMMYID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub Form_Load()
RPC.ROUND_PIC PO(1), 4, 0, 0
RPC.ROUND_PIC PO(3), 4, 0, 0
LOGO = GetSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
IMUSE.PICTURE = LoadPicture(LOGO)
USELOGO.PaintPicture IMUSE.PICTURE, 0, 0, USELOGO.ScaleWidth, USELOGO.ScaleHeight
ICM(0).SETTXT "我的歌单"
ICM(1).SETTXT "我的图片"
ICM(2).SETTXT "我的日志"
ICM(3).SETTXT "我的收藏"
ICM(0).SETMESEL True

ICS(0).SETTXT "添加文件"
ICS(1).SETTXT "删除选中"
ICS(2).SETTXT "清空列表"
ICS(3).SETTXT "刷新"
ICS(4).SETTXT "发布日志"
ICS(5).SETTXT "删除选中"
ICS(6).SETTXT "刷新"
ICS(7).SETTXT "上传图片"
ICS(8).SETTXT "已上传的"
ICS(9).SETTXT "进入音乐窗"

Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
LBUSE.Caption = frmma.Text1.Text
LBPASS.Caption = frmma.Text2.Text
Call PaintPng(App.Path & "\SKIN\ID_T.PNG", Me.hdc, 8, 8)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub ICM_Click(INDEX As Integer)
Dim I As Integer
For I = 0 To ICM.Count - 1
ICM(I).SETMESEL False
ICM(INDEX).SETMESEL True
Next
For I = 0 To PN.Count - 1
PN(I).Visible = False
PN(INDEX).Visible = True
Next
End Sub

Private Sub LA_MouseDown(INDEX As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub LB_MouseMove(INDEX As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call SetHand
End Sub

Private Sub LB_MouseUp(INDEX As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Select Case INDEX
Case 0
Call Frmm.Report
Case 1
Call Frmm.CHANGEPASS
End Select
End Sub

Private Sub LBUSE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseDown(INDEX As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(INDEX As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub x1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then Unload Me
End Sub
