VERSION 5.00
Object = "{9A226D6F-2658-4445-8D35-5C19D42676FE}#1.0#0"; "BSE.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "jp"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   562
   StartUpPosition =   2  '屏幕中心
   Begin BSE_Engine.BSE BSE1 
      Left            =   2400
      Top             =   -120
      _ExtentX        =   6588
      _ExtentY        =   1085
      OverColor       =   8388863
      SchemeStyle     =   5
      PatternBitmap   =   "Form1.frx":0000
   End
   Begin VB.CommandButton KillEngine 
      Caption         =   "关闭"
      Height          =   495
      Left            =   6720
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "BitmapPattern"
      Height          =   375
      Index           =   15
      Left            =   6840
      TabIndex        =   18
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Alien"
      Height          =   375
      Index           =   14
      Left            =   4440
      TabIndex        =   17
      Top             =   4080
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Gradient"
      Height          =   375
      Index           =   13
      Left            =   4440
      TabIndex        =   16
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Height          =   855
      Left            =   6360
      Picture         =   "Form1.frx":07BE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Windows XP Internet Explorer"
      Height          =   375
      Index           =   12
      Left            =   4440
      TabIndex        =   15
      Top             =   2880
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "OfficeXP SystemColor"
      Height          =   375
      Index           =   11
      Left            =   4440
      TabIndex        =   14
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "OfficeXP Silver"
      Height          =   375
      Index           =   10
      Left            =   4440
      TabIndex        =   13
      Top             =   1680
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Office XP Olive Green"
      Height          =   375
      Index           =   9
      Left            =   2520
      TabIndex        =   12
      Top             =   4080
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Office XP Blue"
      Height          =   375
      Index           =   8
      Left            =   2520
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Win 3.x"
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Java"
      Height          =   375
      Index           =   6
      Left            =   2520
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Netscape"
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hover"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Flat"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Windows Xp Silver"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Windows Xp Olive Green"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Windows Xp Blue"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "普通按钮"
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "带图标按钮"
      Height          =   855
      Left            =   3360
      Picture         =   "Form1.frx":14B3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
BSE1.SchemeStyle = 0
BSE1.EndSubClassing
BSE1.InitSubClassing
End Sub

Private Sub Option1_Click(Index As Integer)

BSE1.SchemeStyle = Index
BSE1.EndSubClassing
BSE1.InitSubClassing


End Sub

Private Sub KillEngine_Click()
If BSE1.EngineStarted Then BSE1.EndSubClassing
Unload Me
End Sub


