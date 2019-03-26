VERSION 5.00
Begin VB.Form frmButtonDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JCButton按钮控件(完美支持中文,真彩色透明图标) Ver 1.7"
   ClientHeight    =   7725
   ClientLeft      =   -5205
   ClientTop       =   -2160
   ClientWidth     =   9720
   ForeColor       =   &H00EFEFEF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   515
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   StartUpPosition =   2  '屏幕中心
   Begin prjButton.AquaButton AquaButton3 
      Height          =   375
      Left            =   6120
      TabIndex        =   58
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "AquaButton3"
      CaptionEffects  =   0
      ToolTip         =   "Russian Language"
      TooltipTitle    =   "RUSSIA"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.AquaButton AquaButton2 
      Height          =   375
      Left            =   4560
      TabIndex        =   57
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "禁止"
      CaptionEffects  =   0
      TooltipBackColor=   12513791
   End
   Begin prjButton.AquaButton AquaButton1 
      Height          =   375
      Left            =   3000
      TabIndex        =   56
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "AquaButton1"
      CaptionEffects  =   0
      ToolTip         =   "Drawn by Fred.CPP"
      TooltipTitle    =   "Aqua Button"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton jcbutton1 
      Height          =   2055
      Left            =   7680
      TabIndex        =   55
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3625
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   ""
      PictureNormal   =   "frmButton.frx":0000
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin prjButton.jcbutton jcbutton15 
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Windows XP"
      PictureAlign    =   4
      CaptionEffects  =   0
      ToolTip         =   "The most famous style"
      TooltipTitle    =   "Windows XP"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton cmdTheme 
      Height          =   375
      Index           =   2
      Left            =   15
      TabIndex        =   54
      Top             =   2160
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16641248
      Caption         =   "银色"
      PictureNormal   =   "frmButton.frx":15682
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton cmdTheme 
      Height          =   375
      Index           =   1
      Left            =   15
      TabIndex        =   53
      Top             =   1680
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16641248
      Caption         =   "橄榄绿"
      PictureNormal   =   "frmButton.frx":15A1C
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton jcbutton51 
      Height          =   375
      Left            =   6120
      TabIndex        =   52
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Themed"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ToolTip         =   "Arabic Lang. with RTL support"
      TooltipTitle    =   "Arabic"
      TooltipBackColor=   0
      RightToLeft     =   -1  'True
   End
   Begin prjButton.jcbutton jcbutton50 
      Height          =   375
      Left            =   4560
      TabIndex        =   51
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   13
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "禁止"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin prjButton.jcbutton jcbutton56 
      Height          =   375
      Left            =   3000
      TabIndex        =   50
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   0
      Caption         =   "Themed"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      ToolTip         =   "Uses the current installed theme (Above XP)"
      TooltipTitle    =   "Themed Style"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton jcbutton39 
      Height          =   615
      Left            =   6960
      TabIndex        =   49
      Top             =   6240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "查看帮助"
      PictureNormal   =   "frmButton.frx":15FB6
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton jcbutton41 
      Height          =   375
      Left            =   7800
      TabIndex        =   48
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "退出"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton jcbutton36 
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   ""
      PictureNormal   =   "frmButton.frx":17008
      CaptionEffects  =   0
      MaskColor       =   -1
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton jcbutton45 
      Height          =   615
      Left            =   3000
      TabIndex        =   39
      Top             =   6240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1085
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Check out the cool features of jcbutton"
      HandPointer     =   -1  'True
      PictureNormal   =   "frmButton.frx":175A2
      PictureEffectOnOver=   0
      PictureShadow   =   -1  'True
      CaptionEffects  =   4
      TooltipType     =   1
      TooltipIcon     =   1
   End
   Begin VB.Frame Frame3 
      Caption         =   "Office XP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   35
      Top             =   1560
      Width           =   1935
      Begin prjButton.jcbutton jcbutton11 
         Height          =   405
         Left            =   1230
         TabIndex        =   36
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "1"
         CaptionEffects  =   0
      End
      Begin prjButton.jcbutton jcbutton14 
         Height          =   405
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "0"
         PictureAlign    =   4
         CaptionEffects  =   0
         MaskColor       =   16777215
      End
      Begin prjButton.jcbutton jcbutton27 
         Height          =   405
         Left            =   735
         TabIndex        =   38
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "2"
         Mode            =   1
         CaptionEffects  =   0
      End
   End
   Begin prjButton.jcbutton jcbutton46 
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   10
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "禁止"
      CaptionEffects  =   0
   End
   Begin VB.Frame Frame5 
      Caption         =   "XP 工具栏"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   7680
      TabIndex        =   20
      Top             =   600
      Width           =   1935
      Begin prjButton.jcbutton jcbutton44 
         Height          =   405
         Left            =   1230
         TabIndex        =   21
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "1"
         CaptionEffects  =   0
         ColorScheme     =   3
      End
      Begin prjButton.jcbutton jcbutton42 
         Height          =   405
         Left            =   240
         TabIndex        =   22
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "0"
         PictureAlign    =   4
         CaptionEffects  =   0
         MaskColor       =   16777215
         ColorScheme     =   3
      End
      Begin prjButton.jcbutton jcbutton43 
         Height          =   405
         Left            =   735
         TabIndex        =   23
         Top             =   300
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "2"
         Mode            =   1
         CaptionEffects  =   0
         ColorScheme     =   3
      End
   End
   Begin prjButton.jcbutton cmdVote 
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "源码爱好者"
      CaptionEffects  =   1
      ToolTip         =   "For any updates......."
      TooltipTitle    =   "Goto PSC page"
   End
   Begin prjButton.jcbutton jcbutton40 
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   12
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "禁止"
      CaptionEffects  =   2
   End
   Begin prjButton.jcbutton jcbutton38 
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "3D 悬浮"
      CaptionEffects  =   0
      ColorScheme     =   3
   End
   Begin prjButton.jcbutton cmdTest 
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   7200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "更多演示"
      CaptionEffects  =   0
      ToolTip         =   "Discover what a custom button should be aware of when takling about commercial competent buttons"
      TooltipTitle    =   "Test!"
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vista Aero Toolbar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   7680
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
      Begin prjButton.jcbutton jcbutton20 
         Height          =   405
         Left            =   1230
         TabIndex        =   32
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "1"
         CaptionEffects  =   0
         ColorScheme     =   3
      End
      Begin prjButton.jcbutton jcbutton28 
         Height          =   405
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "0"
         PictureAlign    =   4
         CaptionEffects  =   0
         MaskColor       =   16777215
         ColorScheme     =   3
      End
      Begin prjButton.jcbutton jcbutton30 
         Height          =   405
         Left            =   735
         TabIndex        =   34
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Caption         =   "2"
         Mode            =   1
         CaptionEffects  =   0
         ColorScheme     =   3
      End
   End
   Begin prjButton.jcbutton cmdTheme 
      Height          =   375
      Index           =   0
      Left            =   15
      TabIndex        =   8
      Top             =   1200
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16641248
      Caption         =   "默认(蓝)"
      PictureNormal   =   "frmButton.frx":17EF4
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton jcbutton16 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "平面按钮"
      CaptionEffects  =   0
      ToolTip         =   "As seen in the VB's toolbar"
      TooltipTitle    =   "Flat Style"
      TooltipBackColor=   -2147483624
      ColorScheme     =   3
   End
   Begin prjButton.jcbutton jcbutton10 
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "禁止"
      CaptionEffects  =   2
   End
   Begin prjButton.jcbutton jcbutton8 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "标准按钮"
      CaptionEffects  =   0
      ToolTip         =   "OLD"
      TooltipTitle    =   "Vb's Standard Button"
      TooltipBackColor=   -2147483624
      ColorScheme     =   3
   End
   Begin prjButton.jcbutton jcbutton4 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "支持32bit BMP位图"
      Mode            =   2
      Value           =   -1  'True
      HandPointer     =   -1  'True
      MouseIcon       =   "frmButton.frx":1828E
      PictureNormal   =   "frmButton.frx":183F0
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton jcbutton4 
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   7080
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "开发日志"
      Mode            =   2
      HandPointer     =   -1  'True
      MouseIcon       =   "frmButton.frx":18D42
      PictureNormal   =   "frmButton.frx":18EA4
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton jcbutton4 
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "删除文件"
      Mode            =   2
      HandPointer     =   -1  'True
      MouseIcon       =   "frmButton.frx":197F6
      PictureNormal   =   "frmButton.frx":19958
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton jcbutton4 
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "支持Alpha混合通道"
      Mode            =   2
      HandPointer     =   -1  'True
      MouseIcon       =   "frmButton.frx":1A2AA
      PictureNormal   =   "frmButton.frx":1A40C
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   16777215
      CaptionAlign    =   0
   End
   Begin prjButton.jcbutton jcbutton18 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ShowFocusRect   =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "禁止"
      CaptionEffects  =   2
   End
   Begin prjButton.jcbutton jcbutton7 
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "Install Shield"
      CaptionEffects  =   0
      ToolTip         =   "Install Shield"
      TooltipTitle    =   "As seen in the installation wizard of JetAudio"
      TooltipBackColor=   -2147483624
      ColorScheme     =   2
   End
   Begin prjButton.jcbutton jcbutton6 
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   2
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "禁止"
      CaptionEffects  =   0
   End
   Begin prjButton.jcbutton jcbutton13 
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "禁止"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin prjButton.jcbutton jcbutton19 
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Vista Aero"
      CaptionEffects  =   0
      ToolTip         =   "New Generation Vista Aero Style"
      TooltipTitle    =   "Vista"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton jcbutton21 
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   3
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "禁止"
      CaptionEffects  =   0
   End
   Begin prjButton.jcbutton cmdAqua 
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   24
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Gel"
      CaptionEffects  =   0
      MaskColor       =   16777215
      ToolTip         =   "Well, do you feel it's GELLY?"
      TooltipTitle    =   "GEL!?!"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton jcbutton2 
      Height          =   495
      Left            =   3000
      TabIndex        =   26
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   873
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   "QQ短信"
      Mode            =   1
      PictureNormal   =   "frmButton.frx":1AD5E
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
   End
   Begin prjButton.jcbutton jcbutton3 
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ButtonStyle     =   5
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   "关闭QQ"
      PictureNormal   =   "frmButton.frx":1B2F8
      CaptionEffects  =   0
      MaskColor       =   -1
   End
   Begin prjButton.jcbutton jcbutton32 
      Height          =   495
      Left            =   705
      TabIndex        =   28
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   ""
      PictureNormal   =   "frmButton.frx":1B892
      PictureAlign    =   0
      CaptionEffects  =   0
      MaskColor       =   -1
   End
   Begin prjButton.jcbutton jcbutton47 
      Height          =   495
      Left            =   4230
      TabIndex        =   30
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   873
      ButtonStyle     =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   "登陆QQ"
      PictureNormal   =   "frmButton.frx":1C61C
      CaptionEffects  =   0
      MaskColor       =   -1
   End
   Begin prjButton.jcbutton jcbutton48 
      Height          =   495
      Left            =   -120
      TabIndex        =   31
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   873
      ButtonStyle     =   5
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16572874
      Caption         =   ""
      CaptionEffects  =   0
   End
   Begin prjButton.jcbutton jcbutton22 
      Height          =   375
      Left            =   6120
      TabIndex        =   40
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "平面悬浮"
      CaptionEffects  =   0
      ToolTip         =   "Chinise Language"
      TooltipTitle    =   "Chinese"
      ColorScheme     =   3
   End
   Begin prjButton.jcbutton jcbutton23 
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   41
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "平面"
      PictureAlign    =   6
      CaptionEffects  =   0
      ToolTip         =   "Japanese"
      TooltipTitle    =   "JAPANESE"
      TooltipBackColor=   -2147483624
      ColorScheme     =   3
   End
   Begin prjButton.jcbutton jcbutton24 
      Height          =   375
      Left            =   6120
      TabIndex        =   42
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   0
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   "标准"
      CaptionEffects  =   0
      ToolTip         =   "My National Language"
      TooltipTitle    =   "HINDI"
      TooltipBackColor=   -2147483624
      ColorScheme     =   3
   End
   Begin prjButton.jcbutton jcbutton25 
      Height          =   375
      Left            =   6120
      TabIndex        =   43
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "Install Shield"
      CaptionEffects  =   0
      ToolTip         =   "Vietnaam"
      TooltipTitle    =   "Vietnaam"
      TooltipBackColor=   -2147483624
      ColorScheme     =   2
   End
   Begin prjButton.jcbutton jcbutton26 
      Height          =   375
      Left            =   6120
      TabIndex        =   44
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15199212
      Caption         =   "Windows XP"
      CaptionEffects  =   0
      ToolTip         =   "Tamil Nadu language"
      TooltipTitle    =   "Tamil"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton jcbutton31 
      Height          =   375
      Left            =   6120
      TabIndex        =   45
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Vista Aero"
      CaptionEffects  =   0
      ToolTip         =   "Gujarati Language"
      TooltipTitle    =   "Gujarati"
      TooltipBackColor=   -2147483624
   End
   Begin prjButton.jcbutton cmdAqua 
      Height          =   375
      Index           =   0
      Left            =   6120
      TabIndex        =   46
      Top             =   3720
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Gel"
      CaptionEffects  =   0
      MaskColor       =   16777215
      ToolTip         =   "Greek"
      TooltipTitle    =   "Greek"
      TooltipBackColor=   -2147483624
   End
   Begin VB.Label lblSidebar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "我的面板(Outlook 2007)"
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   120
      TabIndex        =   47
      Top             =   615
      Width           =   1980
   End
   Begin VB.Image imgOutlook 
      Height          =   420
      Left            =   0
      Picture         =   "frmButton.frx":1C9B6
      Stretch         =   -1  'True
      Top             =   510
      Width           =   2745
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   1000
      Y1              =   34
      Y2              =   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   1000
      Y1              =   33
      Y2              =   33
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   7020
      Left            =   0
      Top             =   525
      Width           =   2760
   End
   Begin VB.Menu Menu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu m1 
         Caption         =   "Inbuilt Dropdown feature"
      End
      Begin VB.Menu m2 
         Caption         =   "With Optional DropDown Symbols"
      End
      Begin VB.Menu m3 
         Caption         =   "And Optional Dropdown Separator"
      End
   End
End
Attribute VB_Name = "frmButtonDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' --Just to put appropriate colors when theme is changed at runtime...
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private lOldColor As Long

Private Sub cmdTest_Click()
    frmTest.Show
    Me.Hide
End Sub

Private Sub cmdTheme_Click(Index As Integer)

Dim ctl As Control
Dim m As Integer

    Select Case Index
    Case 0:
        imgOutlook.Picture = LoadPicture(App.Path & "\Resources\Outlook\Blue.BMP")
        SetSysColors 1, COLOR_BTNFACE, RGB(236, 233, 216)
    Case 1:
        imgOutlook.Picture = LoadPicture(App.Path & "\Resources\Outlook\Olive.BMP")
        SetSysColors 1, COLOR_BTNFACE, RGB(236, 233, 216)
    Case 2:
        imgOutlook.Picture = LoadPicture(App.Path & "\Resources\Outlook\Silver.BMP")
        SetSysColors 1, COLOR_BTNFACE, RGB(239, 239, 239)
    End Select

    For Each ctl In frmButtonDemo.Controls
        If TypeOf ctl Is jcbutton Then
            ctl.ColorScheme = Index
        End If
    Next ctl
    
    For m = 0 To 2
        cmdTheme(m).BackColor = vbWhite
    Next m
    
End Sub

Private Sub cmdVote_Click()
    
Dim sAddress As String
    
    sAddress = "http://www.codefans.net"
    'ShellExecute hWnd, "open", URL, vbNullString, vbNullString, 1
    cmdVote.OpenWebsite sAddress

End Sub

Private Sub Form_Load()

    ' --My Nation Language
    jcbutton24.Caption = "HIN: " & ChrW$(&H930) & ChrW$(&H935) & ChrW$(&H93E) & ChrW$(&H917) & ChrW$(&H924)
    ' --My mother tongue
    jcbutton31.Caption = "GUJ: " & ChrW$(2711) & ChrW$(2753) & ChrW$(2716) & ChrW$(2736) & ChrW$(2750) & ChrW$(2724) & ChrW$(2752)
    ' --Japanese
    jcbutton23(0).Caption = "JPN: " & ChrW$(&H3088) & ChrW$(&H3046) & ChrW$(&H3053) & ChrW$(&H305D)
    ' --Vietnaam
    jcbutton25.Caption = "VIE: " & ChrW$(84) & ChrW$(105) & ChrW$(234) & ChrW$(769) & ChrW$(110) & ChrW$(103) & ChrW$(32) & ChrW$(86) & ChrW$(105) & ChrW$(234) & ChrW$(803) & ChrW$(116)
    ' --Tamil
    jcbutton26.Caption = "TAM: " & ChrW$(&HB85) & ChrW$(&HB99) & ChrW$(&HBCD) & ChrW$(&HB95) & ChrW$(&HBBF) & ChrW$(&HB95)
    ' --Greek
    cmdAqua(0).Caption = "GRK: " & ChrW$(&H39A) & ChrW$(&H3B1) & ChrW$(&H3BB) & ChrW$(&H3CE) & ChrW$(&H3C2) & " " & ChrW$(&H3AE)
    ' --Chinese
    jcbutton22.Caption = "CHS: " & ChrW$(&H6B22) & ChrW$(&H8FCE)
    ' --Russian
    AquaButton3.Caption = "RUS: " & ChrW$(1056) & ChrW$(1091) & ChrW$(1089) & ChrW$(1089) & ChrW$(1082) & ChrW$(1080) & ChrW$(1081)
    ' --Arabic
    jcbutton51.Caption = "ARA: " & ChrW$(&H645) & ChrW$(&H640) & ChrW$(&H631) & ChrW$(&H62D) & ChrW$(&H628)
    
    lOldColor = GetSysColor(COLOR_BTNFACE)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SetSysColors 1, COLOR_BTNFACE, lOldColor

End Sub


Private Sub jcbutton39_Click()
    ShellExecute Me.hWnd, "open", App.Path & "\documentation.chm", vbNullString, vbNullString, 1
End Sub

Private Sub jcbutton41_Click()
    Unload Me
End Sub

Private Sub jcbutton45_Click()
    frmFeatures.Show
    Me.Hide
End Sub

