VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "糖果水晶按钮 V2.0"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":058A
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   StartUpPosition =   3  '窗口缺省
   Begin Proyecto1.CandyButton CandyButton14 
      Height          =   960
      Left            =   4320
      TabIndex        =   44
      Top             =   1545
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1693
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "是"
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton19 
      Height          =   960
      Left            =   4320
      TabIndex        =   30
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1693
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "复制"
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton20 
      Height          =   720
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   3120
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   1270
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "水晶大按钮"
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   10485920
      ColorButtonUp   =   8388736
      ColorButtonDown =   6291552
      BorderBrightness=   0
      ColorBright     =   16761087
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton3 
      Height          =   450
      Left            =   3720
      TabIndex        =   22
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "特效"
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   14704640
      ColorButtonUp   =   13668448
      ColorButtonDown =   11108432
      BorderBrightness=   0
      ColorBright     =   16775930
      DisplayHand     =   0   'False
      ColorScheme     =   2
   End
   Begin Proyecto1.CandyButton CandyButton6 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  图书库"
      Picture         =   "Form1.frx":145B8
      PictureAlignment=   2
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16728064
      ColorButtonUp   =   11316396
      ColorButtonDown =   11316396
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   -1  'True
      ColorScheme     =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   1
      Top             =   0
      Width           =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   8775
      Begin Proyecto1.CandyButton CandyButton8 
         Height          =   315
         Index           =   5
         Left            =   4320
         TabIndex        =   31
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CandyButton8"
         Style           =   2
         Checked         =   -1  'True
         ColorButtonHover=   37071
         ColorButtonUp   =   20650
         ColorButtonDown =   37119
         BorderBrightness=   -20
         ColorBright     =   65535
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin Proyecto1.CandyButton CandyButton8 
         Height          =   315
         Index           =   3
         Left            =   4680
         TabIndex        =   32
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CandyButton8"
         Style           =   2
         Checked         =   -1  'True
         ColorButtonHover=   207
         ColorButtonUp   =   2105504
         ColorButtonDown =   255
         BorderBrightness=   -20
         ColorBright     =   12632319
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin Proyecto1.CandyButton CandyButton8 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   33
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CandyButton8"
         Style           =   2
         Checked         =   -1  'True
         ColorButtonHover=   52992
         ColorButtonUp   =   40960
         ColorButtonDown =   61184
         BorderBrightness=   -20
         ColorBright     =   12648384
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin Proyecto1.CandyButton CandyButton8 
         Height          =   315
         Index           =   0
         Left            =   3600
         TabIndex        =   34
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CandyButton8"
         Style           =   2
         Checked         =   -1  'True
         ColorButtonHover=   13602816
         ColorButtonUp   =   14251856
         ColorButtonDown =   16760832
         BorderBrightness=   -20
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "多选信号灯 --> 样式 = 水晶"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   405
         TabIndex        =   21
         Top             =   240
         Width           =   3060
      End
   End
   Begin Proyecto1.CandyButton CandyButton1 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "Form1.frx":1564A
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16728064
      ColorButtonUp   =   11316396
      ColorButtonDown =   11316396
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton1 
      Height          =   375
      Index           =   1
      Left            =   5415
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "Form1.frx":15BE4
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16728064
      ColorButtonUp   =   11316396
      ColorButtonDown =   11316396
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton1 
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "Form1.frx":1617E
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16728064
      ColorButtonUp   =   11316396
      ColorButtonDown =   11316396
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "事件测试"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   8775
      Begin Proyecto1.CandyButton CandyButton5 
         Height          =   255
         Left            =   5760
         TabIndex        =   20
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "清除"
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.TextBox txtEvents 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "VB系统按钮"
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   2055
      End
      Begin Proyecto1.CandyButton CandyButton4 
         Height          =   400
         Left            =   360
         TabIndex        =   43
         Top             =   430
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "水晶按钮"
         Style           =   3
         Checked         =   0   'False
         ColorButtonHover=   15309136
         ColorButtonUp   =   13657888
         ColorButtonDown =   10512144
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "注意! VB系统的按钮使用的坐标值单位是twips(堤), 本水晶按钮使用的为pixels(像素)."
         Height          =   1275
         Left            =   6840
         TabIndex        =   18
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label lblY 
         Caption         =   "0"
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Y:"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "0"
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "X:"
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblMouseLeave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "鼠标离开"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblMouseEnter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "鼠标点击"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "鼠标坐标:"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
   Begin Proyecto1.CandyButton CandyButton10 
      Height          =   345
      Left            =   600
      TabIndex        =   23
      Top             =   165
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "Form1.frx":16718
      Style           =   5
      Checked         =   0   'False
      ColorButtonHover=   14704640
      ColorButtonUp   =   13668448
      ColorButtonDown =   11108432
      BorderBrightness=   50
      ColorBright     =   16775930
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton18 
      Height          =   465
      Left            =   90
      TabIndex        =   24
      Top             =   120
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "Form1.frx":16778
      Style           =   5
      Checked         =   0   'False
      ColorButtonHover=   14704640
      ColorButtonUp   =   13668448
      ColorButtonDown =   11108432
      BorderBrightness=   50
      ColorBright     =   16775930
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton9 
      Height          =   345
      Left            =   7800
      TabIndex        =   25
      Top             =   2040
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   12632256
      ColorButtonUp   =   11579568
      ColorButtonDown =   8421504
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton11 
      Height          =   345
      Left            =   6720
      TabIndex        =   26
      Top             =   2040
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   12632256
      ColorButtonUp   =   11579568
      ColorButtonDown =   8421504
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton1 
      Height          =   375
      Index           =   2
      Left            =   5790
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "Form1.frx":16848
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16728064
      ColorButtonUp   =   11316396
      ColorButtonDown =   11316396
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton20 
      Height          =   720
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   1270
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "水晶大按钮"
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   15305024
      ColorButtonUp   =   14714656
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16767376
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButtonButton2 
      Height          =   975
      Left            =   4320
      TabIndex        =   29
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "退出"
      Picture         =   "Form1.frx":16DE2
      Checked         =   0   'False
      ColorButtonHover=   16728064
      ColorButtonUp   =   11316396
      ColorButtonDown =   11316396
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton20 
      Height          =   450
      Index           =   0
      Left            =   1920
      TabIndex        =   35
      Top             =   1800
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Nero样式"
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   8384
      ColorButtonUp   =   160
      ColorButtonDown =   4272
      BorderBrightness=   0
      ColorBright     =   5292287
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton20 
      Height          =   450
      Index           =   1
      Left            =   1920
      TabIndex        =   36
      Top             =   1200
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "水晶中按钮"
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   9467680
      ColorButtonUp   =   8414992
      ColorButtonDown =   8406784
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton7 
      Height          =   360
      Left            =   6720
      TabIndex        =   38
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "水晶按钮"
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   1
   End
   Begin Proyecto1.CandyButton CandyButton20 
      Height          =   450
      Index           =   3
      Left            =   1920
      TabIndex        =   39
      Top             =   2400
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "水晶中按钮"
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   5
   End
   Begin Proyecto1.CandyButton CandyButton15 
      Height          =   600
      Left            =   5880
      TabIndex        =   40
      Top             =   3960
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " 源码爱好者"
      ForeColor       =   65535
      Picture         =   "Form1.frx":1736E
      PictureAlignment=   2
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   192
      ColorButtonUp   =   128
      ColorButtonDown =   255
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton12 
      Height          =   600
      Left            =   5880
      TabIndex        =   41
      Top             =   3360
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " 源码爱好者"
      ForeColor       =   255
      Picture         =   "Form1.frx":178C7
      PictureAlignment=   3
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   49152
      ColorButtonUp   =   32768
      ColorButtonDown =   65280
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton13 
      Height          =   600
      Left            =   5880
      TabIndex        =   42
      Top             =   2760
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1058
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "  CodeFans.net"
      ForeColor       =   16711680
      Picture         =   "Form1.frx":17A79
      PictureAlignment=   2
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   49344
      ColorButtonUp   =   32896
      ColorButtonDown =   65535
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Proyecto1.CandyButton CandyButton2 
      Height          =   300
      Left            =   120
      TabIndex        =   27
      Top             =   5640
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   335
      X2              =   335
      Y1              =   72
      Y2              =   97
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "XP 工具栏风格："
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   596
      X2              =   596
      Y1              =   72
      Y2              =   97
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   417
      X2              =   417
      Y1              =   73
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   336
      X2              =   596
      Y1              =   97
      Y2              =   97
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   416
      X2              =   416
      Y1              =   73
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   336
      X2              =   596
      Y1              =   71
      Y2              =   71
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "5 种不同的样式 -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'download by http://www.codefans.net

Private Sub CandyButton1_Click(Index As Integer)
    If Index < 3 Then
        CandyButton1(Index).Checked = Not CandyButton1(Index).Checked
        For i = 0 To 2
            If i <> Index Then CandyButton1(i).Checked = False
        Next
    End If
End Sub

Private Sub CandyButton1_DblClick(Index As Integer)
    If Index < 3 Then
        CandyButton1(Index).Checked = Not CandyButton1(Index).Checked
        For i = 0 To 2
            If i <> Index Then CandyButton1(i).Checked = False
        Next
    End If
End Sub

Private Sub CandyButton5_Click()
    txtEvents = ""
End Sub

Private Sub CandyButton6_Click()
    CandyButton6.Checked = Not CandyButton6.Checked
End Sub

Private Sub CandyButton6_DblClick()
    CandyButton6.Checked = Not CandyButton6.Checked
End Sub

Private Sub CandyButton8_Click(Index As Integer)
    CandyButton8(Index).Checked = Not CandyButton8(Index).Checked
End Sub

Private Sub CandyButton4_Click()
    txtEvents = txtEvents & "CandyButton4_Click" & vbCrLf
End Sub

Private Sub CandyButton4_DblClick()
    txtEvents = txtEvents & "CandyButton4_DblClick" & vbCrLf
End Sub

Private Sub CandyButton4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtEvents = txtEvents & "CandyButton4_MouseDown" & vbCrLf
End Sub

Private Sub CandyButton4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtEvents = txtEvents & "CandyButton4_MouseUp" & vbCrLf
End Sub

Private Sub CandyButton4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblX = x
    lblY = y
End Sub

Private Sub CandyButton4_MouseEnter()
    lblMouseEnter.BackColor = vbGreen
    lblMouseLeave.BackColor = vbWhite
End Sub

Private Sub CandyButton4_MouseLeave()
    lblMouseEnter.BackColor = vbWhite
    lblMouseLeave.BackColor = vbGreen
    lblX = "0"
    lblY = "0"
End Sub

Private Sub Command1_Click()
    txtEvents = txtEvents & "Command1_Click" & vbCrLf
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtEvents = txtEvents & "Command1_MouseDown" & vbCrLf
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtEvents = txtEvents & "Command1_MouseUp" & vbCrLf
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblMouseEnter.BackColor = vbButtonFace
    lblMouseLeave.BackColor = vbButtonFace
    lblX = x
    lblY = y
End Sub

