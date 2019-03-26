VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmNetMusic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000DECC5&
   BorderStyle     =   0  'None
   Caption         =   "音乐窗"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   Icon            =   "FrmNetMusic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PSER 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AD7900&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1680
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   119
      Top             =   720
      Visible         =   0   'False
      Width           =   5175
      Begin ICEE.ICEE_KEY ICSER 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   120
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICSER 
         Height          =   495
         Index           =   1
         Left            =   2160
         TabIndex        =   121
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
      End
   End
   Begin ICEE.ICEE_KEY PF 
      Height          =   495
      Left            =   120
      TabIndex        =   117
      Top             =   8760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   8490
      Index           =   10
      Left            =   120
      ScaleHeight     =   566
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   7
      Top             =   9360
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox TXTTEST 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   11040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "FrmNetMusic.frx":038A
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Timer TMACT 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   120
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在加载"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   10
         Left            =   0
         TabIndex        =   102
         Top             =   8880
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.PictureBox PO 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5880
      Index           =   4
      Left            =   5640
      ScaleHeight     =   392
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   88
      Top             =   1440
      Visible         =   0   'False
      Width           =   4575
      Begin ICEE.IList ILIST 
         Height          =   5385
         Left            =   15
         TabIndex        =   89
         Top             =   480
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   9631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ItemHeight      =   18
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认列表"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   101
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6165
      Index           =   9
      Left            =   -3000
      ScaleHeight     =   411
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   18
      Top             =   -6000
      Visible         =   0   'False
      Width           =   4335
      Begin ICEE.IList F_LIST 
         Height          =   6135
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   11218
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ItemHeight      =   18
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   8
      Left            =   1680
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   344
      TabIndex        =   16
      Top             =   15
      Width           =   5160
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   735
         Index           =   0
         Left            =   3960
         TabIndex        =   24
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin VB.TextBox TXTSER 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "<输入歌手或者歌曲名进行搜索>"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Shape SB 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12000
      Top             =   120
   End
   Begin VB.Timer TMP 
      Interval        =   500
      Left            =   12840
      Top             =   7800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6840
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6000
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5520
      Top             =   240
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   240
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7545
      Picture         =   "FrmNetMusic.frx":0393
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   10
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7545
      Picture         =   "FrmNetMusic.frx":0477
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   12
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7545
      Picture         =   "FrmNetMusic.frx":055B
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   11
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICS 
      Height          =   135
      Index           =   2
      Left            =   5048
      TabIndex        =   77
      Top             =   8955
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
   End
   Begin ICEE.ICEE_KEY ICS 
      Height          =   135
      Index           =   1
      Left            =   3488
      TabIndex        =   76
      Top             =   8955
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
   End
   Begin ICEE.ICEE_KEY ICS 
      Height          =   135
      Index           =   0
      Left            =   1928
      TabIndex        =   75
      Top             =   8955
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   7245
      Index           =   1
      Left            =   6000
      ScaleHeight     =   483
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   78
      Top             =   1200
      Width           =   8295
      Begin VB.PictureBox PICLRC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00565656&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   5040
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   110
         Top             =   5160
         Visible         =   0   'False
         Width           =   3135
         Begin VB.PictureBox PC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00AA7402&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   360
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   115
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox PC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H002EBC7C&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   840
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   114
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox PC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00DB59D8&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   1320
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   113
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox PC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   1800
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   112
            Top             =   360
            Width           =   375
         End
         Begin VB.PictureBox PC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   2280
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   111
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.PictureBox PICSET 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   7560
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   109
         Top             =   6240
         Width           =   630
      End
      Begin VB.Timer TMSEEK 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   2880
         Top             =   1320
      End
      Begin VB.Timer TMLRC 
         Interval        =   1000
         Left            =   2880
         Top             =   360
      End
      Begin VB.PictureBox PV 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   60
         Left            =   6000
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   97
         TabIndex        =   83
         Top             =   6600
         Width           =   1455
         Begin VB.Shape SV 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            BorderStyle     =   0  'Transparent
            Height          =   75
            Left            =   0
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00EAC037&
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   11
         Left            =   0
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   551
         TabIndex        =   79
         Top             =   7200
         Width           =   8265
         Begin VB.Shape PR 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00E0E0E0&
            BorderStyle     =   0  'Transparent
            Height          =   135
            Left            =   0
            Top             =   0
            Width           =   855
         End
      End
      Begin ICEE.ICEE_LRC L_LRC 
         Height          =   4545
         Left            =   120
         TabIndex        =   85
         Top             =   1200
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   22463
         _ExtentY        =   8017
         Begin VB.PictureBox PIC_LRC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   900
            Index           =   0
            Left            =   6840
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   87
            Top             =   1320
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.PictureBox PIC_LRC 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   900
            Index           =   1
            Left            =   6840
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   86
            Top             =   2280
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.Image IMFAV 
         Enabled         =   0   'False
         Height          =   240
         Left            =   360
         Picture         =   "FrmNetMusic.frx":063F
         Top             =   720
         Width           =   240
      End
      Begin VB.Label LSER 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手动搜索"
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
         Left            =   4680
         TabIndex        =   91
         Top             =   3480
         Width           =   840
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "暂无歌词"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   5
         Left            =   3360
         TabIndex        =   90
         Top             =   3240
         Width           =   1200
      End
      Begin VB.Image IPP 
         Height          =   720
         Index           =   1
         Left            =   1680
         Top             =   6120
         Width           =   720
      End
      Begin VB.Image IPP 
         Height          =   1080
         Index           =   0
         Left            =   360
         Top             =   6000
         Width           =   1080
      End
      Begin VB.Label LBALL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   7650
         TabIndex        =   84
         Top             =   840
         Width           =   465
      End
      Begin VB.Label LBSONG 
         BackStyle       =   0  'Transparent
         Caption         =   "I BELIVE"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   720
         TabIndex        =   82
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label LBAUTHOR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "萧亚轩"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   81
         Top             =   840
         Width           =   540
      End
      Begin VB.Label LBCOUND 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   6960
         TabIndex        =   80
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   6900
      Index           =   3
      Left            =   6960
      ScaleHeight     =   460
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   1
      Top             =   960
      Width           =   8295
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   7215
         Index           =   5
         Left            =   120
         ScaleHeight     =   481
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   553
         TabIndex        =   104
         Top             =   3960
         Visible         =   0   'False
         Width           =   8295
         Begin ICEE.IList B_LIST 
            Height          =   5460
            Left            =   120
            TabIndex        =   108
            Top             =   1320
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   9631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ItemHeight      =   18
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            Height          =   2730
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   106
            Top             =   3240
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.PictureBox PKB 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1095
            TabIndex        =   105
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "百度新歌榜"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   1
            Left            =   1560
            TabIndex        =   107
            Top             =   480
            Width           =   1425
         End
      End
      Begin VB.PictureBox P_THUMB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2595
         Index           =   4
         Left            =   2880
         Picture         =   "FrmNetMusic.frx":09C9
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.PictureBox P_THUMB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2595
         Index           =   3
         Left            =   2880
         Picture         =   "FrmNetMusic.frx":7199
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.PictureBox P_THUMB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2595
         Index           =   2
         Left            =   2880
         Picture         =   "FrmNetMusic.frx":BBE1
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.PictureBox P_THUMB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2595
         Index           =   1
         Left            =   2880
         Picture         =   "FrmNetMusic.frx":1399E
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.PictureBox P_THUMB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2595
         Index           =   0
         Left            =   2880
         Picture         =   "FrmNetMusic.frx":18FEC
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   173
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   2595
      End
      Begin ICEE.ICEE_WIN8 ICT 
         Height          =   2595
         Index           =   1
         Left            =   2760
         TabIndex        =   29
         Top             =   240
         Width           =   2595
         _ExtentX        =   5741
         _ExtentY        =   2990
      End
      Begin ICEE.ICEE_WIN8 ICT 
         Height          =   2595
         Index           =   2
         Left            =   5400
         TabIndex        =   30
         Top             =   240
         Width           =   2595
         _ExtentX        =   5741
         _ExtentY        =   2990
      End
      Begin ICEE.ICEE_WIN8 ICT 
         Height          =   2595
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   2880
         Width           =   2595
         _ExtentX        =   5741
         _ExtentY        =   2990
      End
      Begin ICEE.ICEE_WIN8 ICT 
         Height          =   2595
         Index           =   4
         Left            =   2760
         TabIndex        =   32
         Top             =   2880
         Width           =   2595
         _ExtentX        =   5741
         _ExtentY        =   2990
      End
      Begin ICEE.ICEE_WIN8 ICT 
         Height          =   2595
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2595
         _ExtentX        =   5741
         _ExtentY        =   2990
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0084536F&
      BorderStyle     =   0  'None
      Height          =   8535
      Index           =   7
      Left            =   0
      ScaleHeight     =   569
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   92
      Top             =   1920
      Visible         =   0   'False
      Width           =   8295
      Begin ICEE.IList LST 
         Height          =   2895
         Left            =   120
         TabIndex        =   94
         Top             =   5040
         Visible         =   0   'False
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5398
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ItemHeight      =   18
      End
      Begin ICEE.IMUSIC IMIC 
         Height          =   3615
         Left            =   2520
         TabIndex        =   116
         Top             =   1320
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   3836
         _ExtentY        =   4260
      End
      Begin VB.PictureBox KPB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   99
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox LISTSER 
         Height          =   2745
         IntegralHeight  =   0   'False
         Left            =   5640
         TabIndex        =   96
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   13
         Left            =   0
         ScaleHeight     =   28
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   553
         TabIndex        =   95
         Top             =   8040
         Width           =   8295
      End
      Begin ICEE.ICEE_KEY ICB 
         Height          =   495
         Index           =   0
         Left            =   6720
         TabIndex        =   93
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICB 
         Height          =   495
         Index           =   1
         Left            =   5160
         TabIndex        =   98
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "还没有文件呦!试试全盘搜索吧"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   7
         Left            =   2160
         TabIndex        =   100
         Top             =   4320
         Width           =   4005
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "本地音乐"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   6
         Left            =   1560
         TabIndex        =   97
         Top             =   480
         Width           =   1200
      End
   End
   Begin ICEE.ICEE_KEY PM 
      Height          =   495
      Left            =   6720
      TabIndex        =   118
      Top             =   8760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00411883&
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   2
      Left            =   1680
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   0
      Top             =   960
      Width           =   8295
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00899F1E&
         BorderStyle     =   0  'None
         Height          =   5460
         Index           =   12
         Left            =   0
         ScaleHeight     =   364
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   552
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   8280
         Begin VB.PictureBox PBK 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00899F1E&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1095
            TabIndex        =   13
            Top             =   120
            Width           =   1095
         End
         Begin MSComctlLib.ImageList IMV 
            Left            =   11400
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   48
            ImageHeight     =   48
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmNetMusic.frx":1EA72
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView LVIEW 
            Height          =   5535
            Left            =   1800
            TabIndex        =   9
            Top             =   1680
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   9763
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "IMV"
            SmallIcons      =   "IMV"
            ForeColor       =   16777215
            BackColor       =   9019166
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "歌手"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "收藏时间"
               Object.Width           =   3969
            EndProperty
         End
         Begin ICEE.ICEE_WIN8 IPLAY 
            Height          =   1215
            Index           =   3
            Left            =   240
            TabIndex        =   72
            Top             =   1680
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   2143
         End
         Begin ICEE.ICEE_WIN8 IPLAY 
            Height          =   1215
            Index           =   4
            Left            =   240
            TabIndex        =   73
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   2143
         End
         Begin ICEE.ICEE_WIN8 IPLAY 
            Height          =   1215
            Index           =   5
            Left            =   240
            TabIndex        =   74
            Top             =   4320
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   2143
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Index           =   2
            Left            =   3240
            TabIndex        =   67
            Top             =   600
            Width           =   270
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "我喜欢的歌手"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   0
            Left            =   1440
            TabIndex        =   14
            Top             =   480
            Width           =   1710
         End
      End
      Begin VB.PictureBox PLEFT 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001F1F1F&
         BorderStyle     =   0  'None
         Height          =   990
         Left            =   120
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   537
         TabIndex        =   35
         Top             =   6480
         Width           =   8055
         Begin VB.ListBox LISTFL 
            Height          =   600
            Left            =   1560
            TabIndex        =   37
            Top             =   5880
            Visible         =   0   'False
            Width           =   1455
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   8
            Left            =   3840
            TabIndex        =   36
            Top             =   0
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   39
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   2
            Left            =   960
            TabIndex        =   40
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   3
            Left            =   1440
            TabIndex        =   41
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   4
            Left            =   1920
            TabIndex        =   42
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   5
            Left            =   2400
            TabIndex        =   43
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   6
            Left            =   2880
            TabIndex        =   44
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   7
            Left            =   3360
            TabIndex        =   45
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   9
            Left            =   4320
            TabIndex        =   46
            Top             =   0
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   10
            Left            =   4800
            TabIndex        =   47
            Top             =   0
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   11
            Left            =   5280
            TabIndex        =   48
            Top             =   0
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   12
            Left            =   5760
            TabIndex        =   50
            Top             =   0
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   13
            Left            =   6240
            TabIndex        =   51
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   14
            Left            =   6720
            TabIndex        =   52
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   15
            Left            =   7200
            TabIndex        =   53
            Top             =   0
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   16
            Left            =   0
            TabIndex        =   54
            Top             =   480
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   17
            Left            =   480
            TabIndex        =   55
            Top             =   480
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   18
            Left            =   960
            TabIndex        =   56
            Top             =   480
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   19
            Left            =   1440
            TabIndex        =   57
            Top             =   480
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   20
            Left            =   1920
            TabIndex        =   58
            Top             =   480
            Width           =   495
            _ExtentX        =   2778
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   21
            Left            =   2400
            TabIndex        =   59
            Top             =   480
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   22
            Left            =   2880
            TabIndex        =   60
            Top             =   480
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   23
            Left            =   3360
            TabIndex        =   61
            Top             =   480
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   24
            Left            =   3840
            TabIndex        =   62
            Top             =   480
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   25
            Left            =   4320
            TabIndex        =   63
            Top             =   480
            Width           =   495
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICC 
            Height          =   495
            Index           =   26
            Left            =   4800
            TabIndex        =   64
            Top             =   480
            Width           =   855
            _ExtentX        =   1720
            _ExtentY        =   873
         End
      End
      Begin VB.PictureBox PIC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   2160
         Left            =   10200
         Picture         =   "FrmNetMusic.frx":2074C
         ScaleHeight     =   144
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   144
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.PictureBox PICFRAME 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   0
         ScaleHeight     =   433
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   553
         TabIndex        =   19
         Top             =   6120
         Width           =   8295
         Begin VB.PictureBox PO 
            AutoRedraw      =   -1  'True
            BackColor       =   &H001F1FE2&
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   7560
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   49
            TabIndex        =   65
            Top             =   0
            Width           =   735
            Begin VB.Label LBS 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   330
               TabIndex        =   66
               Top             =   90
               Width           =   120
            End
         End
         Begin VB.PictureBox PTOP 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00633F0E&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   12000
            Picture         =   "FrmNetMusic.frx":271C3
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   49
            ToolTipText     =   "返回顶部"
            Top             =   4800
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.PictureBox PICSLIDE 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7455
            Left            =   0
            ScaleHeight     =   497
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   553
            TabIndex        =   21
            Top             =   1440
            Width           =   8295
            Begin ICEE.ucScrollbar vsbSlide 
               Height          =   1215
               Left            =   7320
               TabIndex        =   33
               Top             =   2160
               Visible         =   0   'False
               Width           =   375
               _ExtentX        =   873
               _ExtentY        =   2143
            End
            Begin ICEE.ICEE_KEY optThumb 
               Height          =   615
               Index           =   0
               Left            =   0
               TabIndex        =   103
               Top             =   0
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   1085
            End
         End
      End
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         BackColor       =   &H001F1F1F&
         BorderStyle     =   0  'None
         Height          =   7335
         Index           =   6
         Left            =   240
         ScaleHeight     =   489
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   552
         TabIndex        =   22
         Top             =   240
         Width           =   8280
         Begin ICEE.ICEE_WIN8 IPLAY 
            Height          =   1215
            Index           =   0
            Left            =   6960
            TabIndex        =   69
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   2143
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   2730
            Left            =   4200
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   2400
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.PictureBox BKP 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H001F1F1F&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   1095
            TabIndex        =   23
            Top             =   120
            Width           =   1095
         End
         Begin ICEE.IList S_LIST 
            Height          =   5700
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   10160
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ItemHeight      =   18
         End
         Begin ICEE.ICEE_WIN8 IPLAY 
            Height          =   1215
            Index           =   1
            Left            =   6960
            TabIndex        =   70
            Top             =   2760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   2143
         End
         Begin ICEE.ICEE_WIN8 IPLAY 
            Height          =   1215
            Index           =   2
            Left            =   6960
            TabIndex        =   71
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   2143
         End
         Begin VB.Label LA 
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
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Index           =   3
            Left            =   3120
            TabIndex        =   68
            Top             =   600
            Width           =   135
         End
         Begin VB.Label LBNAME 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "歌手名字"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   1440
            TabIndex        =   34
            Top             =   480
            Width           =   1140
         End
      End
   End
End
Attribute VB_Name = "FrmNetMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_TIME As Long, ISMV As Boolean, WIL_B As Integer
Public Will_DL As String, M_N As String, A_N As String
Private WithEvents c_Subclass   As iSubClass
Attribute c_Subclass.VB_VarHelpID = -1
Public cDeskLrc As New clsDeskLrc
Private Const SIZE_SHOW         As Long = 60    '隐藏后留出来的宽度或高度,单位缇
Private Const SHOWHIDE_SPEED    As Long = 30    '(自动显示隐藏速度，单位缇)
Private m_ShowFlag              As Long
Private m_ShowOrient            As Long
Private m_ShowSpeed             As Long
Private m_MoveEnabled           As Boolean
Public PLAYING As Boolean, FAV_ART As Boolean
Private Const LB_INITSTORAGE = &H1A8
Private Const LB_ADDSTRING = &H180
Private Const WM_SETREDRAW = &HB
Private Const WM_VSCROLL = &H115
Private Const SB_BOTTOM = 7
Private Const INVALID_HANDLE_VALUE = -1
Private Const MaxLFNPath = 260
Dim PicHeight%, hLB&, FileSpec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%, Running%
Dim WFD As WIN32_FIND_DATA, hItem&, hFile&
Private Const vbBackslash = "\"
Private Const vbAllFiles = "*.*"
Private Const vbKeyDot = 46
Dim IS_PICTUR_YN As Boolean, ISMOVE As Boolean, Path As String, IS_MV As Boolean
Sub SETCOLOR()
B_LIST.SETCOLOR COLOR_NOR, COLOR_HIGH
ILIST.SETCOLOR COLOR_NOR, COLOR_HIGH
End Sub
Private Sub B_LIST_Click()
On Error Resume Next
Dim MName As String, aname As String
MusicName = B_LIST.List(B_LIST.ListIndex)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
Will_DL = FindMp3URL(MName, aname)
M_N = MName
A_N = aname
If Button = 2 Then Me.PopupMenu Frmm.招商
End Sub

Private Sub B_LIST_DBClick()
On Error Resume Next
Dim MName As String, aname As String, WILL_PLAY As String
MusicName = B_LIST.List(B_LIST.ListIndex)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
WILL_PLAY = FindMp3URL(MName, aname)
If WILL_PLAY = "" Then Exit Sub
frmma.PLIST.AddItem MName, aname, WILL_PLAY, 0
frmma.PLIST.ListIndex = frmma.PLIST.ListCount - 1
frmma.Wm.URL = WILL_PLAY
frmma.PLAYB.Enabled = True
专辑 = ""
年代 = ""
If IS_MINI_LIST = True Then Call FRMLIST.RELIST
Call RELIST
End Sub

Private Sub B_LIST_MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Me.PopupMenu Frmm.招商
End Sub

Private Sub F_LIST_Click()
On Error Resume Next
Dim MName As String, aname As String
MusicName = F_LIST.List(F_LIST.ListIndex)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
Will_DL = FindMp3URL(MName, aname)
M_N = MName
A_N = aname
End Sub

Private Sub F_LIST_DBClick()
On Error Resume Next
Dim MName As String, aname As String, WILL_PLAY As String
MusicName = F_LIST.List(F_LIST.ListIndex)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
WILL_PLAY = FindMp3URL(MName, aname)
If WILL_PLAY = "" Then Exit Sub
frmma.PLIST.AddItem MName, aname, WILL_PLAY, 0
frmma.PLIST.ListIndex = frmma.PLIST.ListCount - 1
frmma.Wm.URL = WILL_PLAY
frmma.PLAYB.Enabled = True
专辑 = ""
年代 = ""
IW(0).SETTXT "搜索"
If IS_MINI_LIST = True Then Call FRMLIST.RELIST
Call RELIST
End Sub

Private Sub F_LIST_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
F_LIST.SetFocus
End Sub
Sub MOVEME()
If frmma.Visible = True Then
If frmma.Left >= Me.Width / 2 Then
Me.Move frmma.Left - Me.Width, frmma.Top
Else
Me.Move frmma.Left + frmma.Width, frmma.Top
End If
Else
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End If
End Sub
Sub SETME()
'If Me.BackColor = COLOR_NOR Then Exit Sub
Me.Cls
Me.BackColor = COLOR_NOR
'Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\M_T.PNG", Me.hdc, 8, 8)
picSlide.BackColor = Me.BackColor
PICFRAME.BackColor = Me.BackColor
PO(5).BackColor = Me.BackColor
PO(3).BackColor = Me.BackColor
PO(1).BackColor = Me.BackColor
PO(10).BackColor = Me.BackColor
KPB.BackColor = Me.BackColor
PKB.BackColor = Me.BackColor
PICSET.BackColor = Me.BackColor
PLEFT.BackColor = Me.BackColor
PO(2).BackColor = Me.BackColor
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
BKP.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", BKP.hdc, 0, 0)
KPB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", KPB.hdc, 0, 0)

Call SETCOLOR
Dim I As Integer
For I = 0 To IW.Count - 1
IW(I).HASLINE = False
IW(I).HASTIP = False
Next
IW(0).SETCOLOR vbBlack, COLOR_HIGH
IW(0).SETTXT "搜索"
For I = 0 To IPLAY.Count - 1
IPLAY(I).IS_PIC = True
IPLAY(I).HASLINE = False
IPLAY(I).SETTXTCOLOR vbWhite, vbWhite
Next
L_LRC.BackColor = Me.BackColor
L_LRC.FOREColor = vbWhite
L_LRC.Karaoke = False

IPLAY(0).SETCOLOR BKP.BackColor, COLOR_HIGH
IPLAY(2).SETCOLOR BKP.BackColor, COLOR_HIGH
IPLAY(1).SETCOLOR BKP.BackColor, COLOR_HIGH
IPLAY(0).SETPNG App.Path & "\SKIN\FA_F.PNG", IPLAY(0).Width / 2 - 24, IPLAY(0).Height / 2 - 24
IPLAY(1).SETPNG App.Path & "\SKIN\S_MARK.PNG", IPLAY(1).Width / 2 - 24, IPLAY(1).Height / 2 - 24
IPLAY(2).SETPNG App.Path & "\SKIN\ADD_ALL.PNG", IPLAY(2).Width / 2 - 24, IPLAY(2).Height / 2 - 24
IPLAY(4).SETPNG App.Path & "\SKIN\DELETE.PNG", IPLAY(4).Width / 2 - 32, IPLAY(4).Height / 2 - 32
IPLAY(5).SETPNG App.Path & "\SKIN\CLOUD.PNG", IPLAY(5).Width / 2 - 32, IPLAY(5).Height / 2 - 32
PICSET.Cls
Call PaintPng(App.Path & "\SKIN\SET_N.PNG", PICSET.hdc, 0, 0)
IPLAY(2).SETTIP "添加全部"
IPLAY(0).SETTIP "收藏Ta"
IPLAY(3).IS_PIC = False
IPLAY(3).SETTXT "移除Ta"
IPLAY(3).SETFONT "微软雅黑", 15, True, 15, False
IPLAY(3).HASTIP = False
IPLAY(4).SETTIP "清空收藏"
IPLAY(5).SETTIP "云同步"
For I = 0 To ICS.Count - 1
ICS(I).SETCOLOR COLOR_HIGH, vbWhite, vbWhite
ICS(I).SETTXT ""
Next
For I = 0 To PIC_LRC.Count - 1
PIC_LRC(I).Cls
PIC_LRC(I).BackColor = COLOR_HIGH
Next
Call PaintPng(App.Path & "\SKIN\ZIN.png", PIC_LRC(0).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\ZOUT.png", PIC_LRC(1).hdc, 0, 0)
PM.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
PF.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
LST.SETCOLOR COLOR_NOR, COLOR_HIGH
PO(7).BackColor = COLOR_NOR
PO(13).BackColor = COLOR_NOR
PO(4).BackColor = COLOR_HIGH
ICB(0).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICB(1).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICB(0).SETTXT "添加本地文件夹"
ICB(1).SETTXT "全盘搜索"
For I = 0 To optThumb.Count - 1
optThumb(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next

For I = 0 To ICT.Count - 1
ICT(I).IS_PIC = True
ICT(I).HASLINE = False
ICT(I).SETCOLOR COLOR_NOR, COLOR_HIGH
ICT(I).SETTXTCOLOR vbWhite, vbWhite
ICT(I).SETIMG P_THUMB(I)
Next

IMIC.SETCOLOR COLOR_NOR, COLOR_HIGH
End Sub
Private Sub Form_Activate()
Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is PictureBox Then
PBOX.Refresh
End If
Next
Call UnHook
H_DOS = 8
gHW = Me.hwnd '鼠标控件
Call Hook '唤醒鼠标滑轮API
Call SETME
Call DRAWPLAYER
End Sub
Sub RELIST()
If frmma.PLIST.ListCount = 0 Then Exit Sub
ILIST.Clear
Dim I As Integer
For I = 0 To frmma.PLIST.ListCount - 1
Me.ILIST.AddItem frmma.PLIST.Title(I)
Next
End Sub

Private Sub Form_Load()

ISMV = False
Call SETCOLOR

IS_NET = True
L_TIME = 0

LVIEW.ColumnHeaders.Add , , "歌手名称", 300
LVIEW.ColumnHeaders.Add , , "添加时间", 100
LVIEW.HideColumnHeaders = True
PM.SETTXT "本地文件"
PF.SETTXT "000"
PO(10).Move 0, 56, Me.ScaleWidth
PO(3).Move 0, 88, Me.ScaleWidth
PO(2).Move PO(3).Width + PO(3).Left, 88, Me.ScaleWidth
PO(1).Move -PO(1).Width, 64, Me.ScaleWidth
PO(7).Move 0, 72, Me.ScaleWidth
PO(13).Move 0, PO(7).ScaleHeight - PO(13).Height, Me.ScaleWidth
PO(5).Move 0, 0
PO(4).Move 240, 200
F_LIST.Move 1, 1
Dim I As Integer, Code As String, MYLEFT, MYTOP, LASTBANG As Integer, a As String

ICT(0).SETTIP "百度网络排行榜"
ICT(1).SETTIP "百度新歌榜"
ICT(2).SETTIP "百度热歌榜"
ICT(3).SETTIP "百度老歌榜"
ICT(4).SETTIP "百度金曲榜"

ScaleMode = vbPixels
PicHeight% = PO(13).Height
hLB& = LISTSER.hwnd
SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200

For I = 0 To ICC.Count - 1
ICC(I).SETCOLOR &H1F1F1F, &HD6AB16, vbWhite
Next
For I = 0 To ICSER.Count - 1
ICSER(I).SETCOLOR Pser.BackColor, &H7E5502, vbWhite
Next
ICSER(0).SETTXT "搜歌手"
ICSER(1).SETTXT "搜歌名"

ICSER(0).IS_SELECT = True

L_LRC.SETFONTSIZE (GetInitEntry("LRC", "SIZE", 14))
PO(9).Move 112, 48
ICC(0).SETTXT "A"
ICC(1).SETTXT "B"
ICC(2).SETTXT "C"
ICC(3).SETTXT "D"
ICC(4).SETTXT "E"
ICC(5).SETTXT "F"
ICC(6).SETTXT "G"
ICC(7).SETTXT "H"
ICC(8).SETTXT "I"
ICC(9).SETTXT "J"
ICC(10).SETTXT "K"
ICC(11).SETTXT "L"
ICC(12).SETTXT "M"
ICC(13).SETTXT "N"
ICC(14).SETTXT "O"
ICC(15).SETTXT "P"
ICC(16).SETTXT "Q"
ICC(17).SETTXT "R"
ICC(18).SETTXT "S"
ICC(19).SETTXT "T"
ICC(20).SETTXT "U"
ICC(21).SETTXT "V"
ICC(22).SETTXT "W"
ICC(23).SETTXT "X"
ICC(24).SETTXT "Y"
ICC(25).SETTXT "Z"
ICC(26).SETTXT "123"
ICC(26).MY_TIT = "0123456789"
ICS(1).IS_SELECT = True
Call MOVEME
MakeTransparent Me.hwnd, 254

'GradateColors gColors, &H231C09, &H7A7417, &H231C09
'DrawProcSpectrum PO(11), 1, gColors

PO(10).Visible = True
TMACT.Enabled = True

Set c_Subclass = New iSubClass
c_Subclass.SetMsgHook Me.hwnd

Me.Show

Call PaintPng(App.Path & "\SKIN\NO_FIND.PNG", PO(2).hdc, 368, 240)
Call PaintPng(App.Path & "\SKIN\MSG_INFO.PNG", PO(2).hdc, 312, 240)

PO(6).Visible = False
PO(5).Visible = False
PICFRAME.Visible = True
PO(6).Move 0, 0, PO(2).ScaleWidth, PO(2).ScaleHeight
PO(12).Move 0, 0, PO(2).ScaleWidth, PO(2).ScaleHeight

Call LOADFAV
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\SINGER_LIST.DAT") = 0 Then
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then  '检查Internet
Call SHOWWRONG("未检测到活动网络,请重试", 2)
PO(10).Visible = False
TMACT.Enabled = False
Else
Call SINGER
End If
Else '如果本地有歌手列表就加载本地列表,这样可以节省时间
Open App.Path & "\MEDIA\MUSICBOX\SINGER_LIST.DAT" For Input As #1
Do Until EOF(1)
Input #1, a
List2.AddItem a
Loop
Close #1
Call CreateThumbs
End If
Call RELIST
Call DRAWPLAYER
IMIC.MUSIC_URL = ""
If PathFileExists(App.Path & "\MEDIA\SEARCH\MP3.CP") = 0 Then
IMIC.Visible = False
LST.Visible = False
Else
LST.Visible = True
IMIC.Visible = True
Open App.Path & "\MEDIA\SEARCH\MP3.CP" For Input As #1
Do Until EOF(1)
Input #1, a
LISTSER.AddItem a
LST.AddItem LastFileName(a)
Loop
Close #1
End If

 If frmma.Wm.URL = "" Then Exit Sub
 If FAV_IT = True Then
 IMFAV.PICTURE = Frmm.PIC(52).PICTURE
Else
IMFAV.PICTURE = Frmm.PIC(54).PICTURE
End If


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub Form_Unload(Cancel As Integer)
lRet = SetInitEntry("MusicBox", "LEFT", Me.Left)
lRet = SetInitEntry("MusicBox", "TOP", Me.Top)
End Sub
Sub LOADFAV()
On Error Resume Next
Dim filem As String, tpStr As String, I As Integer
filem = App.Path & "\MEDIA\Favourite\Favourite.isw"
If PathFileExists(filem) <> 0 Then
LVIEW.ListItems.Clear
'加载收藏夹数据
Open filem For Input As #1
Do While Not EOF(1)
With LVIEW.ListItems.Add()
For I = 0 To 1
Line Input #1, tpStr
If I = 0 Then
.Text = tpStr
Else
.SubItems(I) = tpStr
End If
.SmallIcon = 1
.Icon = 1
Next
End With
Loop
Close #1

End If
IPLAY(1).SETTIP LVIEW.ListItems.Count
LA(2).Caption = LVIEW.ListItems.Count
End Sub
Sub SAVEFAV()
On Error Resume Next
Dim filem As String
filem = App.Path & "\MEDIA\Favourite\Favourite.isw"
Dim I As Integer, tpList As ListItem
Open filem For Output As #1
For Each tpList In LVIEW.ListItems
Print #1, tpList.Text
For I = 0 To 1
Print #1, tpList.SubItems(I)
Next
Next
Close #1
End Sub

Private Sub ICB_Click(Index As Integer)
Select Case Index
Case 0

Case 1
Call START_S
End Select
End Sub

Private Sub ICC_Click(Index As Integer)
Dim I As Integer

If Index < 26 Then
For I = 0 To List2.ListCount - 1
If py(List2.List(I)) = Mid(ICC(Index).MY_TIT, 1, 1) Then LISTFL.AddItem List2.List(I)
Next
Else
For I = 0 To List2.ListCount - 1
LISTFL.AddItem List2.List(I)
If UCase(Left(List2.List(I), 1)) = "A" Then Exit For
Next
End If
Call SetInitEntry("MUSICBOX", "LAST_PAGE", Index)
Call CreateThumbs
End Sub

Private Sub ICS_Click(Index As Integer)
Dim I As Integer
For I = 0 To ICS.Count - 1
ICS(I).IS_SELECT = False
Next
ICS(Index).IS_SELECT = True
Select Case Index
Case 0
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = True
Case 1
Timer2.Enabled = True
Timer1.Enabled = False
Timer3.Enabled = False
Case 2
Timer1.Enabled = True
Timer2.Enabled = False
Timer3.Enabled = False

End Select
End Sub

Private Sub ICSER_Click(Index As Integer)
Dim I As Integer
For I = 0 To ICSER.Count - 1
ICSER(I).IS_SELECT = False
Next
ICSER(Index).IS_SELECT = True
End Sub

Private Sub ICT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Dim Code As String, I As Integer, a As String
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then Call SHOWWRONG("未检测到活动网络,请重试", 2): Exit Sub
WIL_B = Index
B_LIST.Clear
Select Case Index
Case 0
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\网络歌曲.LST") = 0 Then
PO(10).Visible = True
TMACT.Enabled = True
Code = ReadinteFile("http://music.baidu.com/top/netsong")
If Code <> "" Then GetMusicin Code
Else
PO(10).Visible = False
TMACT.Enabled = False
Open App.Path & "\MEDIA\MUSICBOX\网络歌曲.LST" For Input As #1
Do Until EOF(1)
Input #1, a
B_LIST.AddItem a
Loop
Close #1
PO(5).Visible = True
End If
Case 1
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\百度新歌.LST") = 0 Then
PO(10).Visible = True
TMACT.Enabled = True
Code = ReadinteFile("http://music.baidu.com/top/new")
If Code <> "" Then GetMusicin Code
Else
PO(10).Visible = False
TMACT.Enabled = False
Open App.Path & "\MEDIA\MUSICBOX\百度新歌.LST" For Input As #1
Do Until EOF(1)
Input #1, a
B_LIST.AddItem a
Loop
Close #1
PO(5).Visible = True
End If
Case 2
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\百度热歌.LST") = 0 Then
PO(10).Visible = True
TMACT.Enabled = True
Code = ReadinteFile("http://music.baidu.com/top/dayhot")
If Code <> "" Then GetMusicin Code
Else
PO(10).Visible = False
TMACT.Enabled = False
Open App.Path & "\MEDIA\MUSICBOX\百度热歌.LST" For Input As #1
Do Until EOF(1)
Input #1, a
B_LIST.AddItem a
Loop
Close #1
PO(5).Visible = True
End If
Case 3
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\经典歌曲.LST") = 0 Then
PO(10).Visible = True
TMACT.Enabled = True
Code = ReadinteFile("http://music.baidu.com/top/oldsong")
If Code <> "" Then GetMusicin Code
Else
PO(10).Visible = False
TMACT.Enabled = False
Open App.Path & "\MEDIA\MUSICBOX\经典歌曲.LST" For Input As #1
Do Until EOF(1)
Input #1, a
B_LIST.AddItem a
Loop
Close #1
PO(5).Visible = True
End If
Case 4
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\影视金曲.LST") = 0 Then
PO(10).Visible = True
TMACT.Enabled = True
Code = ReadinteFile("http://music.baidu.com/top/yingshijinqu")
If Code <> "" Then GetMusicin Code
Else
PO(10).Visible = False
TMACT.Enabled = False
Open App.Path & "\MEDIA\MUSICBOX\影视金曲.LST" For Input As #1
Do Until EOF(1)
Input #1, a
B_LIST.AddItem a
Loop
Close #1
PO(5).Visible = True
End If
End Select
LA(1).Caption = ICT(Index).MYTIP
End Sub

Private Sub ILIST_DBClick()
If ILIST.ListCount = 0 Then Exit Sub
frmma.Wm.URL = frmma.PLIST.URL(ILIST.ListIndex)
frmma.Wm.Controls.Play
End Sub

Private Sub IMFAV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If FAV_IT = True Then
Call FRMFAV.REMOVE_ITEM(SONGNAME)
Else
Call FRMFAV.ADD_ITEM(SONGNAME, frmma.Wm.URL)
End If
End Sub

Private Sub IPLAY_Click(Index As Integer)
Select Case Index
Case 0
With LVIEW.ListItems.Add()
.Text = List2.Text
.SubItems(1) = Now
.Icon = 1
.SmallIcon = 1
End With
Call SAVEFAV
Call LOADFAV
Case 1
Call LOADFAV
PO(12).Visible = True
Case 2
On Error Resume Next
Dim MName As String, aname As String
Dim I As Integer
For I = 0 To S_LIST.ListCount - 1
MusicName = S_LIST.List(I)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
If FindMp3URL(MName, aname) <> "" Then frmma.PLIST.AddItem MName, aname, FindMp3URL(MName, aname), 0
Next
Case 3
LVIEW.ListItems.REMOVE (LVIEW.SelectedItem.Index)
Call SAVEFAV
Call LOADFAV
Case 4
LVIEW.ListItems.Clear
Call SAVEFAV
Call LOADFAV
Case 5
If frmma.Winsock1.State <> 7 Then Call SHOWWRONG("请先登录服务器!", 0)
End Select
End Sub

Private Sub IPP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
If frmma.Wm.URL = "" Then Exit Sub
If PLAYING = True Then
PLAYING = False
frmma.Wm.Controls.pause
Else
PLAYING = True
frmma.Wm.Controls.Play
End If
Call DRAWPLAYER
Case 1
If LOLIPOP = 3 Or LOLIPOP = 1 Or LOLIPOP = 2 Then
Call frmma.NT(2)
ElseIf LOLIPOP = 0 Then
Call frmma.NT(3)
End If

End Select

End Sub

Private Sub IW_Click(Index As Integer)
Select Case Index
Case 0
If TXTSER.Text = "<输入歌手或者歌曲名进行搜索>" Or Trim(TXTSER.Text) = "" Then Exit Sub
Call FIND_IT
End Select
End Sub

Private Sub IW_MOUSEMOVE(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Pser.Visible = False Then Pser.Visible = True
Pser.ZOrder 0
End Sub

Private Sub L_LRC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
If Button = 2 Then Me.PopupMenu Frmm.歌词
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 4
PM.Visible = True
PM.ZOrder 0
Case 8
PO(7).Visible = True
PO(7).ZOrder 0
Case Else
Call CMV(Me)
End Select
End Sub

Private Sub LBALL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LBAUTHOR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LBCOUND_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LBNAME_Change()
LA(3).Left = LBNAME.Left + LBNAME.Width + 10
LA(2).Left = LA(0).Left + LA(0).Width + 10

End Sub

Private Sub LBNAME_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub


Private Sub LBSONG_Change()
'IMFAV.Left = LBSONG.Left + LBSONG.Width + 10
End Sub

Private Sub LBSONG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LSER_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FRMLRC.Move FrmNetMusic.Left + (FrmNetMusic.Width - FRMLRC.Width) / 2, FrmNetMusic.Top + (FrmNetMusic.Height - FRMLRC.Height) / 2
FRMLRC.Show
End Sub

Private Sub LSER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub LST_Click()
Dim filename As String, FILETIT As String, AUTHOR As String, SONGTIT As String, SINGER_LOGO As String
filename = LISTSER.List(LST.ListIndex)
FILETIT = LST.List(LST.ListIndex)
If MMAIN.PathFileExists(filename) = 0 Then
Exit Sub
Else
IMIC.STOP_IT
ID3V1.filename = filename
ID3V1.ReadTag
AUTHOR = ID3V1.tagArtist
SONGTIT = ID3V1.tagTitle
IMIC.SETTXT SONGTIT, AUTHOR
IMIC.MUSIC_URL = filename
SINGER_LOGO = App.Path & "\MEDIA\MUSICPICTURE\" & AUTHOR & ".BMP"
If PathFileExists(SINGER_LOGO) = 1 Then IMIC.SETPIC SINGER_LOGO Else IMIC.Cls
End If
End Sub

Private Sub LST_DBClick()
If MMAIN.PathFileExists(LISTSER.List(LST.ListIndex)) = 0 Then Call SHOWWRONG("文件可能被删除或者被移到其他位置,无法播放", 2): Exit Sub
frmma.PLIST.AddItem LST.List(LST.ListIndex), "", LISTSER.List(LST.ListIndex)
frmma.Wm.URL = LISTSER.List(LST.ListIndex)
frmma.Wm.Controls.Play
Call RELIST
If IS_MINI_LIST = True Then Call FRMLIST.RELIST
End Sub

Private Sub LVIEW_DblClick()
On Error Resume Next
PO(12).Visible = False
Call SEARCH_ABOUT_TA(LVIEW.SelectedItem.Text)
End Sub

Private Sub LVIEW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub optThumb_Click(Index As Integer)
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then Call SHOWWRONG("未检测到活动网络,请重试", 2): Exit Sub
Call SEARCH_ABOUT_TA(optThumb(Index).MY_TIT)
Debug.Print optThumb(Index).MY_TIT
End Sub
Sub SEARCH_ABOUT_TA(TA As String)
On Error Resume Next
S_LIST.ListIndex = 0
List2.Text = TA
S_LIST.Clear
Dim STRTMB As String
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\" & TA & ".SI") = 0 Then
For I = 0 To 180 Step 20
Call searchsinger(TA, I, S_LIST)
Next
Call SAVE_SONG_LIST(TA, S_LIST) '保存以歌手为名的播放列表
GoTo ERR
Else
Open App.Path & "\MEDIA\MUSICBOX\" & TA & ".SI" For Input As #1
Do Until EOF(1)
Input #1, STRTMB
If Len(STRTMB) < 200 Then S_LIST.AddItem STRTMB
Loop
Close #1
GoTo ERR
End If
ERR:
Call 清除重复(S_LIST)
PO(6).Visible = True
PICFRAME.Visible = False
PLEFT.Visible = False
LBNAME.Caption = TA
PO(10).Visible = False
TMACT.Enabled = False
End Sub
Private Sub optThumb_MOUSEMOVE(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub KPB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISMV = False Then
ISMV = True
KPB.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", KPB.hdc, 0, 0)
End If
End Sub

Private Sub KPB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PO(7).Visible = False
End Sub

Private Sub PBK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISMV = False Then
ISMV = True
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub PBK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PO(12).Visible = False
End Sub
Private Sub BKP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISMV = False Then
ISMV = True
BKP.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", BKP.hdc, 0, 0)
End If
End Sub

Private Sub BKP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PO(6).Visible = False
PLEFT.Visible = True
PICFRAME.Visible = True
End Sub

Private Sub PC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Call SETLRCCOLOR(Index + 1)
cDeskLrc.ReDraw
lRet = SetInitEntry("PLAYER", "LRCSHOW_COLOR", Index + 1)
End Sub

Private Sub PF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PO(7).Visible = True: PO(7).ZOrder 0
End Sub

Private Sub picFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub picFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICSET_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISMV = False Then
ISMV = True
PICSET.Cls
Call PaintPng(App.Path & "\SKIN\SET_H.PNG", PICSET.hdc, 0, 0)
End If
End Sub

Private Sub PICSET_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PICLRC.Visible = True
End Sub

Private Sub PICSLIDE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub picSlide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PKB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISMV = False Then
ISMV = True
PKB.Cls

Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PKB.hdc, 0, 0)
End If
End Sub

Private Sub PKB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PO(5).Visible = False
End Sub

Private Sub PLEFT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PO(4).Visible = True: PO(4).ZOrder 0
End Sub

Private Sub PO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 11 Then
If frmma.Wm.playState = wmppsPlaying Then frmma.Wm.Controls.currentPosition = X * frmma.Wm.currentMedia.duration / PO(11).ScaleWidth
End If
If Index <> 11 Then Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> 8 Then MOVENOW
End Sub
Private Sub Form_Resize()
Dim X       As Long
Dim Y       As Long
Dim lIdx    As Long
Dim lCols   As Long
PICFRAME.Move 0, 0, PO(2).ScaleWidth, PO(2).ScaleHeight - PLEFT.Height
            vsbSlide.Move PICFRAME.ScaleWidth - vsbSlide.Width, 0, vsbSlide.Width, PICFRAME.ScaleHeight
            lCols = Int((PICFRAME.ScaleWidth) / optThumb(0).Width)
            For lIdx = 0 To optThumb.Count - 1
                X = (lIdx Mod lCols) * optThumb(0).Width
                Y = Int(lIdx / lCols) * optThumb(0).Height
                optThumb(lIdx).Move X, Y
            Next lIdx
            picSlide.Width = lCols * optThumb(0).Width
            picSlide.Height = Int(optThumb.Count / lCols) * optThumb(0).Height
            If Int(optThumb.Count / lCols) < (optThumb.Count / lCols) Then
                picSlide.Height = picSlide.Height + optThumb(0).Height
            End If
            vsbSlide.Value = 0
            vsbSlide.Max = picSlide.Height - PICFRAME.ScaleHeight
            If vsbSlide.Max < 0 Then
                vsbSlide.Max = 0
            Else
                vsbSlide.SmallChange = optThumb(0).Height
                vsbSlide.LargeChange = PICFRAME.ScaleHeight
            End If
            PTOP.Move PICFRAME.ScaleWidth - PTOP.Width - 50, PICFRAME.ScaleHeight - PTOP.Height - 19
            PTOP.ZOrder 0
End Sub
Private Sub PTOP_Click()
tmrCheck.Enabled = True
End Sub

Private Sub PV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR
If ISMOVE = True Then
SV.Width = X
If SV.Width <= 0 Then SV.Width = 0: Exit Sub
If SV.Width >= PV.ScaleWidth Then SV.Width = PV.ScaleWidth: ISMOVE = False
frmma.Wm.settings.volume = Int((100 / PV.ScaleWidth) * SV.Width)
End If
ERR:
End Sub

Private Sub PV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ISMOVE = False
lRet = SetInitEntry("PLAYER", "VOLUME", SV.Width)
End Sub
Private Sub PV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ISMOVE = True

End Sub

Private Sub S_LIST_Click()
On Error Resume Next
Dim MName As String, aname As String
MusicName = S_LIST.List(S_LIST.ListIndex)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
Will_DL = FindMp3URL(MName, aname)
M_N = MName
A_N = aname
End Sub

Private Sub S_LIST_DBClick()
On Error Resume Next
Dim MName As String, aname As String, WILL_PLAY As String
MusicName = S_LIST.List(S_LIST.ListIndex)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
WILL_PLAY = FindMp3URL(MName, aname)
If WILL_PLAY = "" Then Exit Sub
frmma.PLIST.AddItem MName, aname, WILL_PLAY, 0
frmma.PLIST.ListIndex = frmma.PLIST.ListCount - 1
frmma.PLIST.PlayIndex = frmma.PLIST.ListCount - 1
frmma.Wm.URL = WILL_PLAY
frmma.PLAYB.Enabled = True
专辑 = ""
年代 = ""
Call RELIST
If IS_MINI_LIST = True Then Call FRMLIST.RELIST
End Sub

Private Sub S_LIST_MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub Timer1_Timer()
If PO(1).Left = 0 Then PO(1).Left = -PO(1).Width
PO(3).Left = PO(3).Left - 80
PO(2).Left = PO(3).Left + PO(3).Width
If PO(3).Left < -PO(3).Width Then
PO(3).Left = -PO(3).Width
PO(2).Left = 0
Timer1.Enabled = False
PO(1).Left = -PO(1).Width
End If
End Sub

Private Sub Timer2_Timer()
PO(3).Left = PO(3).Left + 80
PO(2).Left = PO(3).Left + PO(3).Width
PO(1).Left = PO(3).Left - PO(1).Width
If PO(3).Left > 0 Then
PO(3).Left = 0
Timer2.Enabled = False
PO(1).Left = PO(3).Left - PO(1).Width
End If
End Sub

Private Sub Timer3_Timer()
PO(3).Left = PO(3).Left + 80
PO(2).Left = PO(3).Left + PO(3).Width
PO(1).Left = PO(3).Left - PO(1).Width
If PO(3).Left > Me.ScaleWidth Then
PO(3).Left = Me.ScaleWidth
PO(1).Left = 0
Timer3.Enabled = False
End If

End Sub

Private Sub TMP_Timer()
On Error Resume Next
PF.SETTXT "本地歌曲 " & Format(LST.ListCount, "000")
If LST.ListCount <> 0 Then LA(7).Visible = False Else LA(7).Visible = True
If frmma.Wm.currentMedia.duration > 0 Then PR.Width = PO(11).ScaleWidth / frmma.Wm.currentMedia.duration * frmma.Wm.Controls.currentPosition   '如果进度条未被拖动 则计算出播放进度位置 并移动
PO(11).ToolTipText = frmma.Wm.Controls.currentPositionString & "/" & frmma.Wm.currentMedia.durationString
If frmma.Wm.URL <> "" Then LBCOUND.Caption = frmma.Wm.Controls.currentPositionString Else LBCOUND.Caption = "00:00": PR.Width = 0
LA(3).Caption = S_LIST.ListCount
PM.SETTXT "默认列表 " & Format(ILIST.ListCount, "000")
LBSONG.Caption = SONGNAME
LBAUTHOR.Caption = frmma.FILESINGER
SONGNAME = frmma.Wm.currentMedia.name
If SONGNAME = "" Then LBSONG.Caption = "还没有播放歌曲"
SV.Width = (PV.ScaleWidth * frmma.Wm.settings.volume) / 100
LBALL.Caption = frmma.Wm.currentMedia.durationString
Select Case frmma.Wm.playState
Case wmppsPaused
PLAYING = False
Me.Caption = "暂停播放 -" & SONGNAME
Case wmppsPlaying
PLAYING = True
Me.Caption = "正在播放 -" & SONGNAME
Case 3
Me.Caption = "ICEE音乐 "
PLAYING = False
End Select

'Dim p As POINTAPI, f As RECT
 '   GetCursorPos p '得到MOUSE位置
 '   GetWindowRect Me.hwnd, f '得到窗体的位置
 '   If Me.WindowState <> 1 Then
 '       If p.X > f.Left And p.X < f.Right And p.Y > f.Top And p.Y < f.Bottom Then
'            'MOUSE 在窗体上
'            If Me.Top < 0 Then
'                Me.Top = -10
'                Me.Show
'            ElseIf Me.Left < 0 Then
'                Me.Left = -10
'                Me.Show
'            ElseIf Me.Left + Me.Width >= Screen.Width Then
'                Me.Left = Screen.Width - Me.Width + 10
'                Me.Show
'            End If
'        Else
'            If f.Top <= 4 Then
'                Me.Top = 40 - Me.Height
'            ElseIf f.Left <= 4 Then
'                Me.Left = 40 - Me.Width
'            ElseIf Me.Left + Me.Width >= Screen.Width - 4 Then
'                Me.Left = Screen.Width - 40
'            End If
'        End If
'    End If
End Sub

Private Sub TMACT_Timer()
PO(10).ZOrder 0
L_TIME = L_TIME + 1
PO(10).Cls
Call PaintPng(App.Path & "\SKIN\M_LOAD.PNG", PO(10).hdc, (PO(10).ScaleWidth - 309) / 2, PO(10).ScaleHeight - 150)
Call PaintPng(App.Path & "\SKIN\AB" & L_TIME & ".png", PO(10).hdc, (PO(10).ScaleWidth - 83) / 2, (PO(10).ScaleHeight - 160) / 2)
If L_TIME > 4 Then L_TIME = 0
End Sub


Private Sub TMSEEK_Timer()
L_LRC.SeekLrc frmma.Wm.Controls.currentPosition
If D_L_SHOW = True Then cDeskLrc.SeekLrc frmma.Wm.Controls.currentPosition, False
End Sub

Private Sub TXTSER_GotFocus()
If TXTSER.Text = "<输入歌手或者歌曲名进行搜索>" Then TXTSER.Text = ""
TXTSER.SelStart = 0
TXTSER.SelLength = Len(TXTSER.Text)
If F_LIST.ListCount > 0 Then PO(9).Visible = True
End Sub

Private Sub TXTSER_KeyPress(KeyAscii As Integer)
If TXTSER.Text = "<输入歌手或者歌曲名进行搜索>" Or Trim(TXTSER.Text) = "" Then Exit Sub
If KeyAscii = 13 Then Call FIND_IT
End Sub
Sub FIND_IT()
Dim STRTMB As String
Dim SC_SINGER As String
SC_SINGER = TXTSER.Text
F_LIST.Clear
PO(9).Visible = False
If PathFileExists(App.Path & "\MEDIA\MUSICBOX\" & SC_SINGER & ".SI") = 0 Then
PO(10).Visible = True
TMACT.Enabled = True
For I = 0 To 180 Step 20
If ICSER(0).IS_SELECT = True Then
Call searchsinger(SC_SINGER, I, F_LIST)
Else
Call searchmusic(SC_SINGER, I, F_LIST)
End If
Next
Call SAVE_SONG_LIST(SC_SINGER, F_LIST) '保存以歌手为名的播放列表
Else
Open App.Path & "\MEDIA\MUSICBOX\" & SC_SINGER & ".SI" For Input As #1
Do Until EOF(1)
Input #1, STRTMB
If Len(STRTMB) < 200 Then F_LIST.AddItem STRTMB
Loop
Close #1
End If
If F_LIST.ListCount = 0 Then
Call SHOWWRONG("木有找到" & SC_SINGER & "的相关歌曲呦,亲", 2)
PO(9).Visible = False
Else
PO(9).Visible = True
End If
Call 清除重复(F_LIST)
TMACT.Enabled = False
PO(10).Visible = False
End Sub

Private Sub TXTSER_LostFocus()
If Trim(TXTSER.Text) = "" Then TXTSER.Text = "<输入歌手或者歌曲名进行搜索>"
End Sub

Private Sub TXTSER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PO(9).ZOrder 0
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
If X3.Visible = False Then IStart = 0: Me.Hide
End Sub

'根据歌手搜歌
Public Function searchsinger(ByVal SINGER As String, ByVal START As Integer, List As ILIST)
On Error Resume Next
Dim IEnd As Long
TXTTEST.Text = ""
PO(10).Visible = True
TMACT.Enabled = True
strURL = "http://music.baidu.com/search.key=" & UTF8EncodeURI(SINGER) & "&start=" & START & "&size=20"
strCode = ReadinteFile(strURL)
TXTTEST.Text = strCode
If Len(strCode) > 100 Then
IEnd = 1
Do
IStart = InStr(IEnd + 800, strCode, "author_list")
If IStart = 0 Then Exit Do
IStart = InStr(IStart, strCode, "title=", vbTextCompare)
IEnd = InStr(IStart, strCode, ">")
SINGERNAME = Mid$(strCode, IStart + 7, IEnd - IStart - 8)
IStart = InStr(IEnd - 700, strCode, "data-songdata=")
IStart = InStr(IStart, strCode, "title=")
IEnd = InStr(IStart, strCode, ">")
Music = Mid$(strCode, IStart + 7, IEnd - IStart - 8)
If InStr(1, SINGERNAME, SINGER) > 0 And InStr(1, Music, "审批") < 1 <> "" And SINGERNAME <> "" Then
If Len(Music) <= 200 Then List.AddItem SINGERNAME & " - " & Music
End If
Loop
End If

PO(10).Visible = False
TMACT.Enabled = False

End Function
Sub SAVE_SONG_LIST(TITTLE As String, LISTBOX As ILIST)
If LISTBOX.ListCount = 0 Then Exit Sub
Dim sFile As String
Dim I As Integer
sFile = App.Path & "\MEDIA\MUSICBOX\" & TITTLE & ".SI"

Open sFile For Output As #1
For I = 0 To LISTBOX.ListCount - 1
If Len(Trim(LISTBOX.List(I))) < 200 Then Print #1, LISTBOX.List(I)
Next I
Close #1
End Sub

'根据歌曲搜歌
Public Function searchmusic(ByVal Word As String, ByVal START As Integer, List As LISTBOX)
strURL = "http://music.baidu.com/search.key=" & UTF8EncodeURI(Word) & "&start=" & START & "&size=20"
strCode = ReadinteFile(strURL)
If Len(strCode) > 100 Then
IEnd = 1
Do
IStart = InStr(IEnd, strCode, "data-songdata=")
If IStart = 0 Then Exit Do
IStart = InStr(IStart, strCode, "title=")
IEnd = InStr(IStart, strCode, ">")
Music = Mid$(strCode, IStart + 7, IEnd - IStart - 8)
IStart = InStr(IEnd, strCode, "author_list")
IStart = InStr(IStart, strCode, "title=")
IEnd = InStr(IStart, strCode, ">")
SINGERNAME = Mid$(strCode, IStart + 7, IEnd - IStart - 8)
If InStr(1, Music, "审批") < 1 Then
If SINGERNAME <> "" Then List.AddItem SINGERNAME + " - " + Music
End If
Loop
PO(10).Visible = False
TMACT.Enabled = False
End If
End Function

Private Sub SINGER()
List2.Clear
Dim STRTMB As String
strURL = "http://music.baidu.com/artist"
strCode = ReadinteFile(strURL)
IEnd = InStr(1, strCode, "其他</a>")
Do While InStr(1, Left(strCode, IEnd), "music-foot-alog") < 1
DoEvents
n = InStr(IEnd, strCode, "<a href=")
IStart = InStr(n + 10, strCode, ">")
IEnd = InStr(n + 10, strCode, "</a>")
If Mid(strCode, IStart + 1, IEnd - IStart - 1) <> "" Then List2.AddItem Mid(strCode, IStart + 1, IEnd - IStart - 1)
Loop
Call CreateThumbs
End Sub
Sub SAVE_SINGER_LIST()
On Error Resume Next
Dim a As Integer
If List2.ListCount = 0 Then Call SHOWWRONG("获取歌手列表失败,请重启程序或检查网络连接", 0): Exit Sub
Open App.Path & "\MEDIA\MUSICBOX\SINGER_LIST.DAT" For Output As #1
For a = 0 To List2.ListCount - 1 '保存搜索结果
If Trim(List2.List(a)) <> "" Then Write #1, List2.List(a)
Next a
Close #1
End Sub
Public Function GetMusicin(ByVal Code As String)
On Error Resume Next
Dim lngStart As Long, lngEnd As Long, sTemp As String, filename As String
Dim s As Long, E As Long, sName As String, sArt As String
Dim MName As String, aname As String
B_LIST.Clear
lngEnd = 1
List4.Clear
Do
DoEvents
lngStart = InStr(lngEnd, Code, "<li  data-songitem =") + 20
If lngStart = 20 Then Exit Do
lngEnd = InStr(lngStart, Code, ">")
sTemp = Mid$(Code, lngStart, lngEnd - lngStart)
s = InStr(1, sTemp, "'sname': '") + 10
E = InStr(s, sTemp, "'")
sName = Mid$(sTemp, s, E - s)
s = InStr(1, sTemp, "'author': '") + 11
E = InStr(s, sTemp, "'")
sArt = Mid$(sTemp, s, E - s)
List4.AddItem sArt & " - " & sName
Loop
For I = 0 To List4.ListCount - 1
B_LIST.AddItem List4.List(I)
Next
If List4.ListCount = 0 Then
PO(5).Visible = False
Else
PO(5).Visible = True
End If
B_LIST.ListIndex = 0
PO(10).Visible = False
TMACT.Enabled = False

Select Case WIL_B
Case 0
filename = App.Path & "\MEDIA\MUSICBOX\网络歌曲.LST"
Case 1
filename = App.Path & "\MEDIA\MUSICBOX\百度热歌.LST"
Case 2
filename = App.Path & "\MEDIA\MUSICBOX\百度新歌.LST"
Case 3
filename = App.Path & "\MEDIA\MUSICBOX\经典歌曲.LST"
Case 4
filename = App.Path & "\MEDIA\MUSICBOX\影视金曲.LST"
End Select
If List4.ListCount = 0 Then Exit Function
Open filename For Output As #1
For I = 0 To List4.ListCount - 1
Print #1, List4.List(I)
Next I
Close #1

End Function
Private Sub c_Subclass_GetWindowMessage(Result As Long, ByVal chwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case Message
        Case WM_NCLBUTTONDOWN
            Const HTCAPTION = 2
            If wParam = HTCAPTION Then
                '点击标题栏让所有Timer停止工作
                m_MoveEnabled = True
                tmrCheck.Enabled = False
                tmrMove.Enabled = False
            End If
            
        Case WM_MOVING
            If m_MoveEnabled = False Then Exit Sub
            '这里仅仅是为了不让窗口移出屏幕，可以忽略
            Dim rcMov   As RECT
            Dim rcWnd   As RECT
            Dim lScrW   As Long
            '获取窗口矩形
            Call GetWindowRect(chwnd, rcWnd)
            '//屏幕宽度
            lScrW = Screen.Width / Screen.TwipsPerPixelX
            '获取移动目标位置矩形
            Call CopyMemory(rcMov, ByVal lParam, Len(rcMov))
            With rcMov
                If .Left < 0 Then
                    .Left = 0
                    .Right = rcWnd.Right - rcWnd.Left
                End If
                If .Top < 0 Then
                    .Top = 0
                    .Bottom = rcWnd.Bottom - rcWnd.Top
                End If
                If .Right > lScrW Then
                    .Left = lScrW - (rcWnd.Right - rcWnd.Left)
                    .Right = .Left + (rcWnd.Right - rcWnd.Left)
                End If
            End With
            '//如果窗口的靠在右上角或左上角，则把高度设置为屏幕高度
            Call CopyMemory(ByVal lParam, rcMov, Len(rcMov))
            
        Case WM_EXITSIZEMOVE
            m_MoveEnabled = False
            Call GetWindowRect(chwnd, rcWnd)
            If rcWnd.Left <= 0 Or rcWnd.Top <= 0 Or _
                rcWnd.Right >= Screen.Width / Screen.TwipsPerPixelX Then
                '如果窗口停靠在屏幕边缘
                '让检查鼠标位置的Timer工作
                
                '设置显示方向
                If rcWnd.Left = 0 Then
                    m_ShowOrient = 0
                ElseIf rcWnd.Right >= Screen.Width / Screen.TwipsPerPixelX Then
                    m_ShowOrient = 1
                ElseIf rcWnd.Top = 0 Then
                    m_ShowOrient = 2
                End If
                tmrCheck.Enabled = True
            End If
    End Select
    Result = c_Subclass.CallDefaultWindowProc(chwnd, Message, wParam, lParam)
End Sub

Private Sub tmrCheck_Timer()
Me.vsbSlide.Value = vsbSlide.Value - 250
If vsbSlide.Value = 0 Then tmrCheck.Enabled = False
'    Dim PT As POINTAPI
'    Dim rc As RECT
'    Call GetCursorPos(PT)
'    Call GetWindowRect(Me.hwnd, rc)
'    If PtInRect(rc, PT.X, PT.Y) Then
'        '鼠标停留在窗口上
'        If m_ShowFlag = 1 Then Exit Sub
'        m_ShowSpeed = SHOWHIDE_SPEED
'        m_ShowFlag = 1
'        tmrMove.Enabled = True
'    Else
'        '鼠标不再窗口上
'        If m_ShowFlag = 0 Then Exit Sub
'        m_ShowSpeed = SHOWHIDE_SPEED
'        m_ShowFlag = 0
'        tmrMove.Enabled = True
'    End If
End Sub

Private Sub tmrMove_Timer()
'    Dim nTop    As Long
'    Dim nLeft   As Long
'    m_ShowSpeed = m_ShowSpeed + SHOWHIDE_SPEED
'    '如果大于300T则加快速度
'    If m_ShowSpeed > 300 Then m_ShowSpeed = m_ShowSpeed + m_ShowSpeed * 0.2
'    Select Case m_ShowOrient
'        Case 0  '0  向左
'        '    If m_ShowFlag = 0 Then
        '        nLeft = Me.Left - m_ShowSpeed
        '        If nLeft < -Me.Width + SIZE_SHOW Then nLeft = -Me.Width + SIZE_SHOW: tmrMove.Enabled = False
        '    Else
        '        nLeft = Me.Left + m_ShowSpeed
        '        If nLeft > -SIZE_SHOW Then nLeft = -SIZE_SHOW: tmrMove.Enabled = False
        '    End If
        '    Me.Left = nLeft
            
'        Case 1  '1  向右
        '    If m_ShowFlag = 0 Then
        '        nLeft = Me.Left + m_ShowSpeed
        '        If nLeft > Screen.Width - SIZE_SHOW Then nLeft = Screen.Width - SIZE_SHOW: tmrMove.Enabled = False
        '    Else
        '        nLeft = Me.Left - m_ShowSpeed
        '        If nLeft < Screen.Width - Me.Width + SIZE_SHOW Then nLeft = Screen.Width - Me.Width + SIZE_SHOW: tmrMove.Enabled = False
        '    End If
        '    Me.Left = nLeft
            
'        Case 2  '2  向上
'            If m_ShowFlag = 0 Then
'                nTop = Me.Top - m_ShowSpeed
'                If nTop < -Me.Height + SIZE_SHOW Then nTop = -Me.Height + SIZE_SHOW: tmrMove.Enabled = False
'            Else
'                nTop = Me.Top + m_ShowSpeed
'                If nTop > -SIZE_SHOW Then nTop = -SIZE_SHOW: tmrMove.Enabled = False
'            End If
'            Me.Top = nTop
'
'    End Select
End Sub


Sub MOVENOW()
If ISMV = True Then
ISMV = False
PICSET.Cls
PBK.Cls
PKB.Cls
BKP.Cls
KPB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", BKP.hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", KPB.hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\SET_N.PNG", PICSET.hdc, 0, 0)
End If
If Pser.Visible = True Then Pser.Visible = False
X1.Visible = True
X2.Visible = False
X3.Visible = False
If PO(9).Visible = True Then PO(9).Visible = False
If PO(4).Visible = True Then PO(4).Visible = False
If PICLRC.Visible = True Then PICLRC.Visible = False
For I = 0 To PIC_LRC.Count - 1
If PIC_LRC(I).Visible = True Then PIC_LRC(I).Visible = False
Next
End Sub

Private Sub CreateThumbs()
Dim LAST_PAGE As Integer, I As Integer
PO(10).Visible = True
LISTFL.Clear
TMACT.Enabled = True
Call SAVE_SINGER_LIST
LAST_PAGE = GetInitEntry("MUSICBOX", "LAST_PAGE", 0)
If LAST_PAGE < 26 Then
For I = 0 To List2.ListCount - 1
If py(List2.List(I)) = Mid(ICC(LAST_PAGE).MY_TIT, 1, 1) Or py(List2.List(I)) = Mid(ICC(LAST_PAGE).MY_TIT, 2, 1) Or py(List2.List(I)) = Mid(ICC(LAST_PAGE).MY_TIT, 3, 1) Then LISTFL.AddItem List2.List(I)
LA(10).Caption = "正在分析列表"
Next
Else
For I = 0 To List2.ListCount - 1
LISTFL.AddItem List2.List(I)
If UCase(Left(List2.List(I), 1)) = "A" Then Exit For
LA(10).Caption = "正在分析列表"
Next
End If
LA(10).Caption = "分析完成,开始加载控件"
For I = 0 To ICC.Count - 1
ICC(I).IS_SELECT = False
Next
ICC(LAST_PAGE).IS_SELECT = True
If LISTFL.ListCount = 0 Then PO(10).Visible = False: TMACT.Enabled = False: Exit Sub
Dim lIdx    As Long
    picSlide.Move 0, 0, optThumb(0).Width, optThumb(0).Height
    picSlide.Visible = False
    While optThumb.Count > 1
        Unload optThumb(optThumb.Count - 1)
    Wend
    DoEvents
        For lIdx = 0 To LISTFL.ListCount - 1
            ERR.Clear
                If lPicCnt > 0 Then
                    Load optThumb(lPicCnt)
                    Set optThumb(lPicCnt).Container = picSlide
                End If
                LA(10).Caption = "正在加载控件"
                DoEvents
                optThumb(lPicCnt).L_M_R = 1
                optThumb(lPicCnt).SETTXT LISTFL.List(lIdx)
                optThumb(lPicCnt).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
                optThumb(lPicCnt).Visible = True
                lPicCnt = lPicCnt + 1
        Next lIdx
        PO(10).Visible = False
        TMACT.Enabled = False
        Call Form_Resize
        picSlide.Visible = True
        LBS.Caption = LISTFL.ListCount
        LA(10).Caption = "加载完成"
End Sub

Private Sub vsbSlide_Change()
    picSlide.Top = -vsbSlide.Value
    If vsbSlide.Value = 0 Then PTOP.Visible = False Else PTOP.Visible = True
End Sub
Private Sub vsbSlide_Scroll()
    vsbSlide_Change
End Sub
Public Function py(mystr As String) As String
  If Asc(mystr) < 0 Then
    If Asc(Left(mystr, 1)) < Asc("啊") Then
       py = "0"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("啊") And Asc(Left(mystr, 1)) < Asc("芭") Then
       py = "A"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("芭") And Asc(Left(mystr, 1)) < Asc("擦") Then
       py = "B"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("擦") And Asc(Left(mystr, 1)) < Asc("搭") Then
       py = "C"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("搭") And Asc(Left(mystr, 1)) < Asc("蛾") Then
       py = "D"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("蛾") And Asc(Left(mystr, 1)) < Asc("发") Then
       py = "E"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("发") And Asc(Left(mystr, 1)) < Asc("噶") Then
       py = "F"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("噶") And Asc(Left(mystr, 1)) < Asc("哈") Then
       py = "G"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("哈") And Asc(Left(mystr, 1)) < Asc("击") Then
       py = "H"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("击") And Asc(Left(mystr, 1)) < Asc("喀") Then
       py = "J"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("喀") And Asc(Left(mystr, 1)) < Asc("垃") Then
       py = "K"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("垃") And Asc(Left(mystr, 1)) < Asc("妈") Then
       py = "L"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("妈") And Asc(Left(mystr, 1)) < Asc("拿") Then
       py = "M"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("拿") And Asc(Left(mystr, 1)) < Asc("哦") Then
       py = "N"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("哦") And Asc(Left(mystr, 1)) < Asc("啪") Then
       py = "O"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("啪") And Asc(Left(mystr, 1)) < Asc("期") Then
       py = "P"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("期") And Asc(Left(mystr, 1)) < Asc("然") Then
       py = "Q"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("然") And Asc(Left(mystr, 1)) < Asc("撒") Then
       py = "R"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("撒") And Asc(Left(mystr, 1)) < Asc("塌") Then
       py = "S"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("塌") And Asc(Left(mystr, 1)) < Asc("挖") Then
       py = "T"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("挖") And Asc(Left(mystr, 1)) < Asc("昔") Then
       py = "W"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("昔") And Asc(Left(mystr, 1)) < Asc("压") Then
       py = "X"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("压") And Asc(Left(mystr, 1)) < Asc("匝") Then
       py = "Y"
       Exit Function
    End If
    If Asc(Left(mystr, 1)) >= Asc("匝") Then
       py = "Z"
       Exit Function
    End If
  Else
    If UCase(mystr) <= "Z" And UCase(mystr) >= "A" Then
       py = UCase(Left(mystr, 1))
      Else
       py = mystr
    End If
  End If
End Function
Sub DRAWPLAYER()
PO(1).Cls
Call PaintPng(App.Path & "\SKIN\NEXT.png", PO(1).hdc, IPP(1).Left, IPP(1).Top)
If frmma.Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PAUSE.png", PO(1).hdc, IPP(0).Left, IPP(0).Top) Else Call PaintPng(App.Path & "\SKIN\PLAY.png", PO(1).hdc, IPP(0).Left, IPP(0).Top)
Call PaintPng(App.Path & "\SKIN\VOL.png", PO(1).hdc, 344, 428)
L_LRC.SETPIC PO(1), L_LRC.Left, L_LRC.Top
End Sub
Sub REDRAWPLAYER()
Call PaintPng(App.Path & "\SKIN\NEXT.png", PO(1).hdc, IPP(1).Left, IPP(1).Top)
If frmma.Wm.playState = wmppsPlaying Then Call PaintPng(App.Path & "\SKIN\PAUSE.png", PO(1).hdc, IPP(0).Left, IPP(0).Top) Else Call PaintPng(App.Path & "\SKIN\PLAY.png", PO(1).hdc, IPP(0).Left, IPP(0).Top)
Call PaintPng(App.Path & "\SKIN\VOL.png", PO(1).hdc, 728, 440)
L_LRC.SETPIC PO(1), L_LRC.Left, L_LRC.Top
End Sub
Private Sub L_LRC_MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
For I = 0 To PIC_LRC.Count - 1
If PIC_LRC(I).Visible = False Then PIC_LRC(I).Visible = True
Next
If PICLRC.Visible = True Then PICLRC.Visible = False

End Sub

Private Sub PIC_LRC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
L_LRC.SETFONTSIZE (L_LRC.DE_SIZE + 1)
Case 1
L_LRC.SETFONTSIZE (L_LRC.DE_SIZE - 1)
End Select
End Sub
Private Sub TMLRC_Timer()
If SONGNAME = "" Then L_LRC.Visible = False: Exit Sub
Path = App.Path & "\MEDIA\LRC\" & SONGNAME & ".lrc"
'思路是线检查本地是否有文件,没有的话如果自动搜索值为1且联网的话搜索，否则设置默认信息
If PathFileExists(Path) <> 0 Then
L_LRC.ReadFile (Path)
L_LRC.Visible = True
TMLRC.Enabled = False
TMSEEK.Enabled = True
Else
L_LRC.Visible = False
TMSEEK.Enabled = False
L_LRC.ClearLrc
End If
End Sub
Private Sub SearchDirs(curpath$)
    Dim dirs%, dirbuf$(), I%
    PO(13).Cls
    PO(13).CurrentY = 3
    PO(13).Print "搜索中 " & curpath$
    DoEvents
    If Not Running% Then Exit Sub
    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    TotalDirs% = TotalDirs% + 1
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            ElseIf Not UseFileSpec% Then
                TotalFiles% = TotalFiles% + 1
            End If
        Loop While FindNextFile(hItem&, WFD)
        Call FindClose(hItem&)
    End If
    If UseFileSpec% Then
        SendMessage hLB&, WM_SETREDRAW, 0, 0
        Call SearchFileSpec(curpath$)
        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
        SendMessage hLB&, WM_SETREDRAW, 1, 0
    End If
    
    For I% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(I%) & vbBackslash: Next I%
  
End Sub

Private Sub SearchFileSpec(curpath$)   ' curpath$ is passed w/ trailing "\"
    hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If Not Running% Then Exit Sub
            SendMessage hLB&, LB_ADDSTRING, 0, _
                ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        Loop While FindNextFile(hFile&, WFD)
        Call FindClose(hFile&)
    End If
End Sub

Private Sub START_S()
    If Running% Then: Running% = False: Exit Sub
    Dim drvbitmask&, maxpwr%, pwr%
    FileSpec$ = "*.mp3"
    Running% = True
    UseFileSpec% = True
    LISTSER.Clear
    LST.Clear
    IMIC.Visible = False
    LST.Visible = False
    drvbitmask& = GetLogicalDrives()
    If drvbitmask& Then
        maxpwr% = Int(Log(drvbitmask&) / Log(2))   ' a little math...
        For pwr% = 0 To maxpwr%
            If Running% And (2 ^ pwr% And drvbitmask&) Then _
                Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
        Next
    End If
    Running% = False
    UseFileSpec% = False
    LST.Visible = True
    IMIC.Visible = True
    PO(13).Cls
    PO(13).Print "找到 " & LISTSER.ListCount & "首歌曲"
    Dim I As Integer
    For I = 0 To LISTSER.ListCount - 1
    LST.AddItem LastFileName(LISTSER.List(I))
    Next
On Error Resume Next
Dim a As Integer
Open App.Path & "\MEDIA\SEARCH\MP3.CP" For Output As #1
For a = 0 To LISTSER.ListCount - 1 '保存搜索结果
 Write #1, LISTSER.List(a)
Next a
Close #1
End Sub

Sub CHECK_FAV(name As String)
Dim I As Integer
If LVIEW.ListItems.Count = 0 Then Exit Sub
For I = 1 To LVIEW.ListItems.Count
If LVIEW.ListItems(I).Text = name Then FAV_ART = True Else FAV_ART = False
Next
Call DRAWPLAYER
End Sub
Sub SETLRC()
On Error Resume Next
Dim SIT As Integer
SIT = GetInitEntry("PLAYER", "LRCSHOW_COLOR", 0)
Call SETLRCCOLOR(SIT + 1)
cDeskLrc.ShowText "ICEE音乐,音乐您的生活"
cDeskLrc.FontName = "微软雅黑"
cDeskLrc.FontBold = True
cDeskLrc.Karaoke = True
cDeskLrc.FontSize = 20
cDeskLrc.ReDraw
FRMSHOW.Show
End Sub
Sub SETLRCCOLOR(ByVal Mode As Integer)
    Select Case Mode
        Case 1          '蓝色
            cDeskLrc.BackColor1 = &HFF013C8F
            cDeskLrc.BackColor2 = &HFF0198D4
            cDeskLrc.ForeColor1 = &HFFBCF9FC
            cDeskLrc.ForeColor2 = &HFF67F0FC
            cDeskLrc.LineColor = &H30000000
        Case 2          '绿色
            cDeskLrc.BackColor1 = &HFF87F321
            cDeskLrc.BackColor2 = &HFF0E6700
            cDeskLrc.ForeColor1 = &HFFDCFEAE
            cDeskLrc.ForeColor2 = &HFFE4FE04
            cDeskLrc.LineColor = &H30000000
        Case 3          '红色
            cDeskLrc.BackColor1 = &HFFFECEFC
            cDeskLrc.BackColor2 = &HFFE144CD
            cDeskLrc.ForeColor1 = &HFFFEFE65
            cDeskLrc.LineColor = &H30000000
        Case 4          '白色
            cDeskLrc.BackColor1 = &HFFFBFBFA
            cDeskLrc.BackColor2 = &HFFCBCBCB
            cDeskLrc.ForeColor1 = &HFF62DDFF
            cDeskLrc.ForeColor2 = &HFF229CFE
            cDeskLrc.LineColor = &H30000000
        Case 5          '黄色
            cDeskLrc.BackColor1 = &HFFFE7A00
            cDeskLrc.BackColor2 = &HFFFF0000
            cDeskLrc.ForeColor1 = &HFFFFFF6E
            cDeskLrc.ForeColor2 = &HFFFEA10F
            cDeskLrc.LineColor = &H30000000
    End Select
End Sub

Sub 清除重复(PLIST As ILIST)
On Error Resume Next
Dim n As Integer, m As Integer
For n = 0 To PLIST.ListCount - 1
For m = n To PLIST.ListCount - 1
If PLIST.List(n) = PLIST.List(m) And m <> n Then PLIST.RemoveItem m
Next
Next
End Sub
