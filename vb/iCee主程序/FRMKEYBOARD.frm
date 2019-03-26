VERSION 5.00
Begin VB.Form FRMKEYBOARD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H001F1F1F&
   BorderStyle     =   0  'None
   Caption         =   "软键盘"
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMKEYBOARD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1124
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H001F1F1F&
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   960
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   969
      TabIndex        =   1
      Top             =   360
      Width           =   14535
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   219
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   231
         Left            =   1200
         TabIndex        =   71
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   236
         Left            =   2400
         TabIndex        =   61
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   220
         Left            =   3600
         TabIndex        =   62
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   53
         Left            =   4800
         TabIndex        =   63
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   222
         Left            =   6000
         TabIndex        =   64
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   52
         Left            =   7200
         TabIndex        =   65
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   50
         Left            =   8400
         TabIndex        =   66
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   239
         Left            =   9600
         TabIndex        =   67
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   51
         Left            =   10800
         TabIndex        =   68
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   238
         Left            =   10800
         TabIndex        =   69
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   235
         Left            =   0
         TabIndex        =   72
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   221
         Left            =   1200
         TabIndex        =   73
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   186
         Left            =   6000
         TabIndex        =   74
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   229
         Left            =   2400
         TabIndex        =   75
         Top             =   480
         Width           =   2415
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   192
         Left            =   4800
         TabIndex        =   76
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   54
         Left            =   7200
         TabIndex        =   77
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   49
         Left            =   8400
         TabIndex        =   78
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   191
         Left            =   9600
         TabIndex        =   79
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   56
         Left            =   10800
         TabIndex        =   80
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   55
         Left            =   12000
         TabIndex        =   81
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICM 
         Height          =   495
         Index           =   1
         Left            =   12000
         TabIndex        =   82
         Top             =   0
         Width           =   2415
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   96
         Left            =   0
         TabIndex        =   83
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   97
         Left            =   1200
         TabIndex        =   84
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   99
         Left            =   3600
         TabIndex        =   85
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   100
         Left            =   4800
         TabIndex        =   86
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   101
         Left            =   6000
         TabIndex        =   87
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   102
         Left            =   7200
         TabIndex        =   88
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   103
         Left            =   8400
         TabIndex        =   89
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   104
         Left            =   9600
         TabIndex        =   90
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   105
         Left            =   10800
         TabIndex        =   91
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   110
         Left            =   12000
         TabIndex        =   92
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 小回车 
         Height          =   495
         Left            =   12000
         TabIndex        =   93
         Top             =   960
         Width           =   2415
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   188
         Left            =   0
         TabIndex        =   94
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   190
         Left            =   1200
         TabIndex        =   95
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   57
         Left            =   2400
         TabIndex        =   96
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   48
         Left            =   3600
         TabIndex        =   97
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   106
         Left            =   7200
         TabIndex        =   98
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   111
         Left            =   8400
         TabIndex        =   99
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 数字 
         Height          =   495
         Index           =   98
         Left            =   2400
         TabIndex        =   100
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   187
         Left            =   9600
         TabIndex        =   101
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   189
         Left            =   13200
         TabIndex        =   102
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   237
         Left            =   13200
         TabIndex        =   103
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   107
         Left            =   4800
         TabIndex        =   106
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   109
         Left            =   6000
         TabIndex        =   107
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
   End
   Begin VB.PictureBox PO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H001F1F1F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Index           =   0
      Left            =   0
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1096
      TabIndex        =   0
      Top             =   75
      Width           =   16440
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   27
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   112
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   113
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   114
         Left            =   3600
         TabIndex        =   5
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   115
         Left            =   4800
         TabIndex        =   6
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   116
         Left            =   6000
         TabIndex        =   7
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   117
         Left            =   7200
         TabIndex        =   8
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   118
         Left            =   8400
         TabIndex        =   9
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   119
         Left            =   9600
         TabIndex        =   10
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   120
         Left            =   10800
         TabIndex        =   11
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   121
         Left            =   12000
         TabIndex        =   12
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   122
         Left            =   13200
         TabIndex        =   13
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   123
         Left            =   14400
         TabIndex        =   14
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   9
         Left            =   0
         TabIndex        =   15
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   81
         Left            =   1200
         TabIndex        =   16
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   87
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   69
         Left            =   3600
         TabIndex        =   18
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   82
         Left            =   4800
         TabIndex        =   19
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   84
         Left            =   6000
         TabIndex        =   20
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   89
         Left            =   7200
         TabIndex        =   21
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   85
         Left            =   8400
         TabIndex        =   22
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   73
         Left            =   9600
         TabIndex        =   23
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   79
         Left            =   10800
         TabIndex        =   24
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   80
         Left            =   12000
         TabIndex        =   25
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   44
         Left            =   13200
         TabIndex        =   26
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   19
         Left            =   14400
         TabIndex        =   27
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   65
         Left            =   1200
         TabIndex        =   28
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   83
         Left            =   2400
         TabIndex        =   29
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   68
         Left            =   3600
         TabIndex        =   30
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   70
         Left            =   4800
         TabIndex        =   31
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   71
         Left            =   6000
         TabIndex        =   32
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   72
         Left            =   7200
         TabIndex        =   33
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   74
         Left            =   8400
         TabIndex        =   34
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   75
         Left            =   9600
         TabIndex        =   35
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   76
         Left            =   10800
         TabIndex        =   36
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   8
         Left            =   12000
         TabIndex        =   37
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   90
         Left            =   1200
         TabIndex        =   38
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   88
         Left            =   2400
         TabIndex        =   39
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   67
         Left            =   3600
         TabIndex        =   40
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   86
         Left            =   4800
         TabIndex        =   41
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   66
         Left            =   6000
         TabIndex        =   42
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   78
         Left            =   7200
         TabIndex        =   43
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   77
         Left            =   8400
         TabIndex        =   44
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   33
         Left            =   9600
         TabIndex        =   45
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   36
         Left            =   14400
         TabIndex        =   46
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   45
         Left            =   13200
         TabIndex        =   47
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 转换 
         Height          =   495
         Index           =   164
         Left            =   1200
         TabIndex        =   48
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   32
         Left            =   2400
         TabIndex        =   49
         Top             =   1920
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 徽标 
         Height          =   495
         Left            =   6000
         TabIndex        =   50
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   13
         Left            =   7200
         TabIndex        =   51
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   46
         Left            =   13200
         TabIndex        =   52
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   975
         Index           =   35
         Left            =   14400
         TabIndex        =   53
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   34
         Left            =   10800
         TabIndex        =   54
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 换档 
         Height          =   495
         Index           =   160
         Left            =   0
         TabIndex        =   55
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY 控制 
         Height          =   495
         Index           =   162
         Left            =   0
         TabIndex        =   56
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   38
         Left            =   12000
         TabIndex        =   57
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   37
         Left            =   10800
         TabIndex        =   58
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   39
         Left            =   13200
         TabIndex        =   59
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY LK 
         Height          =   495
         Index           =   40
         Left            =   12000
         TabIndex        =   60
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICM 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   104
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICM 
         Height          =   495
         Index           =   2
         Left            =   8400
         TabIndex        =   105
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
      End
   End
End
Attribute VB_Name = "FRMKEYBOARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Dim Esc, b1, minus, plus, straight, Backspace, TTAB, Lett(65 To 90), num(48 To 57), Caps, Shft, Ctrl, Alt, LBS, RB, Q, a, TD, SBL, SBR, SP, Ent
Dim FF(112 To 123)
Dim i, O, io, S_Change As Boolean, C_CHANGE As Boolean, A_CHANGE As Boolean
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal Enable As Long, ByVal CurrentThread As Long, Enabled As Long) As Long
Private Const SE_DEBUG_PRIVILEGE = &H13
Private Declare Function SetSuspendState Lib "Powrprof" (ByVal Hibernate As Boolean, ByVal ForceCritical As Boolean, ByVal DisableWakeEvent As Boolean) As Boolean
Private Declare Function NtShutdownSystem Lib "ntdll" (ByVal ShutdownAction As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Const shutdown& = 0
Private Const RESTART& = 1
Private Const POWEROFF& = 2
Const EWX_FORCE As Long = 4
Const EWX_LOGOFF As Long = 0
Const EWX_REBOOT As Long = 2
Const EWX_SHUTDOWN As Long = 1

Rem 禁止本窗体拥有输入焦点的常数
Private Const HWND_NOTOPMOST = -2
Private Const WS_DISABLED = &H8000000
Private Const GWL_STYLE = (-16)

Rem 窗口置顶的常数
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Rem 移动没有标题栏窗体的常数
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Rem 模拟按钮常数
Private Const KEYEVENTF_KEYUP = &H2

Private Sub Form_Initialize()
If Screen.Width / 15 < 1024 Then Call SHOWWRONG("对不起,由于您的分辨率过小,无法启动软键盘", 2): Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = COLOR_NOR
Dim re As RECT
Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is ICEE_KEY Then PBOX.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
If TypeOf PBOX Is PictureBox Then PBOX.BackColor = COLOR_NOR
Next
GetWindowRect FindWindow("Shell_TrayWnd", vbNullString), re '获取任务栏信息
    Me.Move 0, Screen.Height - Me.Height - GetTaskbarHeight, Screen.Width
    Call 小写字母
    Call 下档符号
    Call 数字小键盘
    Call SeekMe(Me)
    RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    MakeTransparent Me.hwnd, 255
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_DISABLED
    PO(1).Visible = False
    LK(27).SETTXT "ESC"
    LK(112).SETTXT "F1"
    LK(113).SETTXT "F2"
    LK(114).SETTXT "F3"
    LK(115).SETTXT "F4"
    LK(116).SETTXT "F5"
    LK(117).SETTXT "F6"
    LK(118).SETTXT "F7"
    LK(119).SETTXT "F8"
    LK(120).SETTXT "F9"
    LK(121).SETTXT "F10"
    LK(122).SETTXT "F11"
    LK(123).SETTXT "F12"
    LK(9).SETTXT "Tab"
    LK(13).SETTXT "Enter"
    LK(38).SETTXT "↑"
    LK(39).SETTXT "→"
    LK(37).SETTXT "←"
    LK(40).SETTXT "↓"
    LK(8).SETTXT "退格"
    LK(33).SETTXT "PageUp"
    LK(34).SETTXT "PageDown"
    LK(35).SETTXT "End"
    LK(36).SETTXT "Home"
    LK(19).SETTXT "Pause"
    LK(44).SETTXT "ScreenPrint"
    LK(32).SETTXT ""
    LK(107).SETTXT "+"
    LK(109).SETTXT "-"
    LK(45).SETTXT "Insert"
    LK(46).SETTXT "Delete"
    换档(160).SETTXT "Shift"
    控制(162).SETTXT "Ctrl"
    转换(164).SETTXT "Alt"
    徽标.SETTXT "Win"
    小回车.SETTXT "Enter"
    
    ICM(0).SETTXT "数字/符号"
    ICM(1).SETTXT "返回字母"
    ICM(2).SETTXT "关闭软键盘"
End Sub
Private Sub Form_Resize()
PO(0).Left = (Me.ScaleWidth - PO(0).Width) / 2
PO(1).Left = (Me.ScaleWidth - PO(1).Width) / 2

End Sub


Private Sub ICM_Click(Index As Integer)
Select Case Index
Case 0
PO(0).Visible = False
PO(1).Visible = True
Case 1
PO(1).Visible = False
PO(0).Visible = True
Case 2
Unload Me
End Select
End Sub

Private Sub LK_CLICK(Index As Integer)
    keybd_event Index, 0, 0, 0
    keybd_event Index, 0, KEYEVENTF_KEYUP, 0
    If S_Change = False Then
    
        Call 上档符号键(Index)
        keybd_event 160, 0, KEYEVENTF_KEYUP, 0
        数字小键盘
        下档符号
        小写字母
    Else
        Call 下档符号键(Index)
 End If
End Sub

Private Sub 换档_CLICK(Index As Integer)
If S_Change = False Then
S_Change = True
        keybd_event Index, 0, 0, 0
        阿拉伯数字序号小键盘
        上档符号
        大写字母
    Else
S_Change = False
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0
        数字小键盘
        下档符号
        小写字母
    End If
End Sub
Rem 四组成对的特殊键
Private Sub 控制_Click(Index As Integer)
    If C_CHANGE = True Then
        C_CHANGE = False
        keybd_event Index, 0, 0, 0
    Else
        C_CHANGE = True
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub
Private Sub 数字_CLICK(Index As Integer)
    If S_Change = True Then
        Call 阿拉伯数字序号(Index)
        S_Change = False
        keybd_event 160, 0, KEYEVENTF_KEYUP, 0
        数字小键盘
        下档符号
        小写字母
    Else
        keybd_event Index, 0, 0, 0
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0

    End If
End Sub

Private Sub 转换_Click(Index As Integer)
    If A_CHANGE = False Then
        A_CHANGE = True
        keybd_event Index, 0, 0, 0
    Else
       A_CHANGE = False
       keybd_event Index, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub
Rem 徽标
Private Sub 徽标_Click()
    Rem If 徽标.BackColor = &H00633F0E& Then
    Rem    徽标.BackColor = &H001B9F77&
         keybd_event 91, 0, 0, 0
    Rem Else
        keybd_event 91, 0, KEYEVENTF_KEYUP, 0
    Rem End If
End Sub
Rem 小键盘回车
Private Sub 小回车_Click()
    keybd_event 13, 0, 0, 0
    keybd_event 13, 0, KEYEVENTF_KEYUP, 0
End Sub
Rem 自定义子函数，使得小键盘变成阿拉伯数字序号
Sub 阿拉伯数字序号(ByVal NumIndex As Integer)
        Select Case NumIndex
            Case 96
                SendKeys "⑩"
            Case 97
                SendKeys "①"
            Case 98
                SendKeys "②"
            Case 99
                SendKeys "③"
            Case 100
                SendKeys "④"
            Case 101
                SendKeys "⑤"
            Case 102
                SendKeys "⑥"
            Case 103
                SendKeys "⑦"
            Case 104
                SendKeys "⑧"
            Case 105
                SendKeys "⑨"
            Case 110
                SendKeys "°"
        End Select
End Sub

Rem 自定义子函数，使子母键突出显示大写
Sub 大写字母()
    LK(65).SETTXT "Ａ"
    LK(66).SETTXT "Ｂ"
    LK(67).SETTXT "Ｃ"
    LK(68).SETTXT "Ｄ"
    LK(69).SETTXT "Ｅ"
    LK(70).SETTXT "Ｆ"
    LK(71).SETTXT "Ｇ"
    LK(72).SETTXT "Ｈ"
    LK(73).SETTXT "Ｉ"
    LK(74).SETTXT "Ｊ"
    LK(75).SETTXT "Ｋ"
    LK(76).SETTXT "Ｌ"
    LK(77).SETTXT "Ｍ"
    LK(78).SETTXT "Ｎ"
    LK(79).SETTXT "Ｏ"
    LK(80).SETTXT "Ｐ"
    LK(81).SETTXT "Ｑ"
    LK(82).SETTXT "Ｒ"
    LK(83).SETTXT "Ｓ"
    LK(84).SETTXT "Ｔ"
    LK(85).SETTXT "Ｕ"
    LK(86).SETTXT "Ｖ"
    LK(87).SETTXT "Ｗ"
    LK(88).SETTXT "Ｘ"
    LK(89).SETTXT "Ｙ"
    LK(90).SETTXT "Ｚ"
End Sub
Rem 自定义子函数，使子母键突出显示小写
Sub 小写字母()
    LK(65).SETTXT "ａ"
    LK(66).SETTXT "ｂ"
    LK(67).SETTXT "ｃ"
    LK(68).SETTXT "ｄ"
    LK(69).SETTXT "ｅ"
    LK(70).SETTXT "ｆ"
    LK(71).SETTXT "ｇ"
    LK(72).SETTXT "ｈ"
    LK(73).SETTXT "ｉ"
    LK(74).SETTXT "ｊ"
    LK(75).SETTXT "ｋ"
    LK(76).SETTXT "ｌ"
    LK(77).SETTXT "ｍ"
    LK(78).SETTXT "ｎ"
    LK(79).SETTXT "ｏ"
    LK(80).SETTXT "ｐ"
    LK(81).SETTXT "ｑ"
    LK(82).SETTXT "ｒ"
    LK(83).SETTXT "ｓ"
    LK(84).SETTXT "ｔ"
    LK(85).SETTXT "ｕ"
    LK(86).SETTXT "ｖ"
    LK(87).SETTXT "ｗ"
    LK(88).SETTXT "ｘ"
    LK(89).SETTXT "ｙ"
    LK(90).SETTXT "ｚ"
End Sub
Rem 自定义子函数，使运算符号突出显示上档
Sub 上档符号()
    LK(48).SETTXT "㈩"
    LK(49).SETTXT "㈠"
    LK(50).SETTXT "㈡"
    LK(51).SETTXT "㈢"
    LK(52).SETTXT "㈣"
    LK(53).SETTXT "㈤"
    LK(54).SETTXT "㈥"
    LK(55).SETTXT "㈦"
    LK(56).SETTXT "㈧"
    LK(57).SETTXT "㈨"
    LK(106).SETTXT "×"
    LK(111).SETTXT "÷"
    LK(186).SETTXT ":"
    LK(187).SETTXT "="
    LK(188).SETTXT "《"
    LK(189).SETTXT "——"
    LK(190).SETTXT "》"
    LK(191).SETTXT "."
    LK(192).SETTXT "㏄"
    LK(219).SETTXT "｛"
    LK(220).SETTXT "."
    LK(221).SETTXT "】"
    LK(222).SETTXT "＇"
    LK(229).SETTXT "℃"
    LK(231).SETTXT "｝"
    LK(235).SETTXT "【"
    LK(236).SETTXT """"
    LK(237).SETTXT "；"
    LK(238).SETTXT ","
    LK(239).SETTXT "."
End Sub
Rem 自定义子函数，使运算符号突出显示下档
Sub 下档符号()
    LK(48).SETTXT "）"
    LK(49).SETTXT "!"
    LK(50).SETTXT "＠"
    LK(51).SETTXT "＃"
    LK(52).SETTXT "＄"
    LK(53).SETTXT "％"
    LK(54).SETTXT "＾"
    LK(55).SETTXT "＆"
    LK(56).SETTXT "｜"
    LK(57).SETTXT "（"
    LK(106).SETTXT "＊"
    LK(111).SETTXT "／"
    LK(186).SETTXT ":"
    LK(187).SETTXT "＝"
    LK(188).SETTXT "＜"
    LK(189).SETTXT "__"
    LK(190).SETTXT "＞"
    LK(191).SETTXT "."
    LK(192).SETTXT "~"
    LK(219).SETTXT "｛"
    LK(220).SETTXT "＼"
    LK(231).SETTXT "｝"
    LK(222).SETTXT "＇"
    LK(229).SETTXT "｀"
    LK(221).SETTXT "］"
    LK(235).SETTXT "［"
    LK(236).SETTXT """"
    LK(237).SETTXT "；"
    LK(238).SETTXT ","
    LK(239).SETTXT "．"
    End Sub
Rem 自定义子函数，使LX突出显示数字
Sub 数字小键盘()
    数字(96).SETTXT "０"
    数字(97).SETTXT "１"
    数字(98).SETTXT "２"
    数字(99).SETTXT "３"
    数字(100).SETTXT "４"
    数字(101).SETTXT "５"
    数字(102).SETTXT "６"
    数字(103).SETTXT "７"
    数字(104).SETTXT "８"
    数字(105).SETTXT "９"
    数字(110).SETTXT "．"
End Sub
Rem 自定义子函数，使LX突出显示阿拉伯数字序号
Sub 阿拉伯数字序号小键盘()
    数字(96).SETTXT "⑩"
    数字(97).SETTXT "①"
    数字(98).SETTXT "②"
    数字(99).SETTXT "③"
    数字(100).SETTXT "④"
    数字(101).SETTXT "⑤"
    数字(102).SETTXT "⑥"
    数字(103).SETTXT "⑦"
    数字(104).SETTXT "⑧"
    数字(105).SETTXT "⑨"
    数字(110).SETTXT "°"
End Sub

Rem 自定义子函数，使符号键盘变成上档符号
Sub 上档符号键(ByVal NumIndex As Integer)
        Select Case NumIndex
            Case 49
                SendKeys "㈠"
            Case 50
                SendKeys "㈡"
            Case 51
                SendKeys "㈢"
            Case 52
                SendKeys "㈣"
            Case 53
                SendKeys "㈤"
            Case 54
                SendKeys "㈥"
            Case 55
                SendKeys "㈦"
            Case 56
                SendKeys "㈧"
            Case 57
                SendKeys "㈨"
            Case 48
                SendKeys "㈩"
            Case 106
                SendKeys "×"
            Case 111
                SendKeys "÷"
            Case 186
                SendKeys ":"
            Case 187
                SendKeys "＝"
            Case 188
                SendKeys "《"
            Case 189
                SendKeys "——"
            Case 190
                SendKeys "》"
            Case 191
                SendKeys "."
            Case 192
                SendKeys "㏄"
            Case 219
                SendKeys "｛"
            Case 220
                SendKeys "."
            Case 221
                SendKeys "｝"
            Case 222
                SendKeys "＇"
            Case 229
                SendKeys "℃"
            Case 231
                SendKeys "】"
            Case 235
                SendKeys "【"
            Case 236
                SendKeys """"
            Case 237
                SendKeys "；"
            Case 238
                SendKeys ","
            Case 239
                SendKeys "."
       End Select
End Sub
Rem 自定义子函数，使符号键盘变成下档符号
Sub 下档符号键(ByVal NumIndex As Integer)
        Select Case NumIndex
            Case 49
                SendKeys "!"
            Case 50
                SendKeys "@"
            Case 51
                SendKeys "#"
            Case 52
                SendKeys "$"
            Case 53
                SendKeys "{%}"
            Case 54
                SendKeys "{^}"
            Case 55
                SendKeys "&"
            Case 56
                SendKeys "|"
            Case 57
                SendKeys "{(}"
            Case 48
                SendKeys "{)}"
            Case 106
                SendKeys "*"
            Case 111
                SendKeys "/"
            Case 186
                SendKeys ":"
            Case 187
                SendKeys "="
            Case 188
                SendKeys "<"
            Case 189
                SendKeys "_"
            Case 190
                SendKeys ">"
            Case 191
                SendKeys "."
            Case 192
                SendKeys "{~}"
            Case 219
                SendKeys "{{}"
            Case 220
                SendKeys "\"
            Case 221
                SendKeys "]"
            Case 222
                SendKeys "'"
            Case 229
                SendKeys "`"
            Case 231
                SendKeys "{}}"
            Case 235
                SendKeys "["
            Case 236
                SendKeys """"
            Case 237
                SendKeys ";"
            Case 238
                SendKeys ","
            Case 239
                SendKeys "."
       End Select
End Sub
