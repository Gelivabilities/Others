VERSION 5.00
Begin VB.Form FRMKEYBOARD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H001F1F1F&
   BorderStyle     =   0  'None
   Caption         =   "�����"
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
   StartUpPosition =   3  '����ȱʡ
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
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   96
         Left            =   0
         TabIndex        =   83
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   97
         Left            =   1200
         TabIndex        =   84
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   99
         Left            =   3600
         TabIndex        =   85
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   100
         Left            =   4800
         TabIndex        =   86
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   101
         Left            =   6000
         TabIndex        =   87
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   102
         Left            =   7200
         TabIndex        =   88
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   103
         Left            =   8400
         TabIndex        =   89
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   104
         Left            =   9600
         TabIndex        =   90
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   105
         Left            =   10800
         TabIndex        =   91
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   110
         Left            =   12000
         TabIndex        =   92
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY С�س� 
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
      Begin ICEE.ICEE_KEY ���� 
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
      Begin ICEE.ICEE_KEY ת�� 
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
      Begin ICEE.ICEE_KEY �ձ� 
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
      Begin ICEE.ICEE_KEY ���� 
         Height          =   495
         Index           =   160
         Left            =   0
         TabIndex        =   55
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ���� 
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

Rem ��ֹ������ӵ�����뽹��ĳ���
Private Const HWND_NOTOPMOST = -2
Private Const WS_DISABLED = &H8000000
Private Const GWL_STYLE = (-16)

Rem �����ö��ĳ���
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Rem �ƶ�û�б���������ĳ���
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Rem ģ�ⰴť����
Private Const KEYEVENTF_KEYUP = &H2

Private Sub Form_Initialize()
If Screen.Width / 15 < 1024 Then Call SHOWWRONG("�Բ���,�������ķֱ��ʹ�С,�޷����������", 2): Unload Me
End Sub

Private Sub Form_Load()
Me.BackColor = COLOR_NOR
Dim re As RECT
Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is ICEE_KEY Then PBOX.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
If TypeOf PBOX Is PictureBox Then PBOX.BackColor = COLOR_NOR
Next
GetWindowRect FindWindow("Shell_TrayWnd", vbNullString), re '��ȡ��������Ϣ
    Me.Move 0, Screen.Height - Me.Height - GetTaskbarHeight, Screen.Width
    Call Сд��ĸ
    Call �µ�����
    Call ����С����
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
    LK(38).SETTXT "��"
    LK(39).SETTXT "��"
    LK(37).SETTXT "��"
    LK(40).SETTXT "��"
    LK(8).SETTXT "�˸�"
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
    ����(160).SETTXT "Shift"
    ����(162).SETTXT "Ctrl"
    ת��(164).SETTXT "Alt"
    �ձ�.SETTXT "Win"
    С�س�.SETTXT "Enter"
    
    ICM(0).SETTXT "����/����"
    ICM(1).SETTXT "������ĸ"
    ICM(2).SETTXT "�ر������"
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
    
        Call �ϵ����ż�(Index)
        keybd_event 160, 0, KEYEVENTF_KEYUP, 0
        ����С����
        �µ�����
        Сд��ĸ
    Else
        Call �µ����ż�(Index)
 End If
End Sub

Private Sub ����_CLICK(Index As Integer)
If S_Change = False Then
S_Change = True
        keybd_event Index, 0, 0, 0
        �������������С����
        �ϵ�����
        ��д��ĸ
    Else
S_Change = False
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0
        ����С����
        �µ�����
        Сд��ĸ
    End If
End Sub
Rem ����ɶԵ������
Private Sub ����_Click(Index As Integer)
    If C_CHANGE = True Then
        C_CHANGE = False
        keybd_event Index, 0, 0, 0
    Else
        C_CHANGE = True
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub
Private Sub ����_CLICK(Index As Integer)
    If S_Change = True Then
        Call �������������(Index)
        S_Change = False
        keybd_event 160, 0, KEYEVENTF_KEYUP, 0
        ����С����
        �µ�����
        Сд��ĸ
    Else
        keybd_event Index, 0, 0, 0
        keybd_event Index, 0, KEYEVENTF_KEYUP, 0

    End If
End Sub

Private Sub ת��_Click(Index As Integer)
    If A_CHANGE = False Then
        A_CHANGE = True
        keybd_event Index, 0, 0, 0
    Else
       A_CHANGE = False
       keybd_event Index, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub
Rem �ձ�
Private Sub �ձ�_Click()
    Rem If �ձ�.BackColor = &H00633F0E& Then
    Rem    �ձ�.BackColor = &H001B9F77&
         keybd_event 91, 0, 0, 0
    Rem Else
        keybd_event 91, 0, KEYEVENTF_KEYUP, 0
    Rem End If
End Sub
Rem С���̻س�
Private Sub С�س�_Click()
    keybd_event 13, 0, 0, 0
    keybd_event 13, 0, KEYEVENTF_KEYUP, 0
End Sub
Rem �Զ����Ӻ�����ʹ��С���̱�ɰ������������
Sub �������������(ByVal NumIndex As Integer)
        Select Case NumIndex
            Case 96
                SendKeys "��"
            Case 97
                SendKeys "��"
            Case 98
                SendKeys "��"
            Case 99
                SendKeys "��"
            Case 100
                SendKeys "��"
            Case 101
                SendKeys "��"
            Case 102
                SendKeys "��"
            Case 103
                SendKeys "��"
            Case 104
                SendKeys "��"
            Case 105
                SendKeys "��"
            Case 110
                SendKeys "��"
        End Select
End Sub

Rem �Զ����Ӻ�����ʹ��ĸ��ͻ����ʾ��д
Sub ��д��ĸ()
    LK(65).SETTXT "��"
    LK(66).SETTXT "��"
    LK(67).SETTXT "��"
    LK(68).SETTXT "��"
    LK(69).SETTXT "��"
    LK(70).SETTXT "��"
    LK(71).SETTXT "��"
    LK(72).SETTXT "��"
    LK(73).SETTXT "��"
    LK(74).SETTXT "��"
    LK(75).SETTXT "��"
    LK(76).SETTXT "��"
    LK(77).SETTXT "��"
    LK(78).SETTXT "��"
    LK(79).SETTXT "��"
    LK(80).SETTXT "��"
    LK(81).SETTXT "��"
    LK(82).SETTXT "��"
    LK(83).SETTXT "��"
    LK(84).SETTXT "��"
    LK(85).SETTXT "��"
    LK(86).SETTXT "��"
    LK(87).SETTXT "��"
    LK(88).SETTXT "��"
    LK(89).SETTXT "��"
    LK(90).SETTXT "��"
End Sub
Rem �Զ����Ӻ�����ʹ��ĸ��ͻ����ʾСд
Sub Сд��ĸ()
    LK(65).SETTXT "��"
    LK(66).SETTXT "��"
    LK(67).SETTXT "��"
    LK(68).SETTXT "��"
    LK(69).SETTXT "��"
    LK(70).SETTXT "��"
    LK(71).SETTXT "��"
    LK(72).SETTXT "��"
    LK(73).SETTXT "��"
    LK(74).SETTXT "��"
    LK(75).SETTXT "��"
    LK(76).SETTXT "��"
    LK(77).SETTXT "��"
    LK(78).SETTXT "��"
    LK(79).SETTXT "��"
    LK(80).SETTXT "��"
    LK(81).SETTXT "��"
    LK(82).SETTXT "��"
    LK(83).SETTXT "��"
    LK(84).SETTXT "��"
    LK(85).SETTXT "��"
    LK(86).SETTXT "��"
    LK(87).SETTXT "��"
    LK(88).SETTXT "��"
    LK(89).SETTXT "��"
    LK(90).SETTXT "��"
End Sub
Rem �Զ����Ӻ�����ʹ�������ͻ����ʾ�ϵ�
Sub �ϵ�����()
    LK(48).SETTXT "��"
    LK(49).SETTXT "��"
    LK(50).SETTXT "��"
    LK(51).SETTXT "��"
    LK(52).SETTXT "��"
    LK(53).SETTXT "��"
    LK(54).SETTXT "��"
    LK(55).SETTXT "��"
    LK(56).SETTXT "��"
    LK(57).SETTXT "��"
    LK(106).SETTXT "��"
    LK(111).SETTXT "��"
    LK(186).SETTXT ":"
    LK(187).SETTXT "="
    LK(188).SETTXT "��"
    LK(189).SETTXT "����"
    LK(190).SETTXT "��"
    LK(191).SETTXT "."
    LK(192).SETTXT "�P"
    LK(219).SETTXT "��"
    LK(220).SETTXT "."
    LK(221).SETTXT "��"
    LK(222).SETTXT "��"
    LK(229).SETTXT "��"
    LK(231).SETTXT "��"
    LK(235).SETTXT "��"
    LK(236).SETTXT """"
    LK(237).SETTXT "��"
    LK(238).SETTXT ","
    LK(239).SETTXT "."
End Sub
Rem �Զ����Ӻ�����ʹ�������ͻ����ʾ�µ�
Sub �µ�����()
    LK(48).SETTXT "��"
    LK(49).SETTXT "!"
    LK(50).SETTXT "��"
    LK(51).SETTXT "��"
    LK(52).SETTXT "��"
    LK(53).SETTXT "��"
    LK(54).SETTXT "��"
    LK(55).SETTXT "��"
    LK(56).SETTXT "��"
    LK(57).SETTXT "��"
    LK(106).SETTXT "��"
    LK(111).SETTXT "��"
    LK(186).SETTXT ":"
    LK(187).SETTXT "��"
    LK(188).SETTXT "��"
    LK(189).SETTXT "__"
    LK(190).SETTXT "��"
    LK(191).SETTXT "."
    LK(192).SETTXT "~"
    LK(219).SETTXT "��"
    LK(220).SETTXT "��"
    LK(231).SETTXT "��"
    LK(222).SETTXT "��"
    LK(229).SETTXT "��"
    LK(221).SETTXT "��"
    LK(235).SETTXT "��"
    LK(236).SETTXT """"
    LK(237).SETTXT "��"
    LK(238).SETTXT ","
    LK(239).SETTXT "��"
    End Sub
Rem �Զ����Ӻ�����ʹLXͻ����ʾ����
Sub ����С����()
    ����(96).SETTXT "��"
    ����(97).SETTXT "��"
    ����(98).SETTXT "��"
    ����(99).SETTXT "��"
    ����(100).SETTXT "��"
    ����(101).SETTXT "��"
    ����(102).SETTXT "��"
    ����(103).SETTXT "��"
    ����(104).SETTXT "��"
    ����(105).SETTXT "��"
    ����(110).SETTXT "��"
End Sub
Rem �Զ����Ӻ�����ʹLXͻ����ʾ�������������
Sub �������������С����()
    ����(96).SETTXT "��"
    ����(97).SETTXT "��"
    ����(98).SETTXT "��"
    ����(99).SETTXT "��"
    ����(100).SETTXT "��"
    ����(101).SETTXT "��"
    ����(102).SETTXT "��"
    ����(103).SETTXT "��"
    ����(104).SETTXT "��"
    ����(105).SETTXT "��"
    ����(110).SETTXT "��"
End Sub

Rem �Զ����Ӻ�����ʹ���ż��̱���ϵ�����
Sub �ϵ����ż�(ByVal NumIndex As Integer)
        Select Case NumIndex
            Case 49
                SendKeys "��"
            Case 50
                SendKeys "��"
            Case 51
                SendKeys "��"
            Case 52
                SendKeys "��"
            Case 53
                SendKeys "��"
            Case 54
                SendKeys "��"
            Case 55
                SendKeys "��"
            Case 56
                SendKeys "��"
            Case 57
                SendKeys "��"
            Case 48
                SendKeys "��"
            Case 106
                SendKeys "��"
            Case 111
                SendKeys "��"
            Case 186
                SendKeys ":"
            Case 187
                SendKeys "��"
            Case 188
                SendKeys "��"
            Case 189
                SendKeys "����"
            Case 190
                SendKeys "��"
            Case 191
                SendKeys "."
            Case 192
                SendKeys "�P"
            Case 219
                SendKeys "��"
            Case 220
                SendKeys "."
            Case 221
                SendKeys "��"
            Case 222
                SendKeys "��"
            Case 229
                SendKeys "��"
            Case 231
                SendKeys "��"
            Case 235
                SendKeys "��"
            Case 236
                SendKeys """"
            Case 237
                SendKeys "��"
            Case 238
                SendKeys ","
            Case 239
                SendKeys "."
       End Select
End Sub
Rem �Զ����Ӻ�����ʹ���ż��̱���µ�����
Sub �µ����ż�(ByVal NumIndex As Integer)
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
