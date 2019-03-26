VERSION 5.00
Begin VB.UserControl ICEE_Calender 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0084536F&
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   ControlContainer=   -1  'True
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   403
   Begin ICEE.IList lstYear 
      Height          =   2520
      Left            =   1800
      TabIndex        =   92
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   4868
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
   Begin ICEE.IList lstMonth 
      Height          =   2520
      Left            =   3240
      TabIndex        =   91
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   4868
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
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0084536F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   409
      TabIndex        =   93
      Top             =   600
      Width           =   6135
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "周六"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   255
         Index           =   6
         Left            =   5070
         TabIndex        =   100
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "周五"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   4275
         TabIndex        =   99
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "周四"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   3465
         TabIndex        =   98
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "周三"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2670
         TabIndex        =   97
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "周二"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1875
         TabIndex        =   96
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "周一"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1065
         TabIndex        =   95
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "周日"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   94
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0084536F&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   6060
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D2F0F5&
         Height          =   330
         Left            =   2925
         TabIndex        =   101
         Top             =   180
         Width           =   105
      End
      Begin VB.Label LBXZ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "狮子座"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   600
         MouseIcon       =   "ICEE_Calender.ctx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   90
         Top             =   240
         Width           =   540
      End
      Begin VB.Label LBSX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "鸡年"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   360
      End
      Begin VB.Image imgUp 
         Height          =   240
         Left            =   4162
         MouseIcon       =   "ICEE_Calender.ctx":0152
         MousePointer    =   99  'Custom
         Picture         =   "ICEE_Calender.ctx":02A4
         Top             =   240
         Width           =   240
      End
      Begin VB.Image ImgDown 
         Height          =   240
         Left            =   1642
         MouseIcon       =   "ICEE_Calender.ctx":062E
         MousePointer    =   99  'Custom
         Picture         =   "ICEE_Calender.ctx":0780
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblBackToday 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "今天"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5160
         MouseIcon       =   "ICEE_Calender.ctx":0B0A
         MousePointer    =   99  'Custom
         TabIndex        =   88
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblM 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D2F0F5&
         Height          =   330
         Left            =   3180
         MouseIcon       =   "ICEE_Calender.ctx":0C5C
         MousePointer    =   99  'Custom
         TabIndex        =   87
         Top             =   180
         Width           =   150
      End
      Begin VB.Label lblY 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D2F0F5&
         Height          =   330
         Left            =   2520
         MouseIcon       =   "ICEE_Calender.ctx":0DAE
         MousePointer    =   99  'Custom
         TabIndex        =   86
         Top             =   180
         Width           =   150
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0084536F&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6060
      Begin VB.Timer TMOUT 
         Interval        =   500
         Left            =   5400
         Top             =   480
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   35
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":0F00
         MousePointer    =   99  'Custom
         TabIndex        =   84
         Top             =   3155
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   35
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":1052
         MousePointer    =   99  'Custom
         TabIndex        =   83
         Top             =   3395
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   36
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":11A4
         MousePointer    =   99  'Custom
         TabIndex        =   82
         Top             =   3155
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   36
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":12F6
         MousePointer    =   99  'Custom
         TabIndex        =   81
         Top             =   3395
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   37
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":1448
         MousePointer    =   99  'Custom
         TabIndex        =   80
         Top             =   3155
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   37
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":159A
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   3395
         Width           =   360
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   38
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":16EC
         MousePointer    =   99  'Custom
         TabIndex        =   78
         Top             =   3395
         Width           =   180
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   38
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":183E
         MousePointer    =   99  'Custom
         TabIndex        =   77
         Top             =   3155
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   39
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":1990
         MousePointer    =   99  'Custom
         TabIndex        =   76
         Top             =   3395
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   39
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":1AE2
         MousePointer    =   99  'Custom
         TabIndex        =   75
         Top             =   3155
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   40
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":1C34
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   41
         Left            =   960
         MouseIcon       =   "ICEE_Calender.ctx":1D86
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   3395
         Width           =   180
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   40
         Left            =   840
         MouseIcon       =   "ICEE_Calender.ctx":1ED8
         MousePointer    =   99  'Custom
         TabIndex        =   72
         Top             =   3155
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   41
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":202A
         MousePointer    =   99  'Custom
         TabIndex        =   71
         Top             =   3155
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   28
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":217C
         MousePointer    =   99  'Custom
         TabIndex        =   70
         Top             =   2535
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   29
         Left            =   840
         MouseIcon       =   "ICEE_Calender.ctx":22CE
         MousePointer    =   99  'Custom
         TabIndex        =   69
         Top             =   2535
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   28
         Left            =   960
         MouseIcon       =   "ICEE_Calender.ctx":2420
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   2775
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   29
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":2572
         MousePointer    =   99  'Custom
         TabIndex        =   67
         Top             =   2775
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   30
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":26C4
         MousePointer    =   99  'Custom
         TabIndex        =   66
         Top             =   2535
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   30
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":2816
         MousePointer    =   99  'Custom
         TabIndex        =   65
         Top             =   2775
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   31
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":2968
         MousePointer    =   99  'Custom
         TabIndex        =   64
         Top             =   2535
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   31
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":2ABA
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   2775
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   32
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":2C0C
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   2775
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   32
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":2D5E
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   2535
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   33
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":2EB0
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   2775
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   33
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":3002
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   2535
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   34
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":3154
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   2775
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   34
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":32A6
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   2535
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   21
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":33F8
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   1925
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   22
         Left            =   840
         MouseIcon       =   "ICEE_Calender.ctx":354A
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   1925
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   21
         Left            =   960
         MouseIcon       =   "ICEE_Calender.ctx":369C
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   2165
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   22
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":37EE
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   2165
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   23
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":3940
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   1925
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   23
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":3A92
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   2165
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   24
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":3BE4
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   1925
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   24
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":3D36
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   2165
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   25
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":3E88
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   2165
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   25
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":3FDA
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   1925
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   26
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":412C
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   2165
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   26
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":427E
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   1925
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   27
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":43D0
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   2165
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   27
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":4522
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   1925
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   14
         Left            =   240
         MouseIcon       =   "ICEE_Calender.ctx":4674
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   15
         Left            =   720
         MouseIcon       =   "ICEE_Calender.ctx":47C6
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   14
         Left            =   840
         MouseIcon       =   "ICEE_Calender.ctx":4918
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   1545
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   240
         MouseIcon       =   "ICEE_Calender.ctx":4A6A
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   1545
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   16
         Left            =   1320
         MouseIcon       =   "ICEE_Calender.ctx":4BBC
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   1320
         MouseIcon       =   "ICEE_Calender.ctx":4D0E
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   1545
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   17
         Left            =   2040
         MouseIcon       =   "ICEE_Calender.ctx":4E60
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   2040
         MouseIcon       =   "ICEE_Calender.ctx":4FB2
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   1545
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   2880
         MouseIcon       =   "ICEE_Calender.ctx":5104
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   1545
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   18
         Left            =   2880
         MouseIcon       =   "ICEE_Calender.ctx":5256
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   3600
         MouseIcon       =   "ICEE_Calender.ctx":53A8
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   1545
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   19
         Left            =   3600
         MouseIcon       =   "ICEE_Calender.ctx":54FA
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   20
         Left            =   4440
         MouseIcon       =   "ICEE_Calender.ctx":564C
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1545
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   20
         Left            =   4440
         MouseIcon       =   "ICEE_Calender.ctx":579E
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   1305
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   7
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":58F0
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   685
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   7
         Left            =   4560
         MouseIcon       =   "ICEE_Calender.ctx":5A42
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   925
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   8
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":5B94
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   685
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   3720
         MouseIcon       =   "ICEE_Calender.ctx":5CE6
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   925
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   9
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":5E38
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   685
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   3000
         MouseIcon       =   "ICEE_Calender.ctx":5F8A
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   925
         Width           =   360
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   10
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":60DC
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   925
         Width           =   180
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   10
         Left            =   2160
         MouseIcon       =   "ICEE_Calender.ctx":622E
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   685
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":6380
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   925
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   11
         Left            =   1440
         MouseIcon       =   "ICEE_Calender.ctx":64D2
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   685
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":6624
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   925
         Width           =   360
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   13
         Left            =   960
         MouseIcon       =   "ICEE_Calender.ctx":6776
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   925
         Width           =   180
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   12
         Left            =   840
         MouseIcon       =   "ICEE_Calender.ctx":68C8
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   685
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   13
         Left            =   360
         MouseIcon       =   "ICEE_Calender.ctx":6A1A
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   685
         Width           =   270
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   1005
         MouseIcon       =   "ICEE_Calender.ctx":6B6C
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   320
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一一"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   0
         Left            =   105
         MouseIcon       =   "ICEE_Calender.ctx":6CBE
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   1755
         MouseIcon       =   "ICEE_Calender.ctx":6E10
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   320
         Width           =   360
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卅"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   2685
         MouseIcon       =   "ICEE_Calender.ctx":6F62
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   320
         Width           =   180
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   3435
         MouseIcon       =   "ICEE_Calender.ctx":70B4
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   320
         Width           =   360
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   4035
         MouseIcon       =   "ICEE_Calender.ctx":7206
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   320
         Width           =   360
      End
      Begin VB.Label lblNongLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "廿一"
         ForeColor       =   &H005273E8&
         Height          =   180
         Index           =   6
         Left            =   4995
         MouseIcon       =   "ICEE_Calender.ctx":7358
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   320
         Width           =   360
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   6
         Left            =   5040
         MouseIcon       =   "ICEE_Calender.ctx":74AA
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   75
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   4080
         MouseIcon       =   "ICEE_Calender.ctx":75FC
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   75
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   3480
         MouseIcon       =   "ICEE_Calender.ctx":774E
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   75
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   2640
         MouseIcon       =   "ICEE_Calender.ctx":78A0
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   75
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   1800
         MouseIcon       =   "ICEE_Calender.ctx":79F2
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   75
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   960
         MouseIcon       =   "ICEE_Calender.ctx":7B44
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   75
         Width           =   270
      End
      Begin VB.Label lblYangLi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005273E8&
         Height          =   240
         Index           =   0
         Left            =   240
         MouseIcon       =   "ICEE_Calender.ctx":7C96
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   80
         Width           =   285
      End
      Begin VB.Shape shapeNow 
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H001C1113&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   4440
         Top             =   2520
         Width           =   540
      End
      Begin VB.Shape ShapeSelect 
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H007E5502&
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   4440
         Top             =   1920
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "ICEE_Calender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Event Click()
Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public mDay As Integer
Public mMonth As Integer
Public mYear As Integer
Public mHwnd As Long
Public mNDay As String
Public mNmonth As String
Public mYLongDate As Date
Public mNLongDate As String
Dim pos As POINTAPI '定义这个变量是取得鼠标坐标
Private Const ylData = "AB500D2,4BD0883," _
        & "4AE00DB,A5700D0,54D0581,D2600D8,D9500CC,655147D,56A00D5,9AD00CA,55D027A,4AE00D2," _
        & "A5B0682,A4D00DA,D2500CE,D25157E,B5500D6,56A00CC,ADA027B,95B00D3,49717C9,49B00DC," _
        & "A4B00D0,B4B0580,6A500D8,6D400CD,AB5147C,2B600D5,95700CA,52F027B,49700D2,6560682," _
        & "D4A00D9,EA500CE,6A9157E,5AD00D6,2B600CC,86E137C,92E00D3,C8D1783,C9500DB,D4A00D0," _
        & "D8A167F,B5500D7,56A00CD,A5B147D,25D00D5,92D00CA,D2B027A,A9500D2,B550781,6CA00D9," _
        & "B5500CE,535157F,4DA00D6,A5B00CB,457037C,52B00D4,A9A0883,E9500DA,6AA00D0,AEA0680," _
        & "AB500D7,4B600CD,AAE047D,A5700D5,52600CA,F260379,D9500D1,5B50782,56A00D9,96D00CE," _
        & "4DD057F,4AD00D7,A4D00CB,D4D047B,D2500D3,D550883,B5400DA,B6A00CF,95A1680,95B00D8," _
        & "49B00CD,A97047D,A4B00D5,B270ACA,6A500DC,6D400D1,AF40681,AB600D9,93700CE,4AF057F," _
        & "49700D7,64B00CC,74A037B,EA500D2,6B50883,5AC00DB,AB600CF,96D0580,92E00D8,C9600CD," _
        & "D95047C,D4A00D4,DA500C9,755027A,56A00D1,ABB0781,25D00DA,92D00CF,CAB057E,A9500D6," _
        & "B4A00CB,BAA047B,AD500D2,55D0983,4BA00DB,A5B00D0,5171680,52B00D8,A9300CD,795047D," _
        & "6AA00D4,AD500C9,5B5027A,4B600D2,96E0681,A4E00D9,D2600CE,EA6057E,D5300D5,5AA00CB," _
        & "76A037B,96D00D3,4AB0B83,4AD00DB,A4D00D0,D0B1680,D2500D7,D5200CC,DD4057C,B5A00D4," _
        & "56D00C9,55B027A,49B00D2,A570782,A4B00D9,AA500CE,B25157E,6D200D6,ADA00CA,4B6137B," _
        & "93700D3,49F08C9,49700DB,64B00D0,68A1680,EA500D7,6AA00CC,A6C147C,AAE00D4,92E00CA," _
        & "D2E0379,C9600D1,D550781,D4A00D9,DA400CD,5D5057E,56A00D6,A6C00CB,55D047B,52D00D3," _
        & "A9B0883,A9500DB,B4A00CF,B6A067F,AD500D7,55A00CD,ABA047C,A5A00D4,52B00CA,B27037A," _
        & "69300D1,7330781,6AA00D9,AD500CE,4B5157E,4B600D6,A5700CB,54E047C,D1600D2,E960882," _
        & "D5200DA,DAA00CF,6AA167F,56D00D7,4AE00CD,A9D047D,A2D00D4,D1500C9,F250279,D5200D1"

Private Const ylMd0 = "初一初二初三初四初五初六初七初八初九初十十一十二十三十四十五" _
        & "十六十七十八十九二十廿一廿二廿三廿四廿五廿六廿七廿八廿九三十 "

Private Const ylMn0 = "正二三四五六七八九十冬腊"
Private Const ylTianGan0 = "甲乙丙丁戊已庚辛壬癸"
Private Const ylDiZhi0 = "子丑寅卯辰巳午未申酉戌亥"
Private Const ylShu0 = "鼠牛虎兔龙蛇马羊猴鸡狗猪"


Dim SolarTerms(1960 To 2060) As String
Dim JiaZi60(1 To 60) As String

Private Type LunarData

    lYear As String * 6
    lMonth As String * 6
    lDay As String * 6
    LSX As String * 2
    LMonth_GZ As String * 6
    LDay_GZ As String * 6
    LJieQi As String * 4

End Type

Public Sub SolarTerm_Ini()
'===========================================================================================
'每三个字符代表一个月的两个节气，第一个数为当月第一个节气日期如621表示第一个节气为6号，第二个节气为21号
'========================================================================================

    SolarTerms(1960) = "621519520520521621723723723823722722"
    SolarTerms(1961) = "520419621520621621723823823823722722"
    SolarTerms(1962) = "620419621520621622723823823924823722"
    SolarTerms(1963) = "621419621521622622823824824924823822"
    SolarTerms(1964) = "621519520520521621723723723823722722"
    SolarTerms(1965) = "520419621520621621723823823823722722"
    SolarTerms(1966) = "620419621520621622723823823924823722"
    SolarTerms(1967) = "621419621521622622823824824924823822"
    SolarTerms(1968) = "621519520520521521723723723823722722"
    SolarTerms(1969) = "520419621520621621723823823823722722"
    SolarTerms(1970) = "620419621520621622723823823923823722"
    SolarTerms(1971) = "621419621521622622823824824924823822"
    SolarTerms(1972) = "621519520520521521723723723823722722"
    SolarTerms(1973) = "520419621520521621723823823823722722"
    SolarTerms(1974) = "620419621520621622723823823924823722"
    SolarTerms(1975) = "621419621521622622823824823924823822"
    SolarTerms(1976) = "621519520420521521723723723823722722"
    SolarTerms(1977) = "520419621520521621723723823823722722"
    SolarTerms(1978) = "620419621520621622723823823824823722"
    SolarTerms(1979) = "620419621521621622823824823924823822"
    SolarTerms(1980) = "621519520420521521723723723823722722"
    SolarTerms(1981) = "520419621520521621723723823823722722"
    SolarTerms(1982) = "620419621520621622723823823824822722"
    SolarTerms(1983) = "620419621520621622823824823924823822"
    SolarTerms(1984) = "621419520420521521722723723823722722"
    SolarTerms(1985) = "520419521520521621723723823823722722"
    SolarTerms(1986) = "520419621520621622723823923824822722"
    SolarTerms(1987) = "620419621520621622723824823924823822"
    SolarTerms(1988) = "621419350420521521722723723823722721"
    SolarTerms(1989) = "520419520420521621723723823823722722"
    SolarTerms(1990) = "520419621520621621723823823824822722"
    SolarTerms(1991) = "620419621520621622723823823924823722"
    SolarTerms(1992) = "621419520420521521722723723823722721"
    SolarTerms(1993) = "520418520520521621723723723823722722"
    SolarTerms(1994) = "520419621520621621723823823823822722"
    SolarTerms(1995) = "620419621520621622723823823924823722"
    SolarTerms(1996) = "621419520420521521722723723823722721"
    SolarTerms(1997) = "520418520520521521723723723823722722"
    SolarTerms(1998) = "520419621520621621723823823823722722"
    SolarTerms(1999) = "620419621520621622723823823924823722"
    SolarTerms(2000) = "621419520420521521722723723823722721"
    SolarTerms(2001) = "520418520520521521723723723823722722"
    SolarTerms(2002) = "520419621520621621723823823823722722"
    SolarTerms(2003) = "620419621520621622723823823924823722"
    SolarTerms(2004) = "621419520420521521722723723823722721"
    SolarTerms(2005) = "520418520520521521723723723823722722"
    SolarTerms(2006) = "520419621520521621723823823823722722"
    SolarTerms(2007) = "620419621520621622723823823924823722"
    SolarTerms(2008) = "621419520420521521722723723823722721"
    SolarTerms(2009) = "520418520420521521723723723823722722"
    SolarTerms(2010) = "520419621520521621723723823823722722"
    SolarTerms(2011) = "620419621520621622723823823824823722"
    SolarTerms(2012) = "621419520420520521722723722823722721"
    SolarTerms(2013) = "520418520420521521722723723823722722"    '2013
    SolarTerms(2014) = "520419621520524621723723823823722722"    '2014
    SolarTerms(2015) = "620419621520621622723823823824822722"    '2015
    SolarTerms(2016) = "620419520419520521722723722823722721"    '2016
    SolarTerms(2017) = "520318520420521521722723723823722722"    '2017
    SolarTerms(2018) = "520419521520521621723723823823722722"    '2018
    SolarTerms(2019) = "520419621520621622723823823824822722"    '2019
    SolarTerms(2020) = "620419520419520521622723722823722621"    '2020
    SolarTerms(2021) = "520318520420521521722723723823722721"    '2021
    SolarTerms(2022) = "520419520520521621723723823823722722"    '2022
    SolarTerms(2023) = "520419621520621621723823823824822722"    '2023
    SolarTerms(2024) = "620419520419520521622722722823722621"    '2024
    SolarTerms(2025) = "520318520420521521722723723823722721"    '2025
    SolarTerms(2026) = "520418520520521621723723723823722722"    '2026
    SolarTerms(2027) = "520419621520621621723823823823822722"    '2027
    SolarTerms(2028) = "620419520419520521622722722823722621"    '2028
    SolarTerms(2029) = "520318520420521521722723723823722721"    '2029
    SolarTerms(2030) = "520418520520521521723723723823722722"    '2030
    SolarTerms(2031) = "520418621520621621723823823823722722"    '2031
    SolarTerms(2032) = "620419520419520521677722722823722621"    '2032
    SolarTerms(2033) = "520318520420521521722723723823722721"    '2033
    SolarTerms(2034) = "520418520520521521723723723823722722"    '2034
    SolarTerms(2035) = "520419621520621621723823823823722722"    '2035
    SolarTerms(2036) = "620419520419520521622722722823722621"    '2036
    SolarTerms(2037) = "520318520420521521722723723823722721"    '2037
    SolarTerms(2038) = "520418520520521521723723723823722722"    '2038
    SolarTerms(2039) = "520419621520521621723823823823722722"    '2039
    SolarTerms(2040) = "620419520419520521622722722823722621"    '2040
    SolarTerms(2041) = "520318520420521521722723722823722721"    '2041
    SolarTerms(2042) = "520418520420521521723723723823722722"    '2042
    SolarTerms(2043) = "520419621520521621723723823823722722"    '2043
    SolarTerms(2044) = "620419520419520521622722722723722621"    '2044
    SolarTerms(2045) = "520318520419520521722723722823722721"    '2045
    SolarTerms(2046) = "520418520420521521722723723823722722"    '2046
    SolarTerms(2047) = "520419621520521621723723823823722722"    '2047
    SolarTerms(2048) = "620419520419520520622722722723721621"    '2048
    SolarTerms(2049) = "519318520419520521622722722823722721"    '2049
    SolarTerms(2050) = "520318520420518521722723723823722722"    '2050
    SolarTerms(2051) = "520419520520521621723723723823722722"    '2051
    SolarTerms(2052) = "520419520419520520622722722723721621"    '2052
    SolarTerms(2053) = "519318520419520521622722722823722721"    '2053
    SolarTerms(2054) = "520318520420521521722723723823722722"    '2054
    SolarTerms(2055) = "520419520520521521723723723823722722"    '2055
    SolarTerms(2056) = "520419520419520520622722722723721621"    '2056
    SolarTerms(2057) = "519318520419520521622722722823722621"    '2057
    SolarTerms(2058) = "520318520420521521722723723823722721"    '2058
    SolarTerms(2059) = "520419520520521521723723723823722722"    '2059
    SolarTerms(2060) = "520419520419520520622722722722721621"    '2060


    JiaZi60(1) = "甲子": JiaZi60(2) = "乙丑": JiaZi60(3) = "丙寅": JiaZi60(4) = "丁卯": JiaZi60(5) = "戊辰": JiaZi60(6) = "己巳"
    JiaZi60(7) = "庚午": JiaZi60(8) = "辛未": JiaZi60(9) = "壬申": JiaZi60(10) = "癸酉": JiaZi60(11) = "甲戌": JiaZi60(12) = "乙亥"
    JiaZi60(13) = "丙子": JiaZi60(14) = "丁丑": JiaZi60(15) = "戊寅": JiaZi60(16) = "己卯": JiaZi60(17) = "庚辰": JiaZi60(18) = "辛己"
    JiaZi60(19) = "壬午": JiaZi60(20) = "癸未": JiaZi60(21) = "甲申": JiaZi60(22) = "乙酉": JiaZi60(23) = "丙戌": JiaZi60(24) = "丁亥"
    JiaZi60(25) = "戊子": JiaZi60(26) = "己丑": JiaZi60(27) = "庚寅": JiaZi60(28) = "辛卯": JiaZi60(29) = "壬辰": JiaZi60(30) = "癸巳"
    JiaZi60(31) = "甲午": JiaZi60(32) = "乙未": JiaZi60(33) = "丙申": JiaZi60(34) = "丁酉": JiaZi60(35) = "戊戌": JiaZi60(36) = "己亥"
    JiaZi60(37) = "庚子": JiaZi60(38) = "辛丑": JiaZi60(39) = "壬寅": JiaZi60(40) = "癸丑": JiaZi60(41) = "甲辰": JiaZi60(42) = "乙巳"
    JiaZi60(43) = "丙午": JiaZi60(44) = "丁未": JiaZi60(45) = "戊申": JiaZi60(46) = "己酉": JiaZi60(47) = "庚戌": JiaZi60(48) = "辛亥"
    JiaZi60(49) = "壬子": JiaZi60(50) = "癸丑": JiaZi60(51) = "甲寅": JiaZi60(52) = "乙卯": JiaZi60(53) = "丙辰": JiaZi60(54) = "丁巳"
    JiaZi60(55) = "戊午": JiaZi60(56) = "己未": JiaZi60(57) = "庚申": JiaZi60(58) = "辛酉": JiaZi60(59) = "壬戌": JiaZi60(60) = "癸亥"
End Sub

'公历日期转农历
Public Function GetYLDate(ByVal strDate As String) As String()

    On Error GoTo aErr

    If Not IsDate(strDate) Then Exit Function

    Dim setDate As Date, tYear As Integer, tMonth As Integer, tDay As Integer
    setDate = CDate(strDate)
    tYear = Year(setDate): tMonth = Month(setDate): tDay = Day(setDate)

    '如果不是有效有日期，退出
    If tYear > 2100 Or tYear < 1900 Then Exit Function

    Dim daList() As String * 18, conDate As Date, thisMonths As String
    Dim AddYear As Integer, AddMonth As Integer, AddDay As Integer, getDay As Integer
    Dim YLyear As String, YLShuXing As String
    Dim dd0 As String, mm0 As String, GanZhi(0 To 59) As String * 2
    Dim RunYue As Boolean, RunYue1 As Integer, mDays As Integer, i As Integer

    '加载2年内的农历数据
    ReDim daList(tYear - 1 To tYear)
    daList(tYear - 1) = H2B(Mid(ylData, (tYear - 1900) * 8 + 1, 7))
    daList(tYear) = H2B(Mid(ylData, (tYear - 1900 + 1) * 8 + 1, 7))

    AddYear = tYear

initYL:

    AddMonth = CInt(Mid(daList(AddYear), 15, 2))
    AddDay = CInt(Mid(daList(AddYear), 17, 2))
    conDate = DateSerial(AddYear, AddMonth, AddDay)    '农历新年日期

    getDay = DateDiff("d", conDate, setDate) + 1    '相差天数
    If getDay < 1 Then AddYear = AddYear - 1: GoTo initYL

    thisMonths = Left(daList(AddYear), 14)
    RunYue1 = Val("&H" & Right(thisMonths, 1))    '闰月月份
    If RunYue1 > 0 Then    '有闰月
        thisMonths = Left(thisMonths, RunYue1) & Mid(thisMonths, 13, 1) & Mid(thisMonths, RunYue1 + 1)
    End If
    thisMonths = Left(thisMonths, 13)

    For i = 1 To 13    '计算天数
        mDays = 29 + CInt(Mid(thisMonths, i, 1))
        If getDay > mDays Then
            getDay = getDay - mDays
        Else
            If RunYue1 > 0 Then
                If i = RunYue1 + 1 Then RunYue = True
                If i > RunYue1 Then i = i - 1
            End If

            AddMonth = i
            AddDay = getDay
            Exit For
        End If
    Next

    dd0 = Mid(ylMd0, (AddDay - 1) * 2 + 1, 2)
    mm0 = Mid(ylMn0, AddMonth, 1) + "月"

    For i = 0 To 59
        GanZhi(i) = Mid(ylTianGan0, (i Mod 10) + 1, 1) + Mid(ylDiZhi0, (i Mod 12) + 1, 1)
    Next i

    YLyear = GanZhi((AddYear - 4) Mod 60)
    YLShuXing = Mid(ylShu0, ((AddYear - 4) Mod 12) + 1, 1)
    If RunYue Then mm0 = "闰" & mm0

    Dim C As Integer, Y As Integer, m As Integer, d As Integer, G As Integer, Z As Integer
    C = Left(tYear, 2)
    Y = Right(tYear, 2)
    m = tMonth
    d = tDay
    G = 4 * C + Int(C / 4) + 5 * Y + Int(Y / 4) + Int(3 * (m + 1) / 5) + d - 3
    G = G Mod 10
    If G = 0 Then G = 10
    Z = 8 * C + Int(C / 4) + 5 * Y + Int(Y / 4) + Int(3 * (m + 1) / 5) + d + 7 + IIf(m Mod 2 = 0, 6, 0)
    Z = Z Mod 12
    If Z = 0 Then Z = 12
    Dim res(7) As String
    res(0) = YLyear
    res(1) = YLShuXing
    res(2) = mm0
    res(3) = IIf(dd0 = "初一", mm0, dd0)
    res(4) = GetGanZhiMonth(YLyear, tYear, tMonth, tDay) & "月"
    res(5) = Mid(ylTianGan0, G, 1) & Mid(ylDiZhi0, Z, 1) & "日"
    res(6) = GetSolarTerms(tYear, tMonth, tDay)
    res(7) = GetSolarHoliday(mm0 & dd0)
    GetYLDate = res

aErr:

End Function


'农历转公历日期
'secondMonth 为真，则天示当 tMonth 是闰月时，取第二个月
Public Function GetDate(ByVal tYear As Integer, tMonth As Integer, tDay As Integer, Optional secondMonth As Boolean = False) As String

    On Error GoTo aErr

    If tYear > 2100 Or tYear < 1899 Or tMonth > 12 Or tMonth < 1 Or tDay > 30 Or tDay < 1 Then Exit Function

    Dim thisMonths As String, ylNewYear As Date, toMonth As Integer
    Dim mDays As Integer, RunYue1 As Integer, i As Integer
    thisMonths = H2B(Mid(ylData, (tYear - 1899) * 8 + 1, 7))

    If tDay > 29 + CInt(Mid(thisMonths, tMonth, 1)) Then Exit Function

    ylNewYear = DateSerial(tYear, CInt(Mid(thisMonths, 15, 2)), CInt(Mid(thisMonths, 17, 2)))    '农历新年日期

    thisMonths = Left(thisMonths, 14)
    RunYue1 = Val("&H" & Right(thisMonths, 1))    '闰月月份

    toMonth = tMonth - 1
    If RunYue1 > 0 Then    '有闰月
        thisMonths = Left(thisMonths, RunYue1) & Mid(thisMonths, 13, 1) & Mid(thisMonths, RunYue1 + 1)
        If tMonth > RunYue1 Or (secondMonth And tMonth = RunYue1) Then toMonth = tMonth
    End If
    thisMonths = Left(thisMonths, 13)

    mDays = 0
    For i = 1 To toMonth
        mDays = mDays + 29 + CInt(Mid(thisMonths, i, 1))
    Next
    mDays = mDays + tDay


    GetDate = ylNewYear + mDays - 1

aErr:

End Function

Public Function GetRunYue(tYear As Integer) As Integer
    Dim thisMonths As String
    thisMonths = H2B(Mid(ylData, (tYear - 1899) * 8 + 1, 7))

    thisMonths = Left(thisMonths, 14)
    GetRunYue = Val("&H" & Right(thisMonths, 1))
End Function

Public Function IsBigMonth(tYear As Integer, tMonth As Integer) As Boolean
    Dim thisMonths As String
    thisMonths = H2B(Mid(ylData, (tYear - 1899) * 8 + 1, 7))
    IsBigMonth = Mid(thisMonths, tMonth, 1)
End Function

Public Function GetNongLiNewYear(tYear As Integer) As Date
    Dim thisMonths As String
    thisMonths = H2B(Mid(ylData, (tYear - 1899) * 8 + 1, 7))
    GetNongLiNewYear = CDate(tYear & "-" & Mid(thisMonths, 15, 2) & "-" & Right(thisMonths, 2))
End Function

'将压缩的阴历字符还原
Private Function H2B(ByVal strHex As String) As String
    Dim i As Integer, I1 As Integer, tmpV As String
    Const hStr = "0123456789ABCDEF"
    Const bStr = "0000000100100011010001010110011110001001101010111100110111101111"

    tmpV = UCase(Left(strHex, 3))

    '十六进制转二进制
    For i = 1 To Len(tmpV)
        I1 = InStr(hStr, Mid(tmpV, i, 1))
        H2B = H2B & Mid(bStr, (I1 - 1) * 4 + 1, 4)
    Next

    H2B = H2B & Mid(strHex, 4, 2)

    '十六进制转十进制
    H2B = H2B & "0" & CStr(Val("&H" & Right(strHex, 2)))
End Function

Private Function GetSolarTerms(iYear As Integer, iMonth As Integer, iDay As Integer) As String
    If iYear > 2060 Or iYear < 1960 Then Exit Function
    Dim s As String
    s = SolarTerms(iYear)
    s = Mid(s, (iMonth - 1) * 3 + 1, 3)
    If iDay <> Val(Left(s, 1)) And iDay <> Val(Right(s, 2)) Then
        GetSolarTerms = ""
    Else

        Select Case iMonth
            Case 1
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "小寒", "大寒")
            Case 2
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "立春", "雨水")
            Case 3
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "惊蛰", "春分")
            Case 4
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "清明", "谷雨")
            Case 5
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "立夏", "小满")
            Case 6
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "芒种", "夏至")
            Case 7
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "小暑", "大暑")
            Case 8
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "立秋", "处暑")
            Case 9
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "白露", "秋分")
            Case 10
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "寒露", "霜降")
            Case 11
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "立冬", "小雪")
            Case 12
                GetSolarTerms = IIf(iDay = Val(Left(s, 1)), "大雪", "冬至")
        End Select
    End If

End Function

Private Function GetSolarHoliday(SolarDate As String) As String
    Select Case SolarDate
        Case "正月初一"
            GetSolarHoliday = "春节"
        Case "正月十五"
            GetSolarHoliday = "元宵"
        Case "五月初五"
            GetSolarHoliday = "端午"
        Case "七月初七"
            GetSolarHoliday = "七夕"
        Case "七月十五"
            GetSolarHoliday = "中元"
        Case "八月十五"
            GetSolarHoliday = "中秋"
        Case "九月初九"
            GetSolarHoliday = "重阳"
        Case "腊月初八"
            GetSolarHoliday = "腊八"
        Case "腊月廿四"
            GetSolarHoliday = "小年"
        Case "腊月三十"
            GetSolarHoliday = "除夕"
        Case Else
            GetSolarHoliday = ""
    End Select

End Function

Private Function GetGanZhiMonth(NYear As String, YYear As Integer, YMonth As Integer, YDay As Integer) As String

    Dim tZYGZ As Integer

    Dim i As Integer
    Dim t As Integer
    For i = 1 To 60
        If NYear = JiaZi60(i) Then
            t = i
            Exit For
        End If
    Next

    i = t Mod 10
    Select Case i
        Case 0
            tZYGZ = 51
        Case 1
            tZYGZ = 3
        Case 2
            tZYGZ = 15
        Case 3
            tZYGZ = 27
        Case 4
            tZYGZ = 39
        Case 5
            tZYGZ = 51
        Case 6
            tZYGZ = 3
        Case 7
            tZYGZ = 15
        Case 8
            tZYGZ = 27
        Case 9
            tZYGZ = 39
    End Select
    Dim n As Integer
    n = Val(Mid(SolarTerms(YYear), (YMonth - 1) * 3 + 1, 1))
    If YDay >= n Then
        Dim K As Integer
        K = tZYGZ + YMonth - 2
        If K > 60 Then K = K - 60

    Else
        K = tZYGZ + YMonth - 3
    End If
    GetGanZhiMonth = JiaZi60(K)
End Function

Public Function GetNMonthDay(YDate As String) As String
    Dim strNongLi() As String
    strNongLi = GetYLDate(YDate)
    GetNMonthDay = strNongLi(2) & strNongLi(3)
End Function



Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstYear.Visible = True Then lstYear.Visible = False
If lstMonth.Visible = True Then lstMonth.Visible = False
    
End Sub

Private Sub ImgDown_Click()

    If Val(lblM.Caption) <= 1 Then
        If Val(lblY.Caption) <= 1900 Then Exit Sub
        lblY.Caption = lblY.Caption - 1
        lblM.Caption = "12"
    Else
        lblM.Caption = lblM.Caption - 1
    End If
    shapeNow.Visible = False
    ShapeSelect.Visible = False
    LoadDayList Val(lblY.Caption), Val(lblM.Caption), 1
    mYear = Val(lblY.Caption)
    mMonth = Val(lblM.Caption)
If lstYear.Visible = True Then lstYear.Visible = False
If lstMonth.Visible = True Then lstMonth.Visible = False
    
End Sub

Private Sub imgUp_Click()
If lstYear.Visible = True Then lstYear.Visible = False
If lstMonth.Visible = True Then lstMonth.Visible = False
    
    If Val(lblM.Caption) >= 12 Then
        If Val(lblY.Caption) >= 2100 Then Exit Sub
        lblY.Caption = lblY.Caption + 1
        lblM.Caption = "1"
    Else
        lblM.Caption = lblM.Caption + 1
    End If
    shapeNow.Visible = False
    ShapeSelect.Visible = False
    LoadDayList Val(lblY.Caption), Val(lblM.Caption), 1

End Sub

Private Sub Label4_Click()
    lstMonth.Visible = True
End Sub

Private Sub lblBackToday_Click()

    LoadDayList Year(Now), Month(Now), Day(Now)
    lblY.Caption = Year(Now)
    lblM.Caption = Month(Now)
    Dim tFirst As Integer
    Dim t As Integer
    tFirst = WeekDay(CDate(Year(Now) & "-" & Month(Now) & "-1"))
    t = Day(Now) + tFirst - 2
    shapeNow.Visible = True
    shapeNow.Left = lblNongLi(t).Left + (lblNongLi(t).Width - 540) / 2    '- (lblNongLi(t).Left - 360) / 2 - 70
    shapeNow.Top = lblYangLi(t).Top - 35
    ShapeSelect.Left = lblNongLi(t).Left + (lblNongLi(t).Width - 540) / 2
    ShapeSelect.Top = lblYangLi(t).Top - 35
    '    lblYangLi_Click (t)

End Sub

Private Sub lblM_Change()
    mMonth = Val(lblM.Caption)
End Sub

Private Sub lblM_Click()
    lstMonth.Visible = True
End Sub

Private Sub lblNongLi_Click(Index As Integer)
    ShapeSelect.Left = lblNongLi(Index).Left + (lblNongLi(Index).Width - 540) / 2
    ShapeSelect.Top = lblYangLi(Index).Top - 35
    ShapeSelect.Visible = True
    mDay = Val(lblYangLi(Index).Caption)

    Dim strLongLi() As String
    strLongLi = GetYLDate(mYear & "-" & mMonth & "-" & mDay)
    mNDay = strLongLi(3)
    mNmonth = strLongLi(2)
    mYLongDate = CDate(mYear & "-" & mMonth & "-" & mDay)
    mNLongDate = mNmonth & mNDay
    RaiseEvent Click

End Sub

Private Sub lblY_Change()
    mYear = Val(lblY.Caption)
End Sub

Private Sub lblY_Click()
    lstYear.Visible = True
End Sub

Private Sub lblYangLi_Click(Index As Integer)
    ShapeSelect.Left = lblNongLi(Index).Left + (lblNongLi(Index).Width - 540) / 2
    ShapeSelect.Top = lblYangLi(Index).Top - 35
    ShapeSelect.Visible = True
    mDay = Val(lblYangLi(Index).Caption)

    Dim strLongLi() As String
    strLongLi = GetYLDate(mYear & "-" & mMonth & "-" & mDay)
    mNDay = strLongLi(3)
    mNmonth = strLongLi(2)

    mYLongDate = CDate(mYear & "-" & mMonth & "-" & mDay)
    mNLongDate = mNmonth & mNDay

    RaiseEvent Click

End Sub

Private Sub lstMonth_Click()
    lblM.Caption = lstMonth.List(lstMonth.ListIndex)
    lstMonth.Visible = False
    shapeNow.Visible = False
    ShapeSelect.Visible = False
    LoadDayList Val(lblY.Caption), Val(lblM.Caption), 1
End Sub

Private Sub lstYear_Click()
    lblY.Caption = lstYear.List(lstYear.ListIndex)
    lstYear.Visible = False
    shapeNow.Visible = False
    ShapeSelect.Visible = False
    LoadDayList Val(lblY.Caption), Val(lblM.Caption), 1
End Sub
Private Sub UserControl_Initialize()

On Error Resume Next
    Dim i As Integer
    Dim LAB As Control
    For Each LAB In UserControl.Controls
    If TypeOf LAB Is Label Then
    LAB.FONTNANE = "微软雅黑"
    LAB.MousePointer = 0
    End If
    Next
    lblTop(0).Left = 200
    lblNongLi(0).Caption = "初一一"
    lblYangLi(0).Caption = "1"
    lblNongLi(0).Left = lblTop(0).Left + (lblTop(0).Width - lblNongLi(0).Width) / 2
    lblYangLi(0).Left = lblTop(0).Left + (lblTop(0).Width - lblYangLi(0).Width) / 2

    lblNongLi(7).Caption = "初一八"
    lblYangLi(7).Caption = "8"
    lblNongLi(7).Left = lblTop(0).Left + (lblTop(0).Width - lblNongLi(7).Width) / 2
    lblYangLi(7).Left = lblTop(0).Left + (lblTop(0).Width - lblYangLi(7).Width) / 2

    lblNongLi(14).Caption = "十一五"
    lblYangLi(14).Caption = "15"
    lblNongLi(14).Left = lblTop(0).Left + (lblTop(0).Width - lblNongLi(14).Width) / 2
    lblYangLi(14).Left = lblTop(0).Left + (lblTop(0).Width - lblYangLi(14).Width) / 2

    lblNongLi(21).Caption = "廿一二"
    lblYangLi(21).Caption = "22"
    lblNongLi(21).Left = lblTop(0).Left + (lblTop(0).Width - lblNongLi(21).Width) / 2
    lblYangLi(21).Left = lblTop(0).Left + (lblTop(0).Width - lblYangLi(21).Width) / 2

    lblNongLi(28).Caption = "廿一九"
    lblYangLi(28).Caption = "29"
    lblNongLi(28).Left = lblTop(0).Left + (lblTop(0).Width - lblNongLi(28).Width) / 2
    lblYangLi(28).Left = lblTop(0).Left + (lblTop(0).Width - lblYangLi(28).Width) / 2

    lblNongLi(35).Caption = "廿一九"
    lblYangLi(35).Caption = "29"
    lblNongLi(35).Left = lblTop(0).Left + (lblTop(0).Width - lblNongLi(35).Width) / 2
    lblYangLi(35).Left = lblTop(0).Left + (lblTop(0).Width - lblYangLi(35).Width) / 2

    For i = 1 To 6
        lblTop(i).Top = lblTop(0).Top
        lblTop(i).Left = lblTop(i - 1).Left + 800

        lblNongLi(i).Caption = "初一一"
        lblYangLi(i).Caption = i + 1
        lblNongLi(i).Left = lblTop(i).Left + (lblTop(i).Width - lblNongLi(i).Width) / 2
        lblYangLi(i).Left = lblTop(i).Left + (lblTop(i).Width - lblYangLi(i).Width) / 2

        lblNongLi(i + 7).Caption = "初一一"
        lblYangLi(i + 7).Caption = i + 8
        lblNongLi(i + 7).Left = lblTop(i).Left + (lblTop(i).Width - lblNongLi(i + 7).Width) / 2
        lblYangLi(i + 7).Left = lblTop(i).Left + (lblTop(i).Width - lblYangLi(i + 7).Width) / 2

        lblNongLi(i + 14).Caption = "初一一"
        lblYangLi(i + 14).Caption = i + 15
        lblNongLi(i + 14).Left = lblTop(i).Left + (lblTop(i).Width - lblNongLi(i + 14).Width) / 2
        lblYangLi(i + 14).Left = lblTop(i).Left + (lblTop(i).Width - lblYangLi(i + 14).Width) / 2

        lblNongLi(i + 21).Caption = "初一一"
        lblYangLi(i + 21).Caption = i + 22
        lblNongLi(i + 21).Left = lblTop(i).Left + (lblTop(i).Width - lblNongLi(i + 21).Width) / 2
        lblYangLi(i + 21).Left = lblTop(i).Left + (lblTop(i).Width - lblYangLi(i + 21).Width) / 2

        lblNongLi(i + 28).Caption = "初一一"
        lblYangLi(i + 28).Caption = i + 29
        lblNongLi(i + 28).Left = lblTop(i).Left + (lblTop(i).Width - lblNongLi(i + 28).Width) / 2
        lblYangLi(i + 28).Left = lblTop(i).Left + (lblTop(i).Width - lblYangLi(i + 28).Width) / 2

        lblNongLi(i + 35).Caption = "初一一"
        lblYangLi(i + 35).Caption = i + 36
        lblNongLi(i + 35).Left = lblTop(i).Left + (lblTop(i).Width - lblNongLi(i + 35).Width) / 2
        lblYangLi(i + 35).Left = lblTop(i).Left + (lblTop(i).Width - lblYangLi(i + 35).Width) / 2
    Next
    '======================================================
    lblY.Caption = Year(Now)
    lblM.Caption = Month(Now)
    LoadDayList Year(Now), Month(Now), Day(Now)

    For i = 1960 To 2060
        lstYear.AddItem i
    Next

    For i = 1 To 12
        lstMonth.AddItem i
    Next
    Dim tFirst As Integer
    Dim t As Integer
    tFirst = WeekDay(CDate(Year(Now) & "-" & Month(Now) & "-1"))
    t = Day(Now) + tFirst - 2
    shapeNow.Visible = True
    shapeNow.Left = lblNongLi(t).Left + (lblNongLi(t).Width - 540) / 2
    shapeNow.Top = lblYangLi(t).Top - 35
    ShapeSelect.Left = lblNongLi(t).Left + (lblNongLi(t).Width - 540) / 2

    mHwnd = UserControl.hwnd
    mYear = Year(Now)
    mMonth = Month(Now)
    mDay = Day(Now)

    Dim strLongLi() As String
    strLongLi = GetYLDate(mYear & "-" & mMonth & "-" & mDay)
    mNDay = strLongLi(3)
    mNmonth = strLongLi(2)
    mYLongDate = CDate(mYear & "-" & mMonth & "-" & mDay)
    mNLongDate = mNmonth & mNDay
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstYear.Visible = True Then lstYear.Visible = False
If lstMonth.Visible = True Then lstMonth.Visible = False
    
    RaiseEvent MOUSEMOVE(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 5565
    UserControl.Height = 4440
End Sub

Private Function GetDays(tYear As Integer, tMonth As Integer) As Integer
    Select Case tMonth
        Case 1
            GetDays = 31
        Case 3
            GetDays = 31
        Case 4
            GetDays = 30
        Case 5
            GetDays = 31
        Case 6
            GetDays = 30
        Case 7
            GetDays = 31
        Case 8
            GetDays = 31
        Case 9
            GetDays = 30
        Case 10
            GetDays = 31
        Case 11
            GetDays = 30
        Case 12
            GetDays = 31
        Case 2
            If (tYear Mod 4 = 0 And tYear Mod 100 <> 0) Or (tYear Mod 400 = 0) Then
                GetDays = 29
            Else
                GetDays = 28
            End If
    End Select
End Function

Private Sub LoadDayList(tYear As Integer, tMonth As Integer, tDay As Integer)
    On Error Resume Next
    Dim tFirst As Integer
    Dim tDays As Integer
    Dim t As Integer
    t = 1
    Dim i As Integer
    tFirst = WeekDay(CDate(tYear & "-" & tMonth & "-1"))
    tDays = GetDays(tYear, tMonth) + tFirst - 2

    If tFirst > 1 Then
        For i = 0 To tFirst - 2
            lblNongLi(i).Visible = False
            lblYangLi(i).Visible = False
        Next
    End If

    Dim strLongLi() As String
    For i = tFirst - 1 To tDays

        lblYangLi(i).Visible = True
        lblYangLi(i).Caption = t
        lblNongLi(i).Visible = True
        strLongLi = GetYLDate(tYear & "-" & tMonth & "-" & t)
        If strLongLi(6) = "" Then
            lblNongLi(i).Caption = strLongLi(3)
            If i Mod 7 = 0 Or (i + 1) Mod 7 = 0 Then
                lblNongLi(i).FOREColor = &H1F1FE2
            Else
                lblNongLi(i).FOREColor = &HC0C0C0
            End If
        Else
            lblNongLi(i).Caption = strLongLi(6)
            lblNongLi(i).FOREColor = &H2EBC7C
        End If

        If strLongLi(7) <> "" Then
            lblNongLi(i).Caption = strLongLi(7)
            lblNongLi(i).FOREColor = &H2EBC7C
        End If

        t = t + 1

    Next
    tDays = tDays + 1
    For i = tDays To 41
        lblNongLi(i).Visible = False
        lblYangLi(i).Visible = False
    Next
End Sub


Private Sub TMOUT_Timer()
Dim r As RECT, p As POINTAPI '鼠标移出/移入透明值得改变
GetCursorPos pos
Call 获得生肖(mYear, LBSX)
Call CHECK_XZ
If p.X < r.Left Or p.X > r.Right Or p.Y < r.Top Or p.Y > r.Bottom Then '移出界面
If lstYear.Visible = True Then lstYear.Visible = False
If lstMonth.Visible = True Then lstMonth.Visible = False
End If
End Sub
Public Sub CHECK_XZ() '检测你是什么星座
Dim a As Integer, b As Integer, C As Integer
  a = mMonth
  b = mDay
  C = a * 100 + b
  
  LBXZ.Caption = "魔羯座"
  
  If C > 1221 Then C = 0
   
  If C > 119 Then
    LBXZ.Caption = "水瓶座"
  End If
   
  If C > 218 Then
    LBXZ.Caption = "双鱼座"
  End If
   
  If C > 320 Then
    LBXZ.Caption = "白羊座"
  End If
   
  If C > 420 Then
    LBXZ.Caption = "金牛座"
  End If
   
  If C > 520 Then
    LBXZ.Caption = "双子座"
  End If
   
  If C > 621 Then
    LBXZ.Caption = "巨蟹座"
     
  End If
   
  If C > 722 Then
    LBXZ.Caption = "狮子座"
  End If
   
  If C > 822 Then
    LBXZ.Caption = "处女座"
  End If
   
  If C > 922 Then
    LBXZ.Caption = "天秤座"
  End If
   
  If C > 1022 Then
    LBXZ.Caption = "天蝎座"
  End If
   
  If C > 1121 Then
    LBXZ.Caption = "人马座"
  End If
End Sub
Private Sub 获得生肖(Year As Integer, SHOWINFO As Label)
  Dim name As Integer
  name = Year Mod 12
  Select Case name
    Case 4
      SHOWINFO.Caption = "鼠年"
    Case 5
      SHOWINFO.Caption = "牛年"
    Case 6
      SHOWINFO.Caption = "虎年"
    Case 7
      SHOWINFO.Caption = "兔年"
    Case 8
      SHOWINFO.Caption = "龙年"
    Case 9
      SHOWINFO.Caption = "蛇年"
    Case 10
     SHOWINFO.Caption = "马年"
    Case 11
      SHOWINFO.Caption = "羊年"
    Case 0
     SHOWINFO.Caption = "猴年"
    Case 1
      SHOWINFO.Caption = "鸡年"
    Case 2
     SHOWINFO.Caption = "狗年"
    Case 3
     SHOWINFO.Caption = "猪年"
   End Select
End Sub

