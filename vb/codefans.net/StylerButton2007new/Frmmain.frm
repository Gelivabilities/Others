VERSION 5.00
Begin VB.Form Frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11895
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Frmmain.frx":4B42
   ScaleHeight     =   413
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   793
   StartUpPosition =   2  '屏幕中心
   Begin StylerButton2007.StylerButton StylerButton9 
      Height          =   795
      Left            =   8400
      TabIndex        =   25
      Top             =   2040
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   1402
      Caption         =   "Vista RC2"
      ForeColor       =   16576
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      IconDisableColor=   11711154
      Theme           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton8 
      Height          =   795
      Left            =   8400
      TabIndex        =   24
      Top             =   1080
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1402
      Caption         =   "Office 2007"
      CaptionDisableColor=   13153946
      CaptionEffectColor=   16777215
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton7 
      Height          =   735
      Left            =   8400
      TabIndex        =   23
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      Caption         =   "WMP11"
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton6 
      Height          =   735
      Left            =   8400
      TabIndex        =   19
      Top             =   5160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      Caption         =   "退出"
      ForeColor       =   255
      CaptionDisableColor=   13153946
      CaptionEffectColor=   12632319
      CaptionEffect   =   4
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton5 
      Height          =   675
      Left            =   3000
      TabIndex        =   18
      Top             =   5280
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1191
      Caption         =   "关于 !"
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   12632256
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton4 
      Height          =   900
      Index           =   0
      Left            =   3240
      TabIndex        =   13
      Top             =   2145
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Caption         =   "20"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedValue    =   20
   End
   Begin StylerButton2007.StylerButton StylerButton3 
      Height          =   750
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Top             =   240
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1323
      Caption         =   "激活状态"
      ForeColor       =   33023
      CaptionDisableColor=   12236471
      CaptionEffectColor=   0
      CaptionEffect   =   4
      IconDisableColor=   11711154
      Theme           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "Frmmain.frx":10BF2C
   End
   Begin StylerButton2007.StylerButton StylerButton2 
      Height          =   510
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   900
      Caption         =   "正常"
      ForeColor       =   12582912
      CaptionDisableColor=   13153946
      CaptionEffectColor=   255
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton1 
      Height          =   720
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1270
      Caption         =   "WMP 11"
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   12632256
      CaptionEffect   =   4
      IconDisableColor=   12236471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   0
      Width           =   0
   End
   Begin StylerButton2007.StylerButton StylerButton1 
      Height          =   720
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1270
      Caption         =   "Vista RC2"
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   14737632
      CaptionEffect   =   4
      IconDisableColor=   11711154
      Theme           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton1 
      Height          =   720
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1270
      Caption         =   "Office 2007"
      ForeColor       =   16777215
      CaptionDisableColor=   13153946
      CaptionEffectColor=   14737632
      CaptionEffect   =   4
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton2 
      Height          =   510
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   900
      Caption         =   "浮雕样式"
      CaptionDisableColor=   13153946
      CaptionEffectColor=   16777215
      CaptionEffect   =   2
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton2 
      Height          =   510
      Index           =   2
      Left            =   135
      TabIndex        =   6
      Top             =   4080
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   900
      Caption         =   "雕刻样式"
      ForeColor       =   8421504
      CaptionDisableColor=   13153946
      CaptionEffectColor=   16777215
      CaptionEffect   =   3
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton2 
      Height          =   510
      Index           =   3
      Left            =   105
      TabIndex        =   7
      Top             =   4680
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   900
      Caption         =   "空心字体"
      ForeColor       =   65280
      CaptionDisableColor=   13153946
      CaptionEffectColor=   255
      CaptionEffect   =   4
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton2 
      Height          =   510
      Index           =   4
      Left            =   105
      TabIndex        =   8
      Top             =   5280
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   900
      Caption         =   "阴影样式"
      ForeColor       =   16576
      CaptionDisableColor=   13153946
      CaptionEffectColor=   12632256
      CaptionEffect   =   5
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "黑体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin StylerButton2007.StylerButton StylerButton3 
      Height          =   750
      Index           =   1
      Left            =   3360
      TabIndex        =   10
      Top             =   1080
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   1323
      Caption         =   "禁止状态"
      ForeColor       =   16777215
      CaptionDisableColor=   12236471
      CaptionEffectColor=   65280
      CaptionEffect   =   4
      IconDisableColor=   11711154
      Theme           =   3
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "Frmmain.frx":10C780
   End
   Begin StylerButton2007.StylerButton StylerButton4 
      Height          =   900
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Top             =   2175
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Caption         =   "50"
      ForeColor       =   255
      CaptionDisableColor=   13153946
      CaptionEffectColor=   16777215
      IconDisableColor=   13614497
      Theme           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedValue    =   50
   End
   Begin StylerButton2007.StylerButton StylerButton4 
      Height          =   900
      Index           =   2
      Left            =   5115
      TabIndex        =   15
      Top             =   2205
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1588
      Caption         =   "100"
      ForeColor       =   255
      CaptionDisableColor=   12236471
      CaptionEffectColor=   16777215
      IconDisableColor=   11711154
      Theme           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedValue    =   100
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "各种标题特效演示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2025
      Left            =   2520
      TabIndex        =   22
      Top             =   3255
      Width           =   300
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "按钮风格样式演示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2070
      Left            =   2475
      TabIndex        =   21
      Top             =   360
      Width           =   330
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frmmain.frx":10CFD4
      Height          =   1710
      Left            =   3240
      TabIndex        =   20
      Top             =   3315
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "支持自定义角度值."
      Height          =   270
      Left            =   6120
      TabIndex        =   17
      Top             =   2565
      Width           =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "圆角形状示例"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   6120
      TabIndex        =   16
      Top             =   2130
      Width           =   1860
   End
   Begin VB.Line Line9 
      X1              =   217
      X2              =   524
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "    可以使用自定义颜色显示."
      Height          =   555
      Left            =   6360
      TabIndex        =   12
      Top             =   1350
      Width           =   1230
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   419
      X2              =   524
      Y1              =   85
      Y2              =   85
   End
   Begin VB.Line Line8 
      Index           =   0
      X1              =   216
      X2              =   524
      Y1              =   12
      Y2              =   12
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "图标使用示例"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   915
      Left            =   6330
      TabIndex        =   11
      Top             =   270
      Width           =   1275
   End
   Begin VB.Line Line7 
      X1              =   160
      X2              =   160
      Y1              =   384
      Y2              =   192
   End
   Begin VB.Line Line6 
      X1              =   144
      X2              =   160
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Line Line5 
      X1              =   145
      X2              =   160
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Line Line4 
      X1              =   192
      X2              =   8
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line3 
      X1              =   148
      X2              =   157
      Y1              =   165
      Y2              =   165
   End
   Begin VB.Line Line2 
      X1              =   156
      X2              =   156
      Y1              =   8
      Y2              =   165
   End
   Begin VB.Line Line1 
      X1              =   148
      X2              =   156
      Y1              =   8
      Y2              =   8
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Private Sub Form_Load()
    Dim A As String
    A = App.Major & "." & App.Minor & "." & App.Revision
    Me.Caption = "真彩色时尚按钮控件 2007 vr." & A

End Sub

Private Sub StylerButton5_Click()
    FrmAbout.Show
End Sub

Private Sub StylerButton6_Click()
    End
End Sub
