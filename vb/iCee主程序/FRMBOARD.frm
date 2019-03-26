VERSION 5.00
Begin VB.Form FRMBOARD 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "涂鸦"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11610
   Icon            =   "FRMBOARD.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   774
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00586E74&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   720
      ScaleHeight     =   617
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   769
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   11535
      Begin VB.PictureBox IMODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Index           =   5
         Left            =   7800
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   69
         Top             =   5040
         Width           =   2895
         Begin ICEE.ICEE_COMMAND IOK 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   70
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   19
            Left            =   120
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   18
            Left            =   120
            Top             =   120
            Width           =   2655
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   17
            Left            =   1440
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   180
            Index           =   19
            Left            =   2520
            TabIndex        =   73
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   180
            Index           =   18
            Left            =   240
            TabIndex        =   72
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   180
            Index           =   17
            Left            =   240
            TabIndex        =   71
            Top             =   1440
            Width           =   90
         End
      End
      Begin VB.PictureBox IMODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Index           =   4
         Left            =   4440
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   64
         Top             =   5040
         Width           =   2895
         Begin ICEE.ICEE_COMMAND IOK 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   65
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   16
            Left            =   120
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   15
            Left            =   120
            Top             =   120
            Width           =   2655
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   14
            Left            =   1440
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   180
            Index           =   16
            Left            =   2520
            TabIndex        =   68
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   180
            Index           =   15
            Left            =   240
            TabIndex        =   67
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   180
            Index           =   14
            Left            =   240
            TabIndex        =   66
            Top             =   1440
            Width           =   90
         End
      End
      Begin VB.PictureBox IMODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Index           =   3
         Left            =   1200
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   59
         Top             =   5040
         Width           =   2895
         Begin ICEE.ICEE_COMMAND IOK 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   60
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   13
            Left            =   120
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   12
            Left            =   120
            Top             =   120
            Width           =   2655
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   6
            Left            =   1440
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   180
            Index           =   13
            Left            =   2520
            TabIndex        =   63
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   180
            Index           =   12
            Left            =   240
            TabIndex        =   62
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   180
            Index           =   4
            Left            =   240
            TabIndex        =   61
            Top             =   1440
            Width           =   90
         End
      End
      Begin VB.PictureBox PBK 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00586E74&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   58
         Top             =   120
         Width           =   975
      End
      Begin VB.PictureBox IMODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Index           =   2
         Left            =   7800
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   22
         Top             =   840
         Width           =   2895
         Begin ICEE.ICEE_COMMAND IOK 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   180
            Index           =   11
            Left            =   2520
            TabIndex        =   27
            Top             =   720
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   180
            Index           =   10
            Left            =   2520
            TabIndex        =   26
            Top             =   1560
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   180
            Index           =   9
            Left            =   2520
            TabIndex        =   25
            Top             =   2280
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   180
            Index           =   8
            Left            =   2520
            TabIndex        =   24
            Top             =   3000
            Width           =   90
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   855
            Index           =   11
            Left            =   120
            Top             =   120
            Width           =   2655
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   855
            Index           =   10
            Left            =   120
            Top             =   960
            Width           =   2655
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   735
            Index           =   9
            Left            =   120
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   735
            Index           =   8
            Left            =   120
            Top             =   2520
            Width           =   2655
         End
      End
      Begin VB.PictureBox IMODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Index           =   1
         Left            =   4440
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   17
         Top             =   840
         Width           =   2895
         Begin ICEE.ICEE_COMMAND IOK 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   180
            Index           =   7
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   180
            Index           =   6
            Left            =   240
            TabIndex        =   20
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   180
            Index           =   5
            Left            =   2520
            TabIndex        =   19
            Top             =   1800
            Width           =   90
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   4
            Left            =   1440
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   7
            Left            =   120
            Top             =   120
            Width           =   2655
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   5
            Left            =   120
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.PictureBox IMODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Index           =   0
         Left            =   1200
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   11
         Top             =   840
         Width           =   2895
         Begin ICEE.ICEE_COMMAND IOK 
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   180
            Index           =   3
            Left            =   1560
            TabIndex        =   15
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   13
            Top             =   240
            Width           =   90
         End
         Begin VB.Label LCM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   90
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   3
            Left            =   1440
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   2
            Left            =   120
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   1
            Left            =   1440
            Top             =   120
            Width           =   1335
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00D2F0F5&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Height          =   1575
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用须知:使用后会自动创建700×1500像素的图像,并将之前绘制的图像清空,创建后会自动绘制网格,背景颜色以背景颜色为准,网格颜色为黑色"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   9000
         Width           =   11340
      End
      Begin VB.Shape SB 
         BackColor       =   &H00252E31&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   0
         Top             =   8880
         Width           =   11535
      End
      Begin VB.Label LA 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "漫画模板"
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
         Index           =   0
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   2
      Left            =   120
      ScaleHeight     =   4095
      ScaleWidth      =   6630
      TabIndex        =   76
      Top             =   315
      Visible         =   0   'False
      Width           =   6630
      Begin VB.PictureBox PS 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00565656&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4065
         Index           =   0
         Left            =   15
         ScaleHeight     =   271
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   440
         TabIndex        =   77
         Top             =   15
         Width           =   6600
         Begin VB.PictureBox PSD 
            BackColor       =   &H001F1FE2&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   5280
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   66
            TabIndex        =   78
            ToolTipText     =   "屏幕取色"
            Top             =   960
            Width           =   990
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   8
            Left            =   5640
            TabIndex        =   96
            Top             =   3780
            Width           =   945
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   7
            Left            =   5640
            TabIndex        =   95
            Top             =   3540
            Width           =   945
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   6
            Left            =   5700
            TabIndex        =   94
            Top             =   3180
            Width           =   900
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   5700
            TabIndex        =   93
            Top             =   2940
            Width           =   900
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   5700
            TabIndex        =   92
            Top             =   2700
            Width           =   900
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   5700
            TabIndex        =   91
            Top             =   2460
            Width           =   900
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   5700
            TabIndex        =   90
            Top             =   2100
            Width           =   900
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   5700
            TabIndex        =   89
            Top             =   1860
            Width           =   900
         End
         Begin VB.TextBox t1 
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   5700
            TabIndex        =   88
            Top             =   1620
            Width           =   900
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00565656&
            Caption         =   "G:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   5070
            TabIndex        =   87
            Top             =   2940
            Width           =   495
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00565656&
            Caption         =   "R:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   5070
            TabIndex        =   86
            Top             =   2700
            Width           =   495
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00565656&
            Caption         =   "A:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   5070
            TabIndex        =   85
            Top             =   2460
            Width           =   495
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00565656&
            Caption         =   "B:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   5070
            TabIndex        =   84
            Top             =   2100
            Width           =   495
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00565656&
            Caption         =   "S:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   5070
            TabIndex        =   83
            Top             =   1860
            Width           =   495
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00565656&
            Caption         =   "H:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   5070
            TabIndex        =   82
            Top             =   1620
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.PictureBox p2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   990
            Left            =   5280
            ScaleHeight     =   66
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   66
            TabIndex        =   81
            ToolTipText     =   "画笔颜色"
            Top             =   120
            Width           =   990
         End
         Begin VB.PictureBox p1 
            BackColor       =   &H00231C09&
            BorderStyle     =   0  'None
            Height          =   4080
            Left            =   0
            MousePointer    =   2  'Cross
            ScaleHeight     =   272
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   328
            TabIndex        =   80
            Top             =   0
            Width           =   4920
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00565656&
            Caption         =   "B:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   6
            Left            =   5070
            TabIndex        =   79
            Top             =   3240
            Width           =   495
         End
         Begin VB.Image i0 
            Height          =   300
            Left            =   6480
            Picture         =   "FRMBOARD.frx":0802
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00201400&
            BackStyle       =   0  'Transparent
            Caption         =   "VB"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   3
            Left            =   5070
            TabIndex        =   98
            Top             =   3780
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00201400&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Index           =   2
            Left            =   5070
            TabIndex        =   97
            Top             =   3540
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox PICSOU 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7080
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   100
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7920
      Picture         =   "FRMBOARD.frx":098E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   8040
      Picture         =   "FRMBOARD.frx":0A72
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   8040
      Picture         =   "FRMBOARD.frx":0B56
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox PHELP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   761
      TabIndex        =   7
      Top             =   8640
      Width           =   11415
      Begin VB.Image IM_TYPE 
         Height          =   240
         Left            =   120
         Picture         =   "FRMBOARD.frx":0C3A
         Top             =   75
         Width           =   480
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RGB"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   840
         TabIndex        =   8
         Top             =   120
         Width           =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   0
         X2              =   240
         Y1              =   312
         Y2              =   312
      End
   End
   Begin VB.PictureBox PT 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F5E7D0&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   600
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox PF 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   75
      Width           =   255
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   180
      Width           =   255
   End
   Begin ICEE.ucScrollbar SGRO 
      Height          =   225
      Left            =   0
      TabIndex        =   5
      Top             =   8280
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   397
   End
   Begin ICEE.ucScrollbar SCRO 
      Height          =   6450
      Left            =   11370
      TabIndex        =   6
      Top             =   1830
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   11933
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6810
      Index           =   0
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   454
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   757
      TabIndex        =   4
      Top             =   960
      Width           =   11355
      Begin VB.PictureBox picRulerBaseH 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4500
         Left            =   0
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   150
         Width           =   255
         Begin VB.PictureBox picRulerH 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   4500
            Left            =   0
            ScaleHeight     =   300
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   17
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   0
            Width           =   255
            Begin VB.Shape shpMarkerH 
               BorderColor     =   &H005BB645&
               BorderWidth     =   2
               Height          =   30
               Left            =   90
               Top             =   120
               Width           =   180
            End
         End
      End
      Begin VB.PictureBox picRulerBaseW 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   270
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   400
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
         Width           =   6000
         Begin VB.PictureBox picRulerW 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   17
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   400
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   0
            Width           =   6000
            Begin VB.Shape shpMarkerW 
               BorderColor     =   &H005BB645&
               BorderWidth     =   2
               Height          =   180
               Left            =   150
               Top             =   90
               Width           =   15
            End
         End
      End
      Begin VB.PictureBox PICOPTION 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         Height          =   6015
         Left            =   6240
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   401
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   265
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   105
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin VB.PictureBox PVP 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00565656&
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   120
            ScaleHeight     =   121
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   56
            Top             =   4080
            Width           =   2535
            Begin VB.Image IPR 
               Height          =   855
               Left            =   840
               Stretch         =   -1  'True
               Top             =   480
               Width           =   975
            End
         End
         Begin ICEE.ICEE_KEY IZOOM 
            Height          =   495
            Index           =   0
            Left            =   1680
            TabIndex        =   42
            Top             =   720
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin VB.ComboBox CBSIZE 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00554513&
            Height          =   300
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1440
            Width           =   1215
         End
         Begin VB.ComboBox CBPEN 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00554513&
            Height          =   300
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXTY 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00554513&
            Height          =   270
            Left            =   1920
            TabIndex        =   35
            Text            =   "500"
            Top             =   315
            Width           =   615
         End
         Begin VB.TextBox TXTX 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00554513&
            Height          =   270
            Left            =   720
            TabIndex        =   34
            Text            =   "500"
            Top             =   315
            Width           =   615
         End
         Begin ICEE.ICEE_KEY IZOOM 
            Height          =   495
            Index           =   1
            Left            =   1680
            TabIndex        =   43
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICM 
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   57
            Top             =   2040
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   1
            Left            =   2760
            TabIndex        =   106
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   2
            Left            =   2760
            TabIndex        =   107
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   3
            Left            =   2760
            TabIndex        =   108
            Top             =   1560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   109
            Top             =   2040
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   5
            Left            =   2760
            TabIndex        =   110
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   6
            Left            =   2760
            TabIndex        =   111
            Top             =   3000
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   7
            Left            =   2760
            TabIndex        =   112
            Top             =   3480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin ICEE.ICEE_KEY ICT 
            Height          =   495
            Index           =   8
            Left            =   2760
            TabIndex        =   113
            Top             =   3960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "画笔粗度"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   240
            TabIndex        =   41
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "画笔硬度"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   720
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "高度"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   1440
            TabIndex        =   37
            Top             =   315
            Width           =   360
         End
         Begin VB.Label LA 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "宽度"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   315
            Width           =   360
         End
         Begin VB.Shape SPW 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            Height          =   1935
            Index           =   27
            Left            =   120
            Top             =   120
            Width           =   2655
         End
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   945
         Left            =   10200
         TabIndex        =   32
         Top             =   5280
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1667
      End
      Begin VB.PictureBox PO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6375
         Index           =   1
         Left            =   240
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   425
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   505
         TabIndex        =   48
         Top             =   240
         Width           =   7575
         Begin VB.PictureBox PICTY 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   6855
            Left            =   -240
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   455
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   463
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   -1080
            Width           =   6975
            Begin VB.PictureBox PCUT 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   1215
               Left            =   4080
               ScaleHeight     =   81
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   113
               TabIndex        =   74
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.PictureBox picTemp2 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   0
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.PictureBox picTemp1 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   30
               ScaleHeight     =   25
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   25
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   210
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.PictureBox PTXT 
               AutoRedraw      =   -1  'True
               BackColor       =   &H008BA31F&
               BorderStyle     =   0  'None
               Height          =   1095
               Left            =   1440
               ScaleHeight     =   73
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   169
               TabIndex        =   50
               Top             =   960
               Visible         =   0   'False
               Width           =   2535
               Begin VB.TextBox TXTU 
                  Appearance      =   0  'Flat
                  BackColor       =   &H008BA31F&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H00FFFFFF&
                  Height          =   480
                  Left            =   15
                  TabIndex        =   53
                  Top             =   0
                  Width           =   2415
               End
               Begin VB.ComboBox CBFS 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00554513&
                  Height          =   300
                  Index           =   1
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.ComboBox CBFS 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00554513&
                  Height          =   300
                  Index           =   0
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   51
                  Top             =   600
                  Width           =   1095
               End
            End
            Begin ICEE.ICEE_KEY ICM 
               Height          =   495
               Index           =   6
               Left            =   1560
               TabIndex        =   75
               Top             =   240
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   873
            End
            Begin VB.Shape SCUT 
               DrawMode        =   5  'Not Copy Pen
               Height          =   1215
               Left            =   480
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
         End
      End
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   4
      Left            =   9720
      TabIndex        =   99
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   101
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   102
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   103
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   104
      Top             =   480
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   5
      Left            =   6000
      Picture         =   "FRMBOARD.frx":12C4
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   2
      Left            =   6000
      Picture         =   "FRMBOARD.frx":194E
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   0
      Left            =   4800
      Picture         =   "FRMBOARD.frx":1FD8
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   1
      Left            =   5400
      Picture         =   "FRMBOARD.frx":2662
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   3
      Left            =   4800
      Picture         =   "FRMBOARD.frx":2CEC
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IA 
      Height          =   240
      Index           =   4
      Left            =   5400
      Picture         =   "FRMBOARD.frx":3376
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件名"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   6
      Left            =   10080
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
   Begin VB.Image IU 
      Height          =   705
      Left            =   10800
      MouseIcon       =   "FRMBOARD.frx":3A00
      OLEDropMode     =   1  'Manual
      Picture         =   "FRMBOARD.frx":3D0A
      ToolTipText     =   "退出"
      Top             =   75
      Width           =   750
   End
End
Attribute VB_Name = "FRMBOARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
   ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
   ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
   ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal w As Long, _
   ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal I As Long, ByVal U As Long, _
   ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal cp As Long, ByVal Q As Long, _
   ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Type TRIVERTEX
   X As Long
   Y As Long
   Red0 As Byte
   Red1 As Byte
   Green0 As Byte
   Green1 As Byte
   Blue0 As Byte
   Blue1 As Byte
   Alpha0 As Byte
   Alpha1 As Byte
End Type
Private Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type

Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As Any, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Private Enum GradientFillStyle
 GRADIENT_FILL_RECT_H = 0&
 GRADIENT_FILL_RECT_V = 1&
End Enum

Private Const d_BorderP = &HC56A31
Private Const d_SprtP = &H99A8AC
Private Const d_HlP = &HEDD2C1
Private Const d_CheckedP = &HD7DDFA
Private Const d_PressedP = &HE2B598
Private Const d_Gripper = &H764127
Private Const d_CtrlBorder = &HB99D7F
Private Const d_Icon_Grayscale = &HFFE7D5
Private Const GCL_STYLE As Long = -26
Private Const CS_DROPSHADOW = &H20000
Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CreateActCtxW Lib "kernel32.dll" (ByRef pActCtx As ACTCTXW) As Long
Private Declare Function ActivateActCtx Lib "kernel32.dll" (ByVal hActCtx As Long, ByRef lplpCookie As Long) As Long
Private Type ACTCTXW
 CBSIZE As Long
 dwFlags As Long
 lpcwstrSource As Long
 wProcessorArchitecture As Integer
 wLangId As Integer
 lpcwstrAssemblyDirectory As Long
 lpcwstrResourceName As Long
 lpcwstrApplicationName As Long
 hModule As Long
End Type
Private Const ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID As Long = 1
Private Const ACTCTX_FLAG_LANGID_VALID As Long = 2
Private Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID As Long = 4
Private Const ACTCTX_FLAG_RESOURCE_NAME_VALID As Long = 8
Private Const ACTCTX_FLAG_SET_PROCESS_DEFAULT As Long = 16
Private Const ACTCTX_FLAG_APPLICATION_NAME_VALID As Long = 32
Private Const ACTCTX_FLAG_HMODULE_VALID As Long = 128
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim RGBColor As Long, Red As Long, Green As Long, Blue As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private bm0 As New cDIBSection
Private bm As New cDIBSection
Private bm2 As New cDIBSection
Private m_rOld As Single, m_gOld As Single, m_bOld As Single, m_aOld As Single
Private m_r As Single, m_g As Single, m_b As Single, m_a As Single
Private m_HSB_H As Single, m_HSB_S As Single, m_HSB_B As Single
Private bInteger As Boolean, bClamp As Boolean, bAlpha As Boolean
Private nSelType As Long
Private idxHl As Long
Private d() As Long
Const LOGPIXELSX As Long = 88
Const LOGPIXELSY As Long = 90
Const m_MinFormW = 9120
Const m_MinFormH = 5550
Const m_rulerWH = 17
Dim m_origW As Long
Dim m_origH As Long
Dim m_Magnify As Long
Dim m_Left As Long
Dim m_Top As Long
Dim m_LogH As Long
Dim m_LogV As Long
Dim m_UnitH As Long
Dim m_UnitV As Long
Dim m_MarkIntervalH As Long
Dim m_MarkIntervalV As Long
Dim m_ValueIntervalH As Long
Dim m_ValueIntervalV As Long
Dim m_LongMarkH As Long
Dim m_LongMarkV As Long
Dim X As Long, Y As Long
Dim w As Long, H As Long
Dim XHi, XLo, YHi, YLo, StopDraw
Dim PENSTYLE As Integer, ZOOM_P As Double
Dim MouseDown As Boolean, MYLEFT, MYTOP, C_PEN_SIZE As Long, T_CAN_M As Boolean
Dim PsdFile As New cPSD
Private Type DrawPos
    StartX As Single
    StartY As Single
    OldX As Single
    OldY As Single
    X1 As Single
    y1 As Single
End Type
Private mDrawPos As DrawPos
Dim IS_M As Boolean, OLDH As Long, OLDW As Long
Dim OX As Integer, oy As Integer, mX As Integer, mY As Integer, Color As Long

Private Sub CBFS_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
TXTU.FontSize = CBFS(0).Text
lRet = SetInitEntry("DRAWIT", "FONTSIZE", CBFS(0).ListIndex)
PICTY.FontSize = CBFS(0).Text
Case 1
TXTU.FontName = CBFS(1).Text
lRet = SetInitEntry("DRAWIT", "FONTNAME", CBFS(1).ListIndex)
PICTY.FontName = CBFS(1).Text
End Select
End Sub

Private Sub Form_Activate()
Dim I As Integer
Call UnHook
H_DOS = 3
gHW = Me.hwnd '鼠标控件
Call Hook '唤醒鼠标滑轮API
Me.BackColor = COLOR_NOR
PHELP.BackColor = COLOR_NOR
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
For I = 0 To ICM.Count - 1
ICM(I).M_STYLE = 2
ICM(I).L_M_R = 1
ICM(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next
For I = 0 To ICT.Count - 1
ICT(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next
For I = 0 To IZOOM.Count - 1
IZOOM(I).SETCOLOR vbWhite, COLOR_HIGH, vbBlack
Next
PICOPTION.BackColor = COLOR_NOR
IW.SETCOLOR COLOR_NOR, COLOR_HIGH
End Sub

Private Sub Form_Load()
Dim I As Integer
IU.PICTURE = X1.PICTURE
IS_FULLSCREEN = False
m_Magnify = 1
IS_M = True
MYLEFT = GetInitEntry("DRAWIT", "LEFT", (Screen.Width - Me.Width) / 2)
MYTOP = GetInitEntry("DRAWIT", "TOP", (Screen.Height - Me.Height) / 2)
LA(8).Caption = "R:0G:0B:0 " & "     图像尺寸:" & PICTY.Width & " × " & PICTY.Height

If LONELY_MODE = False Then
If frmma.Left > Me.Width Then Me.Move frmma.Left - Me.Width, frmma.Top Else Me.Move frmma.Left + frmma.Width, frmma.Top
Else
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Load Frmm
End If

LA(6).Caption = "[空白]"
PENSTYLE = 0
StopDraw = 1
SCUT.Move 0, 0, 0, 0
C_PEN_SIZE = 0
ZOOM_P = 1
ICM(0).SETTXT "打开"
ICM(1).SETTXT "保存"
ICM(2).SETTXT "打印"
ICM(3).SETTXT "分享"
ICM(4).SETTXT "全屏模式"
ICM(5).SETTXT "漫画模板"
ICM(6).SETTXT "裁切"

ICT(0).SETTXT "填充"
ICT(1).SETTXT "马克笔"
ICT(2).SETTXT "文本"
ICT(3).SETTXT "铅笔"
ICT(4).SETTXT "马赛克"
ICT(5).SETTXT "喷枪"
ICT(6).SETTXT "直线"
ICT(7).SETTXT "圆形"
ICT(8).SETTXT "裁切"

ICT(1).IS_SELECT = True
IW.HASLINE = False
IW.HASTIP = False
IW.SETTIP "工具栏"
IW.SETTXT "工具栏"

IW.SETTXTCOLOR vbWhite, vbWhite
PICTY.Height = GetInitEntry("DrawIt", "HEIGHT", 500)
PICTY.Width = GetInitEntry("DrawIt", "WIDTH", 500)

TXTX.Text = PICTY.ScaleWidth
TXTY.Text = PICTY.ScaleHeight

PENSTYLE = 1
SCRO.Value = 0
SCRO.LargeChange = 200
SGRO.Value = 0
SGRO.LargeChange = 200
SGRO.Orientation = oHorizontal
Call oMagneticWnd.AddWindow(Me.hwnd)

For I = 10 To 100
CBPEN.AddItem (I)
Next

For I = 0 To IZOOM.Count - 1
IZOOM(I).SETCOLOR vbWhite, &H8BA31F, vbBlack
Next

IZOOM(0).SETTXT "+"
IZOOM(1).SETTXT "原始大小"

For I = 1 To 20
CBSIZE.AddItem (I)
Next

For I = 0 To IOK.Count - 1
IOK(I).HASLINE = False
IOK(I).SETTXT "使用此模板"
Next

For I = 8 To 100
CBFS(0).AddItem I
Next
Call FillComboWithFonts(CBFS(1))

CBFS(0).ListIndex = GetInitEntry("DRAWIT", "FONTSIZE", 5)

CBPEN.ListIndex = GetInitEntry("DRAWIT", "PEN", 90)
CBSIZE.ListIndex = GetInitEntry("DRAWIT", "SIZE", 5)

FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">打开了涂鸦画板"

PF.BackColor = GetInitEntry("DRAWIT", "PENCOLOR", vbBlack)
PB.BackColor = GetInitEntry("SYSTEM", "PICCOLOR", vbWhite)

PICTY.BackColor = PB.BackColor
PT.BackColor = PICTY.BackColor
CBFS(1).ListIndex = GetInitEntry("DRAWIT", "FONTNAME", 0)
Call DRAW_COLOR_PAN
End Sub

Private Sub Form_LostFocus()
lRet = SetInitEntry("MsgBOX", "LEFT", (Screen.Width - FrmWrong.Width) / 2)
lRet = SetInitEntry("MsgBOX", "TOP", (Screen.Height - FrmWrong.Height) / 2)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_FULLSCREEN = False Then Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PO(2).Visible = True Then PO(2).Visible = False
If PICOPTION.Visible = True Then PICOPTION.Visible = False
If IU.PICTURE <> X1.PICTURE Then IU.PICTURE = X1.PICTURE
Call HideRulerMarkers
If IS_M = True Then
IS_M = False
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PICTY_OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Cls
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
PTXT.Line (0, 0)-(PTXT.ScaleWidth - 1, PTXT.ScaleHeight - 1), COLOR_HIGH, B
PICOPTION.Line (0, 0)-(PICOPTION.ScaleWidth - 1, PICOPTION.ScaleHeight - 1), COLOR_HIGH, B
PHELP.Move 1, Me.ScaleHeight - PHELP.Height - 1
IU.Move Me.ScaleWidth - IU.Width - 1, 1
LA(6).Move IU.Left - LA(6).Width - 5
ICM(4).Move IU.Left - ICM(4).Width - 5
SGRO.Move 2, PHELP.Top - SGRO.Height - 1
PO(0).Move 1, 64, Me.ScaleWidth - SCRO.Width - 2, SGRO.Top - PO(0).Top
SGRO.Width = PO(0).Width
SCRO.Move Me.ScaleWidth - SCRO.Width - 2, 64, 15, PO(0).Height
PM.Move 1, 1, Me.ScaleWidth - 2, Me.ScaleHeight - 2
If PICOPTION.Visible = True Then PICOPTION.Visible = False
IW.Move PO(0).ScaleWidth - 63, PO(0).ScaleHeight - 63
End Sub

Private Sub Form_Terminate()
Set PICTY = Nothing
Set PT = Nothing
Set FRMBOARD = Nothing

End Sub
Private Sub ICM_Click(Index As Integer)
Select Case Index
Case 0
Dim filename As String
filename = ShowOpen(Me.hwnd, "*.Bmp;*.JPG;*.GIF;*.PNG;*.PSD" & Chr(0) & "*.Bmp;*.JPG;*.GIF;*.PNG;*.PSD", "打开")
If filename = "" Then Exit Sub
Call OpenFile(filename)
Case 1
Call SAVETY
Case 3
DefCOM = 0
Call SavePicture(PICTY.image, App.Path & "\THUMBS\THUMBS.Bmp")
Call frmma.SHAREIT(App.Path & "\THUMBS\THUMBS.Bmp")
Case 4
If IS_FULLSCREEN = False Then
IS_FULLSCREEN = True
ICM(4).SETTXT "退出全屏"
Me.Move 0, 0, Screen.Width, Screen.Height - GetTaskbarHeight
Else
IS_FULLSCREEN = False
ICM(4).SETTXT "全屏模式"
Me.Move MYLEFT, MYTOP, 11610, 9390
End If
Case 2
On Error GoTo ERRHAND:
PrintPictureToFitPage Printer, PICTY.PICTURE
Printer.EndDoc
ERRHAND:
Call SHOWWRONG("打印机错误:" & ERR.Description, 0)
Case 5
PM.Visible = True
Case 6
PCUT.Move 0, 0, SCUT.Width, SCUT.Height
L = BitBlt(PCUT.hdc, 0, 0, SCUT.Width, SCUT.Height, PICTY.hdc, SCUT.Left, SCUT.Top, &HCC0020)
PCUT.Refresh
Call SavePicture(PCUT.image, App.Path & "\THUMBS\THUMBS_CUT.BMP")
Call OpenFile(App.Path & "\THUMBS\THUMBS_CUT.BMP")
ICM(6).Visible = False
SCUT.Visible = False
End Select
PICOPTION.Visible = False
End Sub



Private Sub ICT_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Dim I As Integer
For I = 0 To ICT.Count - 1
ICT(I).IS_SELECT = False
Next
ICT(Index).IS_SELECT = True
PENSTYLE = Index
PTXT.Visible = False
Select Case Index
Case 2
PTXT.Visible = True
End Select
End Sub

Private Sub IM_TYPE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub IMODE_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub IOK_CLICK(Index As Integer)
PICTY.Cls
TXTX.Text = 600
TXTY.Text = 800
PICTY.Height = 800
PICTY.Width = 600
PICTY.DrawWidth = 5
PICTY.BackColor = vbWhite
Select Case Index
Case 0
PICTY.Line (0, 0)-(300, 400), vbBlack, B
PICTY.Line (300, 0)-(600, 800), vbBlack, B
PICTY.Line (0, 0)-(300, 800), vbBlack, B
PICTY.Line (300, 400)-(600, 800), vbBlack, B
Case 1
PICTY.Line (0, 0)-(600, 400), vbBlack, B
PICTY.Line (0, 400)-(300, 800), vbBlack, B
PICTY.Line (300, 400)-(600, 800), vbBlack, B
Case 2
PICTY.Line (0, 0)-(600, 200), vbBlack, B
PICTY.Line (0, 200)-(600, 400), vbBlack, B
PICTY.Line (0, 400)-(600, 600), vbBlack, B
PICTY.Line (0, 600)-(600, 800), vbBlack, B
End Select
PM.Visible = False
End Sub

Private Sub IU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE = X2.PICTURE Then IU.PICTURE = X3.PICTURE
End Sub
Private Sub IU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE = X1.PICTURE Then IU.PICTURE = X2.PICTURE
End Sub
Private Sub IU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If IU.PICTURE = X3.PICTURE Then IU.PICTURE = X1.PICTURE
If Button = 1 Then Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim I As Integer
lRet = SetInitEntry("DRAWIT", "LEFT", Me.Left)
lRet = SetInitEntry("DRAWIT", "TOP", Me.Top)
lRet = SetInitEntry("DRAWIT", "PENCOLOR", PF.BackColor)
lRet = SetInitEntry("DRAWIT", "PEN", CBPEN.ListIndex)
lRet = SetInitEntry("DRAWIT", "SIZE", CBSIZE.ListIndex)
lRet = SetInitEntry("SYSTEM", "PICCOLOR", PB.BackColor)
lRet = SetInitEntry("DrawIt", "HEIGHT", PICTY.Height)
lRet = SetInitEntry("DrawIt", "WIDTH", PICTY.Width)
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">关闭了涂鸦画板"
IS_FULLSCREEN = False
PsdFile.FreePsd
If LONELY_MODE = True Then End
End Sub

Private Sub IW_Click()
PICOPTION.Move PO(0).ScaleWidth - PICOPTION.Width, IW.Top + IW.Height - PICOPTION.Height
PICOPTION.Visible = Not PICOPTION.Visible
End Sub

Private Sub IZOOM_Click(Index As Integer)
On Error Resume Next
Call SavePicture(PICTY.image, App.Path & "\THUMBS\THUMBS_ZOOM.BMP")

Select Case Index
Case 0
ZOOM_P = ZOOM_P + 0.1
PICTY.Move 0, 0, PT.Width * ZOOM_P, PT.Height * ZOOM_P
Set PICTY.PICTURE = PICTY.image
Call DrawPictureByNum(PICTY.hdc, App.Path & "\THUMBS\THUMBS_ZOOM.BMP", 0, 0, ZOOM_P)
Case 1
Call SavePicture(PICSOU.image, App.Path & "\THUMBS\THUMBS_ZOOM.BMP")
PICTY.Move 0, 0, OLDW, OLDH
ZOOM_P = 1
'Call DrawPictureByNum(PICTY.hdc, App.Path & "\THUMBS\THUMBS_ZOOM.BMP", 0, 0, 1)
PICTY.PaintPicture PICSOU.PICTURE, 0, 0, OLDW, OLDH
Set PICTY.PICTURE = PICTY.image
End Select


Kill App.Path & "\THUMBS\THUMBS_ZOOM.BMP"

End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_FULLSCREEN = False Then Call CMV(Me)
End Sub

Private Sub LA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 4, 0
Call SetHand
End Select
End Sub
Private Sub p2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PF.BackColor = p2.BackColor
End Sub

Private Sub PB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
On Error GoTo ERR
PB.BackColor = frmma.ShowColor(Me)
PT.BackColor = PB.BackColor
ERR:
Exit Sub
End Sub

Private Sub PB_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PICTY_OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub PBK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PM.Visible = False
End Sub

Private Sub PBK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_M = False Then
IS_M = True
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PBK.hdc, 0, 0)
End If
End Sub
Private Sub PF_Change()
On Error Resume Next
PICTY.FOREColor = PF.BackColor
End Sub

Private Sub PF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PO(2).Visible = False Then PO(2).Visible = True
PO(2).ZOrder 0
End Sub


Private Sub PF_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PICTY_OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub PHELP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_FULLSCREEN = False Then Call CMV(Me)
End Sub

Private Sub PHELP_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PICTY_OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub PICTY_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Select Case PENSTYLE
Case 0
Call Filling(PICTY, PF.BackColor, PICTY.POINT(X, Y), 0, X, Y)
Case 1
drawAirbrush Me.PICTY.hdc, CLng(X), CLng(Y), 1, PF.BackColor, CBPEN.Text
drawAirbrush Me.PT.hdc, CLng(X), CLng(Y), 1, PF.BackColor, CBPEN.Text
MouseDown = True
PICTY.PaintPicture PT.PICTURE, 0, 0
Case 2
PTXT.Visible = True
PTXT.Move X, Y
TXTU.SetFocus
Case 3
MouseDown = True
OX = X
oy = Y
Case 4
Dim xl As Integer, yl As Integer
Dim K As Integer, j As Integer
If PICTY.ScaleWidth - 1 - mX >= 30 And PICTY.ScaleHeight - 1 - mY >= 30 Then
'检测当前鼠标的位置,防止处理边缘溢出
For yl = mY To mY + 20 Step 10  '局部马赛克处理
 For xl = mX To mX + 20 Step 10
  Color = GetPixel(PICTY.hdc, xl + 15, yl + 15)
  '取每个马赛克小块的中心象素的颜色为填充整个小块的颜色
  r = (Color Mod 256)
  b = (Int(Color / 65536))
  G = Int((Color - (b * 65536) - r) / 256)
  For K = 0 To 9  '填充整个马赛克小块的颜色
   For j = 0 To 9
     Call SetPixel(PICTY.hdc, xl + K, yl + j, RGB(r, G, b))
   Next j
  Next K
  PICTY.Refresh  '图象刷新
 Next xl
Next yl
End If
Case 5
MouseDown = True
Call 喷笔(PICTY, PF.BackColor, X, Y)
Case 6
MouseDown = True
PICTY.FillStyle = vbFSTransparent
PICTY.AutoRedraw = False
    With mDrawPos
        .StartX = X
        .StartY = Y
        .OldX = .StartX
        .OldX = .StartY
    End With
Case 7

Case 8
StopDraw = 0
XLo = X
YLo = Y
XHi = X
YHi = Y
SCUT.Width = Abs(XHi - XLo)
SCUT.Height = Abs(YHi - YLo)

End Select
End If
If Button = 2 Then Me.PopupMenu Frmm.特效
End Sub

Private Sub PICTY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If PO(2).Visible = True Then PO(2).Visible = False
If PENSTYLE = 8 Then
SCUT.Visible = True
Else
SCUT.Visible = False
ICM(6).Visible = False
End If
    shpMarkerW.Visible = True
    shpMarkerH.Visible = True
    shpMarkerW.Left = X - SGRO.Value       ' Remember our rulers are limited to viewport w/h the most
    shpMarkerH.Top = Y - SCRO.Value

If PICOPTION.Visible = True Then PICOPTION.Visible = False
Dim rc As Long, GC As Long, BC As Long, CC As Long
CC = PICTY.POINT(X, Y)
rc = CC Mod 256
GC = CC \ 256 Mod 256
BC = CC \ 256 \ 256
LA(8).Caption = "R:" & rc & " G:" & GC & " B:" & BC & "   图像尺寸:" & PICTY.Width & " × " & PICTY.Height
Select Case PENSTYLE
Case 1
If MouseDown = False Then Exit Sub
C_PEN_SIZE = C_PEN_SIZE + 1
If C_PEN_SIZE >= CBSIZE.Text Then C_PEN_SIZE = CBSIZE.Text
drawAirbrush PICTY.hdc, CLng(X), CLng(Y), C_PEN_SIZE, PF.BackColor, CBPEN.Text
drawAirbrush PT.hdc, CLng(X), CLng(Y), C_PEN_SIZE, PF.BackColor, CBPEN.Text
PICTY.PaintPicture PT.PICTURE, 0, 0
Case 2

Case 3
If MouseDown = False Then Exit Sub
PICTY.DrawWidth = CBSIZE.Text + 5
PICTY.Line (OX, oy)-(X, Y)
PICTY.FOREColor = PF.BackColor
OX = X
oy = Y
Case 4
mX = X - 20
mY = Y - 20
Case 5
If MouseDown = False Then Exit Sub
Call 喷笔(PICTY, PF.BackColor, X, Y)
Case 6
If MouseDown = False Then Exit Sub
PICTY.Refresh
PICTY.DrawWidth = CBSIZE.Text
PICTY.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
Case 7

Case 8
XHi = X
YHi = Y
If XHi < 0 Then XHi = 0
If YHi < 0 Then YHi = 0
If XHi > PPV.ScaleWidth Then XHi = PPV.ScaleWidth
If YHi > PPV.ScaleHeight Then YHi = PPV.ScaleHeight
If StopDraw <> 0 Then Exit Sub
SCUT.Width = Abs(XHi - XLo)
SCUT.Height = Abs(YHi - YLo)
SCUT.Visible = True
        If XHi > XLo And YHi > YLo Then
            SCUT.Top = YLo
            SCUT.Left = XLo
        End If
        If XHi > XLo And YHi < YLo Then
            SCUT.Top = YHi
            SCUT.Left = XLo
        End If
        If XHi < XLo And YHi < YLo Then
            SCUT.Top = YHi
            SCUT.Left = XHi
        End If
        If XHi < XLo And YHi > YLo Then
            SCUT.Top = YLo
            SCUT.Left = XHi
        End If
        ICM(6).Move SCUT.Left, SCUT.Top
End Select
End Sub

Private Sub PICTY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
MouseDown = False
PICTY.AutoRedraw = True
If PENSTYLE = 6 Then PICTY.Line (mDrawPos.StartX, mDrawPos.StartY)-(X, Y)
If PENSTYLE = 8 Then
StopDraw = 1
ICM(6).Visible = True
End If
C_PEN_SIZE = 0
PICTY.DrawWidth = 1
If PT.AutoRedraw = True Then PT.AUTOSIZE = False
Call PictureBoxSaveJPG(PICTY.image, App.Path & "\MEDIA\Paint\AutoSave.JPG", 100)
If PICTY.Height < 240 Or PICTY.Width < 60 Then
Exit Sub
Else
Call SETPRE
Set PICSOU.PICTURE = PICTY.image
End If
End Sub
Sub OpenFile(filename As String)
On Error Resume Next
If filename = "" Or PathFileExists(filename) = 0 Then Exit Sub
Select Case UCase(Right(filename, 3))
Case "PNG"
Call OPENISPNG(PT, filename)
IM_TYPE.PICTURE = IA(4).PICTURE
Case "PSD"
PICTY.Cls
PsdFile.LoadPsdFile (filename)
PT.Move 0, 0, PsdFile.Width, PsdFile.Height
PsdFile.DrawToDC PT.hdc, 0, 0
IM_TYPE.PICTURE = IA(2).PICTURE
Case "BMP"
PT.PICTURE = LoadPicture(filename)
IM_TYPE.PICTURE = IA(0).PICTURE
Case "JPG"
PT.PICTURE = LoadPicture(filename)
IM_TYPE.PICTURE = IA(3).PICTURE
Case "GIF"
PT.PICTURE = LoadPicture(filename)
IM_TYPE.PICTURE = IA(1).PICTURE
Case Else
IM_TYPE.PICTURE = IA(5).PICTURE
End Select

Set PICTY.PICTURE = Nothing
PT.AUTOSIZE = True
PICTY.Move 0, 0, PT.ScaleWidth, PT.ScaleHeight
PICTY.PaintPicture PT.image, 0, 0, PT.ScaleWidth, PT.ScaleHeight
Set PT.PICTURE = Nothing
PT.AUTOSIZE = False
PB.BackColor = PT.POINT(0, 0)
LA(6).Caption = "[" & ShortName(filename) & "]"
Call MMAIN.PictureBoxSaveJPG(PICTY.image, App.Path & "\MEDIA\Paint\AutoSave.JPG", 100)
SCRO.Value = 0
SGRO.Value = 0
TXTX.Text = PICTY.ScaleWidth
TXTY.Text = PICTY.ScaleHeight
LA(8).Caption = "R:0G:0B:0 " & "     图像尺寸:" & PICTY.Width & " × " & PICTY.Height
ZOOM_P = 1
OLDW = PICTY.Width
OLDH = PICTY.Height
Set PICSOU.PICTURE = PICTY.image
Me.Show
End Sub


Sub SAVETY()
Dim filename As String, SB As String
filename = ShowSave(Me.hwnd, "Bmp" & Chr(0) & "*.Bmp" & Chr(0) & "JEPG" & Chr(0) & "*.JPG", "保存")
If filename = "" Then Exit Sub
SB = UCase(Right(filename, 3))
Select Case SB
Case "BMP"
Call SavePicture(PICTY.image, filename)
Case "JPG"
Call PictureBoxSaveJPG(PICTY.image, filename, 100)
End Select
End Sub

Private Sub PICTY_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strpath As String
If Data.files.Count = 0 Then Exit Sub
strpath = Data.files(1)
Select Case LCase$(Right$(Data.files(1), 3))
Case "bmp", "dib", "gif", "jpg", "png", "psd"
Call OpenFile(strpath)
End Select
End Sub

Private Sub PICTY_Resize()
m_origW = PICTY.ScaleWidth
m_origH = PICTY.ScaleHeight
SCRO.Max = PICTY.Height - PO(0).ScaleHeight + 250
SGRO.Max = PICTY.Width - PO(0).ScaleWidth + 250
PT.AUTOSIZE = False
PT.Move 0, 0, PICTY.Width, PICTY.Height
Call AlignRulers
Call SETPRE
End Sub
Sub SETPRE()
On Error Resume Next
'IPR.PICTURE = PICTY.image
If PICTY.Height > PVP.ScaleHeight Or PICTY.Width > PVP.ScaleWidth Then
IPR.Height = PVP.ScaleHeight
IPR.Width = PVP.ScaleWidth * (IPR.Height / PICTY.ScaleHeight)
Dimention2 IPR, PICTY, PICTY.ScaleWidth * (IPR.Height / PICTY.ScaleHeight), IPR.Height
IPR.Move (PVP.ScaleWidth - IPR.Width) / 2, 0
Else
Dimention2 IPR, PICTY, PICTY.Width, PICTY.Height
IPR.Move (PVP.ScaleWidth - IPR.Width) / 2, (PVP.ScaleHeight - IPR.Height) / 2
End If
End Sub
Private Sub PM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_M = True Then
IS_M = False
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub PM_Resize()
SB.Move 0, (PM.ScaleHeight - SB.Height), PM.ScaleWidth
IMODE(1).Move (PM.ScaleWidth - IMODE(1).Width) / 2
IMODE(4).Move (PM.ScaleWidth - IMODE(4).Width) / 2
IMODE(0).Move IMODE(1).Left - IMODE(0).Width - 10
IMODE(3).Move IMODE(4).Left - IMODE(3).Width - 10
IMODE(2).Move IMODE(1).Left + IMODE(2).Width + 10
IMODE(5).Move IMODE(4).Left + IMODE(5).Width + 10
LA(4).Move 5, PM.ScaleHeight - LA(4).Height - 5, PM.ScaleWidth
End Sub

Private Sub PO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PO(2).Visible = True Then PO(2).Visible = False
If PICOPTION.Visible = True Then PICOPTION.Visible = False
Call HideRulerMarkers
If IS_M = True Then
IS_M = False
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub PO_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PICTY_OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub PO_Resize(Index As Integer)
PO(1).Move picRulerBaseW.Height, picRulerBaseH.Width, PO(0).ScaleWidth - picRulerBaseH.Width, PO(0).ScaleHeight - picRulerBaseW.Height
picRulerBaseH.Top = picRulerBaseW.Height
picRulerBaseW.Left = picRulerBaseH.Width
Call AlignRulers

End Sub
Private Sub SCRO_Change()

PICTY.Top = -SCRO.Value
PlotRulers

End Sub

Private Sub SCRO_Scroll()
SCRO_Change
End Sub

Private Sub SGRO_Change()

PICTY.Left = -SGRO.Value
PlotRulers

End Sub

Private Sub SGRO_Scroll()
SGRO_Change
End Sub

Private Sub TXTU_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
PTXT.Visible = False
PICTY.CurrentX = PTXT.Left
PICTY.CurrentY = PTXT.Top
PICTY.FOREColor = PF.BackColor
PICTY.Print TXTU.Text
End If
End Sub

Private Sub TXTX_Change()
If Trim(TXTX.Text) = "" Then Exit Sub
PICTY.Move 0, 0
If Int(TXTX.Text) > 3600 Then TXTX.Text = 3600
If Int(TXTY.Text) > 3600 Then TXTY.Text = 3600
PICTY.Width = Int(TXTX.Text)
PT.Width = Int(TXTX.Text)
SCRO.Value = 0
SGRO.Value = 0
lRet = SetInitEntry("DrawIt", "WIDTH", PICTY.ScaleWidth)

End Sub

Private Sub TXTX_KeyPress(KeyAscii As Integer)
KeyAscii = VailText(KeyAscii, "0123456789", True)
End Sub

Private Sub TXTY_Change()
If Trim(TXTY) = "" Then Exit Sub
PICTY.Move 0, 0
PICTY.Height = Int(TXTY.Text)
PT.Height = Int(TXTY.Text)
SCRO.Value = 0
SGRO.Value = 0
lRet = SetInitEntry("DrawIt", "HEIGHT", PICTY.ScaleHeight)
End Sub

Private Sub TXTY_KeyPress(KeyAscii As Integer)
KeyAscii = VailText(KeyAscii, "0123456789", True)
End Sub
Private Sub HideRulerMarkers()
    shpMarkerW.Visible = False
    shpMarkerH.Visible = False
End Sub
Private Sub AlignRulers()
    Dim w As Long
    Dim H As Long
    
    w = m_origW
    H = m_origH
    
    If PICTY.PICTURE = 0 Then
        HideRulerMarkers
        
        PICTY.Move 0, 0
        
        picRulerBaseW.Left = PO(1).Left
        picRulerBaseW.Top = PO(1).Top - picRulerBaseW.Height
    
        picRulerBaseH.Left = PO(1).Left - picRulerBaseH.Width
        picRulerBaseH.Top = PO(1).Top
    
        picRulerBaseW.Width = PO(1).Width
        picRulerBaseH.Height = PO(1).Height
    Else
          ' Size ruler
        If PICTY.ScaleWidth > PO(1).ScaleWidth Then
             picRulerBaseW.Width = PO(1).Width
        Else
             picRulerBaseW.Left = picRulerBaseW.Left - 1
             picRulerBaseW.Top = picRulerBaseW.Top - 1
             
             picRulerBaseW.Width = PICTY.Width + 1
        End If
        If PICTY.ScaleHeight > PO(1).ScaleHeight Then
             picRulerBaseH.Height = PO(1).Height
        Else
             picRulerBaseH.Left = picRulerBaseH.Left - 1
             picRulerBaseH.Top = picRulerBaseH.Top - 1
        
             picRulerBaseH.Height = PICTY.Height + 1
        End If
        
          ' Pos ruler
        If PICTY.Left > 0 Then
            picRulerBaseW.Left = PO(1).Left + PICTY.Left
        Else
            picRulerBaseW.Left = PO(1).Left
        End If
        If PICTY.Top > 0 Then
            picRulerBaseW.Top = (PO(1).Top + PICTY.Top) - picRulerBaseW.Height
        Else
            picRulerBaseW.Top = PO(1).Top - picRulerBaseW.Height
        End If
        
        If PICTY.Left > 0 Then
            picRulerBaseH.Left = (PO(1).Left + PICTY.Left) - picRulerBaseH.Width
        Else
            picRulerBaseH.Left = PO(1).Left - picRulerBaseH.Width
        End If
        If PICTY.Top > 0 Then
            picRulerBaseH.Top = PO(1).Top + PICTY.Top
        Else
            picRulerBaseH.Top = PO(1).Top
        End If
    End If
    
    picRulerW.Width = PO(1).Width
    picRulerH.Height = PO(1).Height
    
    PlotRulers
End Sub



Private Sub PlotRulers()
    Dim mIndex As Long
    Dim X As Long, Y As Long
    Dim w As Long, H As Long
    Dim mUnitH2 As Double
    Dim mUnitV2 As Double
    Dim mStart000X As Long
    Dim mStart000Y As Long
    Dim mFractionX As Long
    Dim mFractionY As Long
    Dim mText As String
    Dim mTextW As Single
    Dim cy As Long
    Dim K As Long
    Dim I As Long, j As Long
    m_UnitH = 100
    m_UnitV = 100
    m_MarkIntervalH = 10
    m_MarkIntervalV = 10
    m_LongMarkH = 50
    m_LongMarkV = 50
    w = Fix(PO(1).ScaleWidth / (100 * m_Magnify)) * (100 * m_Magnify)
    H = Fix(PO(1).ScaleHeight / (100 * m_Magnify)) * (100 * m_Magnify)
    w = w + (100 * m_Magnify) * 2
    H = H + (100 * m_Magnify) * 2
    picTemp1.PICTURE = LoadPicture()
    picTemp2.PICTURE = LoadPicture()
    picTemp1.Width = picTemp1.Width - picTemp1.ScaleWidth + w
    picTemp1.Height = picTemp1.Height - picTemp1.ScaleHeight + (m_rulerWH - 2)
    picTemp2.Width = picTemp2.Width - picTemp2.ScaleWidth + (m_rulerWH - 2)
    picTemp2.Height = picTemp2.Height - picTemp2.ScaleHeight + H
        w = PICTY.ScaleWidth
        H = PICTY.ScaleHeight
        
        X = SGRO.Value / m_Magnify
        Y = SCRO.Value / m_Magnify
        X = Abs(X)
        Y = Abs(Y)
             mStart000X = Fix(X / m_UnitH) * m_UnitH
             mStart000Y = Fix(Y / m_UnitV) * m_UnitV
        If X Mod m_UnitH <> 0 Then mFractionX = (X Mod m_UnitH) * m_Magnify
        If Y Mod m_UnitV <> 0 Then mFractionY = (Y Mod m_UnitV) * m_Magnify
    picTemp1.FontName = "Arial"
    picTemp1.FontBold = False
    picTemp1.FontSize = 7
    picTemp1.BackColor = picRulerW.BackColor
    picTemp1.FOREColor = COLOR_NOR
    For I = 0 To picTemp1.ScaleWidth Step (m_MarkIntervalH * m_Magnify)
        If I Mod (m_LongMarkH * m_Magnify) = 0 Then
             picTemp1.Line (I, 0)-(I, 4)
             picTemp1.Line (I, m_rulerWH - 6)-(I, m_rulerWH - 1)
        Else
             picTemp1.Line (I, 0)-(I, 2)
             picTemp1.Line (I, m_rulerWH - 4)-(I, m_rulerWH - 1)
        End If
    Next I
    m_ValueIntervalH = (m_UnitH * m_Magnify)           ' Print "000" mark at intervals
    m_ValueIntervalV = (m_UnitV * m_Magnify)           ' Print "000" mark at intervals
    
    picTemp1.FOREColor = COLOR_NOR
    cy = (picTemp1.ScaleHeight - picTemp1.TextHeight("2")) / 2
    
    K = picTemp1.ScaleWidth / m_ValueIntervalH

    For I = 0 To K
        j = I * m_ValueIntervalH                          ' The print pos
        If mIndex = 0 Then
             mText = CStr(I * m_UnitH + mStart000X)       ' ' NB: Here "i * 100" only
        ElseIf mIndex = 1 Then
             mText = CStr(I + mStart000X)
        ElseIf mIndex = 2 Then
             mText = CStr(I + mStart000X)
        End If
        
        mTextW = picTemp1.TextWidth(mText)
        
        If Len(mText) < 2 Then
            picTemp1.CurrentX = j - 1
        Else
            picTemp1.CurrentX = (j - (mTextW / 2))
        End If
        picTemp1.CurrentY = cy
        picTemp1.Print mText
    Next I
    picTemp1.PICTURE = picTemp1.image
     
     ' Transfer image of picTemp1 to picRulerW, taking into account of mFractionX
    picRulerW.PICTURE = LoadPicture()
     ' "+..." below so that during scrolling, whole width on picRulerW is covered
    w = picRulerW.ScaleWidth + (100 * m_Magnify)
    H = picRulerW.ScaleHeight
    If mFractionX = 0 Then
        BitBlt picRulerW.hdc, 0, 0, w, H, picTemp1.hdc, 0, 0, vbSrcCopy
    Else
        BitBlt picRulerW.hdc, -mFractionX, 0, w, H, picTemp1.hdc, 0, 0, vbSrcCopy
    End If
    picRulerW.PICTURE = picRulerW.image

    picTemp2.FontName = "Arial"
    picTemp2.FontSize = 7
    picTemp2.BackColor = picRulerH.BackColor
    picTemp2.FOREColor = COLOR_NOR
    
    For I = 0 To picTemp2.ScaleHeight Step (m_MarkIntervalV * m_Magnify)
        If I Mod (m_LongMarkV * m_Magnify) = 0 Then
             picTemp2.Line (0, I)-(4, I)
             picTemp2.Line (m_rulerWH - 6, I)-(m_rulerWH - 1, I)
        Else
             picTemp2.Line (0, I)-(2, I)
             picTemp2.Line (m_rulerWH - 4, I)-(m_rulerWH - 1, I)
        End If
    Next I

    picTemp2.FOREColor = COLOR_NOR
    MarkVertical picTemp2, mStart000Y
    picRulerH.PICTURE = LoadPicture()
    w = picRulerH.ScaleWidth
    H = picRulerH.ScaleHeight + (100 * m_Magnify)
    If mFractionY = 0 Then
        BitBlt picRulerH.hdc, 0, 0, w, H, picTemp2.hdc, 0, 0, vbSrcCopy
    Else
        BitBlt picRulerH.hdc, 0, -mFractionY, w, H, picTemp2.hdc, 0, 0, vbSrcCopy
    End If
    picRulerH.PICTURE = picRulerH.image
    DoEvents
End Sub
Private Sub MarkVertical(inPic As PictureBox, ByVal inStart000Val As Long)
    Const FW_NORMAL = 400
    Const FONT_SIZE = 10            ' Approx Pic.FontSize of 7, the smallest suitable
    Const FONT_NAME = "Arial"
    Dim mIndex As Long
    Dim mFont As Long
    Dim mFontOld As Long
    Dim mText As String
    Dim mTextH As Single
    Dim K As Long
    Dim X As Single
    Dim I As Long
'    On Error GoTo ErrHandler
    X = (inPic.ScaleWidth - picRulerW.TextWidth("2")) / 2 - 1
    K = inPic.ScaleHeight / m_ValueIntervalV
    
    For I = 0 To K
       If mIndex = 0 Then
             mText = CStr(I * m_UnitV + inStart000Val)       ' ' NB: Here "i * 100" only
        ElseIf mIndex = 1 Then
             mText = CStr(I + inStart000Val)
        ElseIf mIndex = 2 Then
             mText = CStr(I + inStart000Val)
        End If
        
        mFont = CreateFont(FONT_SIZE, 0, 900, 900, FW_NORMAL, False, False, False, _
                0, 0, 0, 0, 0, FONT_NAME)
                
        mFontOld = SelectObject(inPic.hdc, mFont)
        mTextH = inPic.TextWidth(mText)        ' Refer to TextWidth, not TextHeight here
        
        inPic.CurrentX = X
        inPic.CurrentY = (I * m_ValueIntervalV) + (mTextH / 2)
        inPic.Print mText
        
        SelectObject inPic.hdc, mFontOld
        DeleteObject mFont
    Next I

End Sub
Private Sub getLogPixels(inX As Long, inY As Long)
    Dim mHDC As Long
    mHDC = GetDC(0&)
    inX = GetDeviceCaps(mHDC, LOGPIXELSX)
    inY = GetDeviceCaps(mHDC, LOGPIXELSY)
End Sub

Sub DRAW_COLOR_PAN()
Dim I As Long
bm0.CreateFromPicture i0.PICTURE
bm.Create p1.ScaleWidth, p1.ScaleHeight
t1(4).Text = GetInitEntry("DRAWIT", "R", "255")
t1(5).Text = GetInitEntry("DRAWIT", "G", "255")
t1(6).Text = GetInitEntry("DRAWIT", "B", "255")
pRedraw
pChange -1
opt1(3).Visible = bAlpha
t1(3).Visible = bAlpha
p2.BackColor = PF.BackColor
End Sub
Friend Property Get ClampBorder() As Boolean
ClampBorder = bClamp
End Property

Friend Property Let ClampBorder(ByVal b As Boolean)
bClamp = b
End Property

Friend Property Get UseInteger() As Boolean
UseInteger = bInteger
End Property

Friend Property Let UseInteger(ByVal b As Boolean)
bInteger = b
End Property

Friend Property Get UseAlpha() As Boolean
UseAlpha = bAlpha
End Property

Friend Property Let UseAlpha(ByVal b As Boolean)
bAlpha = b
End Property
Friend Function GetColorData() As Long
On Error Resume Next
Dim rgbRed As Long, rgbGreen As Long, rgbBlue As Long, rgbReserved As Long
rgbRed = m_r
rgbGreen = m_g
rgbBlue = m_b
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
If bAlpha Then
 rgbReserved = m_a
 If rgbReserved < 0 Then rgbReserved = 0 Else If rgbReserved > 255 Then rgbReserved = 255
 If rgbReserved > 127 Then rgbReserved = rgbReserved - 256
End If
GetColorData = rgbRed Or _
(rgbGreen * &H100&) Or _
(rgbBlue * &H10000) Or _
(rgbReserved * &H1000000)
End Function

Friend Sub SetColorData(ByVal clr As Long)
m_r = clr And &HFF&
m_g = (clr And &HFF00&) \ &H100&
m_b = (clr And &HFF0000) \ &H10000
If bAlpha Then m_a = (clr And &HFF000000) \ &H1000000 Else m_a = 255
If m_a < 0 Then m_a = m_a + 256
m_rOld = m_r
m_gOld = m_g
m_bOld = m_b
m_aOld = m_a
pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Sub

Friend Sub GetColor(ByRef rgbRed As Single, ByRef rgbGreen As Single, ByRef rgbBlue As Single, ByRef rgbReserved As Single)
rgbRed = m_r
rgbGreen = m_g
rgbBlue = m_b
If bAlpha Then rgbReserved = m_a
End Sub

Friend Sub SETCOLOR(ByVal rgbRed As Single, ByVal rgbGreen As Single, ByVal rgbBlue As Single, ByVal rgbReserved As Single)
m_r = rgbRed
m_g = rgbGreen
m_b = rgbBlue
If bAlpha Then m_a = rgbReserved Else m_a = 255
m_rOld = rgbRed
m_gOld = rgbGreen
m_bOld = rgbBlue
m_aOld = m_a
pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Sub
Private Sub opt1_Click(Index As Integer)
nSelType = Index
pRedraw
End Sub

Private Sub p1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X < 268 Then idxHl = 1 Else idxHl = 2
If Button = 1 Then p1_MouseMove Button, Shift, X, Y
Me.PF.BackColor = GetColorData
p2.BackColor = GetColorData
lRet = SetInitEntry("system", "pencolor", GetColorData)
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 If idxHl = 1 Then 'left
  pClickLeft X - 4, Y - 4
  pRedraw
  pChange -1
 ElseIf idxHl = 2 Then 'right
  pClickRight Y - 4
  pRedraw
  pChange -1
 End If
End If
End Sub
Private Sub p1_Paint()
bm.PaintPicture p1.hdc
End Sub
Private Sub PSD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then FRMBOARD.PF.BackColor = PSD.BackColor Else Sleep 200: FRMCOLOR.Show
End Sub
Private Sub t1_Change(Index As Integer)
On Error Resume Next
Dim s As String
Dim f As Single, b As Long
Select Case Index
Case 0 'H
 f = Val(t1(Index).Text)
 If f < 0 Then f = 0: b = -1 Else If f > 360 Then f = 360: b = -1
 m_HSB_H = f
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 1 'S
 f = Val(t1(Index).Text)
 If f < 0 Then f = 0: b = -1 Else If f > 100 Then f = 100: b = -1
 m_HSB_S = f
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 2 'B
 f = Val(t1(Index).Text)
 If f < 0 Then f = 0: b = -1 Else If f > 100 Then f = 100: b = -1
 m_HSB_B = f
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 3 'A
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_a = f
Case 4 'R
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_r = f
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 5 'G
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_g = f
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 6 'B
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_b = f
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 7 'web
 s = Replace(t1(Index).Text, "#", "")
 s = Replace(s, " ", "")
 m_r = Val("&H" + Mid(s, 1, 2))
 m_g = Val("&H" + Mid(s, 3, 2))
 m_b = Val("&H" + Mid(s, 5, 2))
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 8 'VB
 b = Val(t1(Index).Text)
 m_r = (b And &HFF&)
 m_g = (b And &HFF00&) \ &H100&
 m_b = (b And &HFF0000) \ &H10000
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
 b = 0
End Select
pChange Index Or b
pRedraw

lRet = SetInitEntry("DRAWIT", "R", t1(4).Text)
lRet = SetInitEntry("DRAWIT", "G", t1(5).Text)
lRet = SetInitEntry("DRAWIT", "B", t1(6).Text)

End Sub

Private Sub pChange(ByVal Index As Long)
On Error Resume Next
Dim s As String, I As Long
I = m_r
If I < 0 Then I = 0 Else If I > 255 Then I = 255
If I < 16 Then s = s + "0" + Hex(I) Else s = s + Hex(I)
I = m_g
If I < 0 Then I = 0 Else If I > 255 Then I = 255
If I < 16 Then s = s + "0" + Hex(I) Else s = s + Hex(I)
I = m_b
If I < 0 Then I = 0 Else If I > 255 Then I = 255
If I < 16 Then s = s + "0" + Hex(I) Else s = s + Hex(I)
If Index <> 0 Then If bInteger Then t1(0).Text = CStr(Round(m_HSB_H)) Else t1(0).Text = CStr(m_HSB_H)
If Index <> 1 Then If bInteger Then t1(1).Text = CStr(Round(m_HSB_S)) Else t1(1).Text = CStr(m_HSB_S)
If Index <> 2 Then If bInteger Then t1(2).Text = CStr(Round(m_HSB_B)) Else t1(2).Text = CStr(m_HSB_B)
If Index <> 3 Then If bInteger Then t1(3).Text = CStr(Round(m_a)) Else t1(3).Text = CStr(m_a)
If Index <> 4 Then If bInteger Then t1(4).Text = CStr(Round(m_r)) Else t1(4).Text = CStr(m_r)
If Index <> 5 Then If bInteger Then t1(5).Text = CStr(Round(m_g)) Else t1(5).Text = CStr(m_g)
If Index <> 6 Then If bInteger Then t1(6).Text = CStr(Round(m_b)) Else t1(6).Text = CStr(m_b)
If Index <> 7 Then t1(7).Text = s
If Index <> 8 Then t1(8).Text = "&H" + Mid(s, 5, 2) + Mid(s, 3, 2) + Mid(s, 1, 2)
End Sub

Private Sub t1_GotFocus(Index As Integer)
On Error Resume Next
With t1(Index)
 .SelStart = 0
 .SelLength = Len(.Text)
End With
End Sub

Private Sub pRedraw()
On Error Resume Next
Dim r As RECT, hBr As Long
'draw back
hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
r.Right = bm.Width
r.Bottom = bm.Height
FillRect bm.hdc, r, hBr
DeleteObject hBr
'draw border
hBr = CreateSolidBrush(d_CtrlBorder)
r.Right = bm2.Width
r.Bottom = bm2.Height
FrameRect bm2.hdc, r, hBr
r.Left = 3
r.Top = 3
r.Right = 261
r.Bottom = 261
FrameRect bm.hdc, r, hBr
r.Left = 279
r.Right = 305
FrameRect bm.hdc, r, hBr
DeleteObject hBr
'draw
pDrawLeft
pDrawRight
pDrawColor
'over
p1_Paint
End Sub
Private Sub pDrawLeft()
On Error Resume Next
Dim I As Long, clr As Long
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
Dim nSelX As Long, nSelY As Long
'draw left
Select Case nSelType
Case 0 'hsbH
 pHSB2RGB m_HSB_H, 100, 100, rgbRed, rgbGreen, rgbBlue
 For I = 0 To 255
  clr = ((rgbRed * I / 255) And &HFF&) Or _
  (((rgbGreen * I / 255) And &HFF&) * &H100&) Or _
  (((rgbBlue * I / 255) And &HFF&) * &H10000)
  GradientFillRect bm.hdc, 4, 259 - I, 260, 260 - I, I * &H10101, clr, GRADIENT_FILL_RECT_H
 Next I
 nSelX = m_HSB_S / 100 * 255
 nSelY = (1 - m_HSB_B / 100) * 255
Case 1 'hsbS
 For I = 0 To 255
  pHSB2RGB I / 255 * 360, m_HSB_S, 100, rgbRed, rgbGreen, rgbBlue
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  GradientFillRect bm.hdc, 4 + I, 4, 5 + I, 260, clr, vbBlack, GRADIENT_FILL_RECT_V
 Next I
 nSelX = m_HSB_H / 360 * 255
 nSelY = (1 - m_HSB_B / 100) * 255
Case 2, 3 'hsbB
 For I = 0 To 255
  pHSB2RGB I / 255 * 360, 100, m_HSB_B, rgbRed, rgbGreen, rgbBlue
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  GradientFillRect bm.hdc, 4 + I, 4, 5 + I, 260, clr, CLng(m_HSB_B / 100 * 255) * &H10101, GRADIENT_FILL_RECT_V
 Next I
 nSelX = m_HSB_H / 360 * 255
 nSelY = (1 - m_HSB_S / 100) * 255
Case 4 'r
 rgbRed = m_r
 If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
 clr = (rgbRed And &HFF&)
 For I = 0 To 255
  GradientFillRect bm.hdc, 4, 259 - I, 260, 260 - I, clr, clr Or &HFF0000, GRADIENT_FILL_RECT_H
  clr = clr + &H100&
 Next I
 nSelX = m_b
 nSelY = 255 - m_g
Case 5 'g
 rgbGreen = m_g
 If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
 clr = (rgbGreen And &HFF&) * &H100&
 For I = 0 To 255
  GradientFillRect bm.hdc, 4, 259 - I, 260, 260 - I, clr, clr Or &HFF0000, GRADIENT_FILL_RECT_H
  clr = clr + &H1&
 Next I
 nSelX = m_b
 nSelY = 255 - m_r
Case 6 'b
 rgbBlue = m_b
 If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
 clr = (rgbBlue And &HFF&) * &H10000
 For I = 0 To 255
  GradientFillRect bm.hdc, 4, 259 - I, 260, 260 - I, clr, clr Or &HFF&, GRADIENT_FILL_RECT_H
  clr = clr + &H100&
 Next I
 nSelX = m_r
 nSelY = 255 - m_g
End Select
'draw selected
clr = CreateRectRgn(4, 4, 260, 260)
SelectClipRgn bm.hdc, clr
bm0.PaintPicture bm.hdc, nSelX - 1, nSelY - 1, 11, 11, 0, 9, vbSrcInvert
SelectClipRgn bm.hdc, 0
End Sub

Private Sub pClickLeft(ByVal nSelX As Long, ByVal nSelY As Long)
If nSelX < 0 Then nSelX = 0 Else If nSelX > 255 Then nSelX = 255
If nSelY < 0 Then nSelY = 0 Else If nSelY > 255 Then nSelY = 255
Select Case nSelType
Case 0 'hsbH
 m_HSB_S = nSelX / 255 * 100
 m_HSB_B = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 1 'hsbS
 m_HSB_H = nSelX / 255 * 360
 m_HSB_B = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 2, 3 'hsbB
 m_HSB_H = nSelX / 255 * 360
 m_HSB_S = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 4 'r
 m_b = nSelX
 m_g = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 5 'g
 m_b = nSelX
 m_r = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 6 'b
 m_r = nSelX
 m_g = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Select
End Sub

Private Sub pClickRight(ByVal nSelY As Long)
If nSelY < 0 Then nSelY = 0 Else If nSelY > 255 Then nSelY = 255
Select Case nSelType
Case 0 'hsbH
 m_HSB_H = (255 - nSelY) / 255 * 360
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 1 'hsbS
 m_HSB_S = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 2 'hsbB
 m_HSB_B = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 3 'alpha!!!
 m_a = nSelY
Case 4 'r
 m_r = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 5 'g
 m_g = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 6 'b
 m_b = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Select
End Sub

Private Sub pDrawRight()
On Error Resume Next
Dim r As RECT
Dim I As Long, clr As Long, clr2 As Long
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
'draw right
Select Case nSelType
Case 0 'hsbH
 r.Left = 280
 r.Right = 304
 For I = 0 To 255
  pHSB2RGB I / 255 * 360, 100, 100, rgbRed, rgbGreen, rgbBlue
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  clr = CreateSolidBrush(clr)
  r.Top = 259 - I
  r.Bottom = r.Top + 1
  FillRect bm.hdc, r, clr
  DeleteObject clr
 Next I
 I = (1 - m_HSB_H / 360) * 255
Case 1 'hsbS
 rgbRed = m_HSB_B
 If rgbRed < 20 Then rgbRed = 20
 I = CLng(rgbRed / 100 * 255) * &H10101
 pHSB2RGB m_HSB_H, 100, rgbRed, rgbRed, rgbGreen, rgbBlue
 clr = (rgbRed And &HFF&) Or _
 ((rgbGreen And &HFF&) * &H100&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr, I, GRADIENT_FILL_RECT_V
 I = (1 - m_HSB_S / 100) * 255
Case 2 'hsbB
 pHSB2RGB m_HSB_H, m_HSB_S, 100, rgbRed, rgbGreen, rgbBlue
 clr = (rgbRed And &HFF&) Or _
 ((rgbGreen And &HFF&) * &H100&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr, vbBlack, GRADIENT_FILL_RECT_V
 I = (1 - m_HSB_B / 100) * 255
Case 3 'alpha!!!
 '////////
 For I = 0 To 255
  r.Top = 259 - I
  r.Bottom = r.Top + 1
  'calc blend
  rgbRed = I + m_r
  rgbGreen = I + m_g
  rgbBlue = I + m_b
  If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
  If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
  If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  clr = CreateSolidBrush(clr)
  rgbBlue = I / 2
  rgbRed = rgbBlue + m_r
  rgbGreen = rgbBlue + m_g
  rgbBlue = rgbBlue + m_b
  If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
  If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
  If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
  clr2 = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  clr2 = CreateSolidBrush(clr2)
  If I And 8& Then
   clr = clr Xor clr2
   clr2 = clr2 Xor clr
   clr = clr Xor clr2
  End If
  r.Left = 280
  r.Right = 288
  FillRect bm.hdc, r, clr2
  r.Left = 288
  r.Right = 296
  FillRect bm.hdc, r, clr
  r.Left = 296
  r.Right = 304
  FillRect bm.hdc, r, clr2
  DeleteObject clr
  DeleteObject clr2
 Next I
 '////////
 I = m_a
Case 4 'r
 rgbGreen = m_g
 rgbBlue = m_b
 If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
 If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
 clr = ((rgbGreen And &HFF&) * &H100&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr + &HFF&, clr, GRADIENT_FILL_RECT_V
 I = 255 - m_r
Case 5 'g
 rgbRed = m_r
 rgbBlue = m_b
 If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
 If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
 clr = (rgbRed And &HFF&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr + &HFF00&, clr, GRADIENT_FILL_RECT_V
 I = 255 - m_g
Case 6 'b
 rgbRed = m_r
 rgbGreen = m_g
 If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
 If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
 clr = (rgbRed And &HFF&) Or _
 ((rgbGreen And &HFF&) * &H100&)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr + &HFF0000, clr, GRADIENT_FILL_RECT_V
 I = 255 - m_b
End Select
'draw selected
TransparentBlt bm.hdc, 269, I, 9, 9, bm0.hdc, 0, 0, 9, 9, vbGreen
TransparentBlt bm.hdc, 306, I, 9, 9, bm0.hdc, 8, 0, 9, 9, vbGreen
End Sub

Private Sub pDrawColor()
On Error Resume Next
Dim r As RECT
Dim I As Long, j As Long, clr As Long
Dim f As Single
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
'calc new color
f = 255 - m_a
rgbRed = f + m_r
rgbGreen = f + m_g
rgbBlue = f + m_b
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j
 r.Bottom = r.Top + 8
 For I = (j And 15&) To 57 Step 16
  r.Left = I
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next I
Next j
DeleteObject clr
f = f / 2
rgbRed = f + m_r
rgbGreen = f + m_g
rgbBlue = f + m_b
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j
 r.Bottom = r.Top + 8
 For I = ((j + 8&) And 15&) To 57 Step 16
  r.Left = I
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next I
Next j
DeleteObject clr
'calc old color
f = 255 - m_aOld
rgbRed = f + m_rOld
rgbGreen = f + m_gOld
rgbBlue = f + m_bOld
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j + 32
 r.Bottom = r.Top + 8
 For I = (j And 15&) To 57 Step 16
  r.Left = I
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next I
Next j
DeleteObject clr
f = f / 2
rgbRed = f + m_rOld
rgbGreen = f + m_gOld
rgbBlue = f + m_bOld
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j + 32
 r.Bottom = r.Top + 8
 For I = ((j + 8&) And 15&) To 57 Step 16
  r.Left = I
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next I
Next j
DeleteObject clr
End Sub

Private Sub pRGB2HSB(ByVal rgbRed As Single, ByVal rgbGreen As Single, ByVal rgbBlue As Single, ByRef hsbH As Single, ByRef hsbS As Single, ByRef hsbB As Single)
Dim fMax As Single, nMax As Long, fMin As Single
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
If rgbRed > rgbGreen Then
 If rgbRed > rgbBlue Then
  fMax = rgbRed
  nMax = 1
  If rgbGreen > rgbBlue Then fMin = rgbBlue Else fMin = rgbGreen
 Else
  fMax = rgbBlue
  nMax = 3
  fMin = rgbGreen
 End If
Else
 If rgbGreen > rgbBlue Then
  fMax = rgbGreen
  nMax = 2
  If rgbRed > rgbBlue Then fMin = rgbBlue Else fMin = rgbRed
 Else
  fMax = rgbBlue
  nMax = 3
  fMin = rgbRed
 End If
End If
hsbB = fMax * 100 / 255
If fMax = fMin Then
 hsbS = 0
Else
 fMin = fMax - fMin
 hsbS = 100 * fMin / fMax
 Select Case nMax
 Case 1
  fMax = (rgbGreen - rgbBlue) * 60 / fMin
 Case 2
  fMax = 120 + (rgbBlue - rgbRed) * 60 / fMin
 Case Else
  fMax = 240 + (rgbRed - rgbGreen) * 60 / fMin
 End Select
 If fMax > 360 Then fMax = fMax - 360 Else If fMax < 0 Then fMax = fMax + 360
 hsbH = fMax
End If
End Sub

Private Sub pHSB2RGB(ByVal hsbH As Single, ByVal hsbS As Single, ByVal hsbB As Single, ByRef rgbRed As Single, ByRef rgbGreen As Single, ByRef rgbBlue As Single)
Dim nHue As Long, fMin As Single
hsbH = hsbH - Int(hsbH / 360) * 360
If hsbS < 0 Then hsbS = 0 Else If hsbS > 100 Then hsbS = 100
If hsbB < 0 Then hsbB = 0 Else If hsbB > 100 Then hsbB = 100
hsbB = hsbB * 255 / 100
If hsbS = 0 Then
 rgbRed = hsbB
 rgbGreen = hsbB
 rgbBlue = hsbB
Else
 hsbH = hsbH / 60
 nHue = Int(hsbH)
 hsbH = hsbH - nHue
 hsbS = hsbS / 100
 fMin = hsbB * (1 - hsbS)
 If nHue And 1& Then
  hsbS = hsbB * (1 - hsbS * hsbH)
  hsbH = hsbB
  hsbB = hsbS
 Else
  hsbH = hsbB * (1 - hsbS * (1 - hsbH))
 End If
 If nHue < 2 Then
  rgbRed = hsbB
  rgbGreen = hsbH
  rgbBlue = fMin
 ElseIf nHue < 4 Then
  rgbGreen = hsbB
  rgbBlue = hsbH
  rgbRed = fMin
 Else
  rgbBlue = hsbB
  rgbRed = hsbH
  rgbGreen = fMin
 End If
End If
End Sub
Private Sub GetRGBColors(ByVal RGBColor As Long, ByRef RedColor As Long, ByRef GreenColor As Long, ByRef BlueColor As Long)
RedColor = RGBColor Mod 256
GreenColor = (RGBColor \ &H100) Mod 256
BlueColor = (RGBColor \ &H10000) Mod 256
End Sub

Private Sub PS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_FULLSCREEN = False Then Call CMV(Me)
End Sub

Private Sub PS_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE <> Me.X1.PICTURE Then IU.PICTURE = Me.X1.PICTURE
End Sub
Sub NewLoadManifest()
On Error GoTo a
Dim t As ACTCTXW, s As String
Dim H As Long, I As Long
Debug.Print 1 \ 0
InitCommonControls
t.CBSIZE = Len(t)
s = Space(1024)
GetSystemDirectory s, 1024
I = InStr(s, vbNullChar)
If I > 0 Then s = Left(s, I - 1)
s = s + "\shell32.dll"
t.lpcwstrSource = StrPtr(s)
t.lpcwstrResourceName = 124
t.dwFlags = ACTCTX_FLAG_RESOURCE_NAME_VALID
H = CreateActCtxW(t)
If H <> -1 And H <> 0 Then ActivateActCtx H, I
a:
End Sub
Sub EnableDropShadow(ByVal hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Sub DisbleDropShadow(ByVal hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) And Not CS_DROPSHADOW
End Sub

Private Function pFindWindowByClassName(ByVal sClassName As String) As Long
Dim hwd As Long, PID As Long
Dim p As Long
PID = GetCurrentProcessId
Do
 hwd = FindWindowEx(0, hwd, sClassName, vbNullString)
 If hwd <> 0 Then
  GetWindowThreadProcessId hwd, p
  If p = PID Then
   pFindWindowByClassName = hwd
   Exit Function
  End If
 End If
Loop Until hwd = 0
End Function

Public Function EnableTooltipDropShadow() As Boolean
Dim hwd As Long
Static b As Boolean
If b Then
 EnableTooltipDropShadow = True
 Exit Function
End If
hwd = pFindWindowByClassName("VBBubble")
If hwd <> 0 Then
 EnableDropShadow hwd
 b = True
End If
hwd = pFindWindowByClassName("VBBubbleRT5")
If hwd <> 0 Then
 EnableDropShadow hwd
 b = True
End If
hwd = pFindWindowByClassName("VBBubbleRT6")
If hwd <> 0 Then
 EnableDropShadow hwd
 b = True
End If
EnableTooltipDropShadow = b
End Function
Sub GrayscaleBitmap(bm As cDIBSection, bmOut As cDIBSection, ByVal clr As Long, ByVal clrTrans As Long)
Dim I As Long, j As Long, ii As Long, m As Long
Dim nClrBlue As Long, nClrGreen As Long, nClrRed As Long
Dim nTransBlue As Long, nTransGreen As Long, nTransRed As Long
Dim K As Long
Dim b() As Byte
If bm.Width <= 0 Or bm.Height <= 0 Then Exit Sub
bmOut.Create bm.Width, bm.Height
m = bm.BytesPerScanLine
ReDim b(m - 1)
nClrBlue = (clr And &HFF0000) \ &H10000
nClrGreen = (clr And &HFF00&) \ &H100&
nClrRed = clr And &HFF&
nTransBlue = (clrTrans And &HFF0000) \ &H10000
nTransGreen = (clrTrans And &HFF00&) \ &H100&
nTransRed = clrTrans And &HFF&
For j = 0 To bm.Height - 1
 CopyMemory b(0), ByVal bm.DIBSectionBitsPtr + j * m, m
 ii = 0
 For I = 0 To bm.Width - 1
  If b(ii) <> nTransBlue Or b(ii + 1) <> nTransGreen Or b(ii + 2) <> nTransRed Then
'   k = b(ii) * 146& + b(ii + 1) * 1454& + b(ii + 2) * 456& + 512&
   K = (CLng(b(ii)) + b(ii + 1) + b(ii + 2)) * 685& + 512&
   b(ii) = (nClrBlue * K) \ 524288
   b(ii + 1) = (nClrGreen * K) \ 524288
   b(ii + 2) = (nClrRed * K) \ 524288
  End If
  ii = ii + 3
 Next I
 CopyMemory ByVal bmOut.DIBSectionBitsPtr + j * m, b(0), m
Next j
End Sub


