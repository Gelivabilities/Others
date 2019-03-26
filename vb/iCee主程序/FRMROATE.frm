VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGraphic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "图像处理"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13560
   Icon            =   "FRMROATE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   904
   Begin VB.PictureBox PICBROWSER 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   7665
      Left            =   45
      ScaleHeight     =   511
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   898
      TabIndex        =   19
      Top             =   825
      Visible         =   0   'False
      Width           =   13470
      Begin VB.PictureBox PICVIEW 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   7605
         Left            =   0
         ScaleHeight     =   507
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   898
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   13470
         Begin VB.PictureBox PO 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   0
            Left            =   0
            ScaleHeight     =   73
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   905
            TabIndex        =   50
            Top             =   0
            Width           =   13575
            Begin VB.PictureBox PBK 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               Height          =   900
               Index           =   0
               Left            =   120
               ScaleHeight     =   900
               ScaleWidth      =   900
               TabIndex        =   69
               Top             =   120
               Width           =   900
            End
            Begin VB.Label LB_FN 
               BackStyle       =   0  'Transparent
               Caption         =   "文件名称"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1320
               TabIndex        =   52
               Top             =   480
               Width           =   8160
            End
            Begin VB.Label LB_CT 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "1/90"
               ForeColor       =   &H00C0C0C0&
               Height          =   180
               Left            =   12960
               TabIndex        =   51
               Top             =   600
               Width           =   360
            End
         End
         Begin VB.PictureBox PO 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   6615
            Index           =   2
            Left            =   0
            ScaleHeight     =   441
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   897
            TabIndex        =   55
            Top             =   1080
            Width           =   13455
            Begin VB.FileListBox filHidden 
               Height          =   5670
               Left            =   120
               Pattern         =   "*.bmp;*.dib;*.rle;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur;*.png"
               TabIndex        =   59
               Top             =   0
               Visible         =   0   'False
               Width           =   3075
            End
            Begin VB.PictureBox PICSEE 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               Height          =   4020
               Left            =   3600
               ScaleHeight     =   268
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   476
               TabIndex        =   56
               Top             =   1440
               Width           =   7140
            End
            Begin ICEE.GIF pic_gif 
               Height          =   1935
               Left            =   5280
               Top             =   2880
               Width           =   2295
               _extentx        =   4048
               _extenty        =   3413
               gif             =   "FRMROATE.frx":0802
            End
         End
      End
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1095
         Index           =   4
         Left            =   0
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   897
         TabIndex        =   110
         Top             =   0
         Width           =   13455
         Begin ICEE.ICEE_KEY ICZ 
            Height          =   495
            Index           =   4
            Left            =   12120
            TabIndex        =   112
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
         End
         Begin VB.Label LBCO 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   240
            TabIndex        =   111
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.PictureBox picFrame 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6585
         Left            =   0
         ScaleHeight     =   439
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   900
         TabIndex        =   20
         Top             =   1080
         Width           =   13500
         Begin ICEE.ucScrollbar vsbSlide 
            Height          =   7095
            Left            =   13080
            TabIndex        =   29
            Top             =   120
            Visible         =   0   'False
            Width           =   255
            _extentx        =   450
            _extenty        =   12515
         End
         Begin VB.PictureBox picSlide 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00231C09&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5655
            Left            =   0
            ScaleHeight     =   377
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   232
            TabIndex        =   21
            Top             =   0
            Width           =   3480
            Begin ICEE.ICEE_WIN8 OPTTHUMB 
               Height          =   1890
               Index           =   0
               Left            =   0
               TabIndex        =   113
               Top             =   0
               Width           =   1890
               _extentx        =   3334
               _extenty        =   3334
            End
         End
      End
   End
   Begin VB.PictureBox PICCODE 
      AutoRedraw      =   -1  'True
      BackColor       =   &H001F1F1F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   7665
      Left            =   45
      ScaleHeight     =   511
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   898
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   13470
      Begin ICEE.ICHECK CHECK1 
         Height          =   375
         Left            =   4080
         TabIndex        =   109
         Top             =   1440
         Width           =   3135
         _extentx        =   5530
         _extenty        =   661
      End
      Begin VB.PictureBox PBK 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001F1F1F&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   900
         Index           =   4
         Left            =   120
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   73
         Top             =   120
         Width           =   900
      End
      Begin VB.PictureBox PCODE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4500
         Left            =   4080
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   15
         Top             =   2100
         Width           =   4500
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00808080&
            Height          =   4500
            Left            =   0
            Top             =   0
            Width           =   4500
         End
      End
      Begin VB.ComboBox cmb1 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Version"
         Top             =   1440
         Width           =   3495
      End
      Begin VB.ComboBox cmb1 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Error correction level"
         Top             =   2880
         Width           =   3495
      End
      Begin VB.ComboBox cmb1 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Mask type"
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox TXTCODE 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2265
         Left            =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   4335
         Width           =   3465
      End
      Begin VB.ComboBox cmb1 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Text encoding"
         Top             =   3600
         Width           =   3495
      End
      Begin ICEE.ICEE_COMMAND LB 
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   6600
         Width           =   3495
         _extentx        =   2778
         _extenty        =   873
      End
      Begin ICEE.ICEE_COMMAND LB 
         Height          =   495
         Index           =   4
         Left            =   4080
         TabIndex        =   14
         Top             =   6600
         Width           =   4500
         _extentx        =   2778
         _extenty        =   873
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生成二维码"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   9
         Left            =   1440
         TabIndex        =   77
         Top             =   360
         Width           =   1575
      End
      Begin VB.Shape SOSO 
         BorderColor     =   &H00808080&
         Height          =   2295
         Index           =   0
         Left            =   240
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内容"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   4080
         Width           =   360
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数据格式"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   3360
         Width           =   720
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "信息量"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   26
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编码安全程度"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   2640
         Width           =   1080
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "信息量"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "300×300 Pix"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   1
         Left            =   3240
         TabIndex        =   16
         Top             =   600
         Width           =   1080
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1740
         Left            =   11640
         Stretch         =   -1  'True
         Top             =   1680
         Visible         =   0   'False
         Width           =   1740
      End
   End
   Begin VB.PictureBox PIC_MAIN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   901
      TabIndex        =   63
      Top             =   7200
      Visible         =   0   'False
      Width           =   13515
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1500
         Index           =   0
         Left            =   1680
         TabIndex        =   64
         Top             =   240
         Width           =   1500
         _extentx        =   3969
         _extenty        =   3969
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1500
         Index           =   1
         Left            =   3240
         TabIndex        =   65
         Top             =   240
         Width           =   1500
         _extentx        =   3969
         _extenty        =   3969
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1500
         Index           =   2
         Left            =   4800
         TabIndex        =   66
         Top             =   240
         Width           =   1500
         _extentx        =   3969
         _extenty        =   3969
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1500
         Index           =   3
         Left            =   6360
         TabIndex        =   67
         Top             =   240
         Width           =   1500
         _extentx        =   3969
         _extenty        =   3969
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1500
         Index           =   4
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1500
         _extentx        =   3969
         _extenty        =   3969
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1500
         Index           =   5
         Left            =   7920
         TabIndex        =   107
         Top             =   240
         Width           =   1500
         _extentx        =   3969
         _extenty        =   3969
      End
   End
   Begin VB.PictureBox PICSUN 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7665
      Left            =   45
      ScaleHeight     =   511
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   898
      TabIndex        =   18
      Top             =   825
      Visible         =   0   'False
      Width           =   13470
      Begin VB.PictureBox PBK 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   900
         Index           =   3
         Left            =   120
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   72
         Top             =   120
         Width           =   900
      End
      Begin VB.PictureBox picCurve 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   9360
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   257
         TabIndex        =   31
         Top             =   2880
         Width           =   3855
      End
      Begin ICEE.ICEE_COMMAND ICT 
         Height          =   495
         Index           =   5
         Left            =   9360
         TabIndex        =   32
         Top             =   6720
         Width           =   3855
         _extentx        =   6800
         _extenty        =   873
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4530
         Left            =   1920
         ScaleHeight     =   302
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   402
         TabIndex        =   30
         Top             =   1200
         Visible         =   0   'False
         Width           =   6030
      End
      Begin VB.PictureBox PO 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   6555
         Index           =   1
         Left            =   0
         ScaleHeight     =   437
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   897
         TabIndex        =   53
         Top             =   1080
         Width           =   13455
         Begin VB.PictureBox picMain 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   4530
            Left            =   3960
            ScaleHeight     =   302
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   361
            TabIndex        =   54
            Top             =   1320
            Width           =   5415
         End
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "调整曝光度"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   0
         Left            =   1440
         TabIndex        =   75
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   75
      Left            =   120
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   890
      TabIndex        =   22
      Top             =   9240
      Visible         =   0   'False
      Width           =   13350
      Begin VB.PictureBox picProgressSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   135
         Left            =   -360
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   201
         TabIndex        =   23
         Top             =   0
         Width           =   3015
      End
   End
   Begin ICEE.ICEE_WIN8 ICM 
      Height          =   495
      Left            =   12000
      TabIndex        =   68
      Top             =   8640
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   12795
      Picture         =   "FRMROATE.frx":081A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   62
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   12795
      Picture         =   "FRMROATE.frx":08FE
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   61
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   12795
      Picture         =   "FRMROATE.frx":09E2
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   60
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox PICTALK 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   7785
      Left            =   45
      ScaleHeight     =   519
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   898
      TabIndex        =   17
      Top             =   825
      Visible         =   0   'False
      Width           =   13470
      Begin ICEE.ICEE_COMMAND ICT 
         Height          =   375
         Index           =   7
         Left            =   12000
         TabIndex        =   46
         Top             =   7200
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
      End
      Begin ICEE.ICEE_COMMAND ICT 
         Height          =   375
         Index           =   6
         Left            =   780
         TabIndex        =   45
         Top             =   7200
         Width           =   675
         _extentx        =   2355
         _extenty        =   661
      End
      Begin ICEE.ICEE_COMMAND ICT 
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   47
         Top             =   7200
         Width           =   675
         _extentx        =   2355
         _extenty        =   661
      End
      Begin VB.PictureBox PBK 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   900
         Index           =   2
         Left            =   120
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   71
         Top             =   120
         Width           =   900
      End
      Begin VB.PictureBox PIC_S 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   9840
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   49
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox PIC_T 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   11040
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   1080
      End
      Begin MSComctlLib.ProgressBar PG1 
         Height          =   135
         Left            =   1560
         TabIndex        =   43
         Top             =   7200
         Visible         =   0   'False
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   2400
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   3600
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   4800
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   6000
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   5
         Left            =   12000
         TabIndex        =   38
         Top             =   1200
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   6
         Left            =   12000
         TabIndex        =   39
         Top             =   2400
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   7
         Left            =   12000
         TabIndex        =   40
         Top             =   3600
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   8
         Left            =   12000
         TabIndex        =   41
         Top             =   4800
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin ICEE.ICEE_PIC PIC_DEMO 
         Height          =   1215
         Index           =   9
         Left            =   12000
         TabIndex        =   42
         Top             =   6000
         Width           =   1335
         _extentx        =   2355
         _extenty        =   2143
      End
      Begin VB.PictureBox PO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6030
         Index           =   3
         Left            =   1560
         ScaleHeight     =   402
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   689
         TabIndex        =   57
         Top             =   1200
         Width           =   10335
         Begin VB.PictureBox PDEMO 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   3495
            Left            =   3240
            ScaleHeight     =   233
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   241
            TabIndex        =   58
            Top             =   1200
            Width           =   3615
            Begin VB.Timer TIMAUTO 
               Enabled         =   0   'False
               Interval        =   500
               Left            =   3120
               Top             =   120
            End
         End
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "图像特效"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   8
         Left            =   1440
         TabIndex        =   76
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.PictureBox picBottom 
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   7665
      Left            =   45
      ScaleHeight     =   511
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   898
      TabIndex        =   0
      Top             =   825
      Visible         =   0   'False
      Width           =   13470
      Begin VB.PictureBox PBK 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   900
         Index           =   1
         Left            =   120
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   70
         Top             =   120
         Width           =   900
      End
      Begin ICEE.ICEE_COMMAND LB 
         Height          =   495
         Index           =   0
         Left            =   8640
         TabIndex        =   10
         Top             =   360
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
      End
      Begin ICEE.ICEE_COMMAND LB 
         Height          =   495
         Index           =   1
         Left            =   11760
         TabIndex        =   11
         Top             =   360
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
      End
      Begin ICEE.ICEE_COMMAND LB 
         Height          =   495
         Index           =   2
         Left            =   10200
         TabIndex        =   12
         Top             =   360
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
      End
      Begin VB.PictureBox picClip 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6600
         Left            =   0
         ScaleHeight     =   440
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   897
         TabIndex        =   1
         Top             =   1080
         Width           =   13455
         Begin VB.PictureBox picData 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   3255
            Left            =   4320
            ScaleHeight     =   217
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   281
            TabIndex        =   2
            Top             =   1320
            Width           =   4215
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   60
         Left            =   8640
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   106
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "旋转图像"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Index           =   7
         Left            =   1440
         TabIndex        =   78
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.PictureBox PTXTPIC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   45
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   7695
      ScaleWidth      =   13335
      TabIndex        =   79
      Top             =   840
      Visible         =   0   'False
      Width           =   13335
      Begin VB.PictureBox PBK 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   900
         Index           =   5
         Left            =   120
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   114
         Top             =   120
         Width           =   900
      End
      Begin ICEE.ICHECK CHECK2 
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   3480
         Width           =   1455
         _extentx        =   2566
         _extenty        =   873
      End
      Begin ICEE.ICEE_KEY ICZ 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   101
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1920
         TabIndex        =   92
         Text            =   "M"
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2520
         TabIndex        =   91
         Text            =   "A"
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3120
         TabIndex        =   90
         Text            =   "#"
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   3720
         TabIndex        =   89
         Text            =   "9"
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   4320
         TabIndex        =   88
         Text            =   "l"
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5040
         TabIndex        =   87
         Text            =   "o"
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   5880
         TabIndex        =   86
         Text            =   ":"
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   6480
         TabIndex        =   85
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   83
         Top             =   7680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8880
         TabIndex        =   81
         Text            =   "150"
         Top             =   120
         Width           =   735
      End
      Begin ICEE.ICEE_KEY ICZ 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   102
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
      End
      Begin ICEE.ICEE_KEY ICZ 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   103
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
      End
      Begin ICEE.ICEE_KEY ICZ 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   104
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
      End
      Begin ICEE.ICHECK CHECK3 
         Height          =   375
         Left            =   120
         TabIndex        =   106
         Top             =   3960
         Width           =   1455
         _extentx        =   2566
         _extenty        =   873
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   1800
         ScaleHeight     =   4695
         ScaleWidth      =   4575
         TabIndex        =   80
         Top             =   480
         Width           =   4575
      End
      Begin VB.PictureBox PicP 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4695
         Left            =   1680
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   4695
         ScaleWidth      =   4935
         TabIndex        =   82
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "字符画"
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
         Index           =   10
         Left            =   120
         TabIndex        =   108
         Top             =   7080
         Width           =   855
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "黑"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   1560
         TabIndex        =   100
         Top             =   120
         Width           =   180
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "红"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   2280
         TabIndex        =   99
         Top             =   120
         Width           =   180
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "绿"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   2880
         TabIndex        =   98
         Top             =   120
         Width           =   180
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "黄"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   3480
         TabIndex        =   97
         Top             =   120
         Width           =   180
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "蓝"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   14
         Left            =   4080
         TabIndex        =   96
         Top             =   120
         Width           =   180
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "紫红"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   13
         Left            =   4680
         TabIndex        =   95
         Top             =   120
         Width           =   360
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "青蓝"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   5400
         TabIndex        =   94
         Top             =   120
         Width           =   360
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "白"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   6240
         TabIndex        =   93
         Top             =   120
         Width           =   180
      End
      Begin VB.Label LA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "精度(最大150/最小20)"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   6960
         TabIndex        =   84
         Top             =   120
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X  As Long
    Y  As Long
End Type
Dim MyWH As Long
Dim IS_MV_ON As Boolean
Private I As Long, j As Long
Private cx As Long, cy As Long
Dim SPATH As String, AUTO_T As Long, FILE_T As String
Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const maxNPoints As Byte = 32
Dim nPoints As Byte
Private iX() As Single
Private iY() As Single
Private p() As Single
Private U() As Single
Dim isMouseDown As Boolean  'Track mouse status between MouseDown and MouseMove events
Dim selPoint As Long        'Currently selected knot in the spline
Private results(-1 To 256) As Integer   'Stores the y-values for each x-value in the final spline
Dim MinX As Integer, MaxX As Integer    'Used for calculating leading and trailing values
Private Const mouseAccuracy As Byte = 6 'How close to a knot the user must click to select that knot
Public clsDIB As New CLSPICDIBS
Private clsRotateDIB As CLSPICDIBS
Public sngAngle As Single
Public lngBackColor As Long
Private WithEvents clsProcess As cDIBProcess
Attribute clsProcess.VB_VarHelpID = -1
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByRef lpUsedDefaultChar As Any) As Long
Private Const CP_UTF8 As Long = 65001
Private obj As New clsQRCode
Public Select_Pic As String
Private Function drawCubicSpline()
    
    'Tanner's inserted code: draw the background grid
    picCurve.Cls
    Dim I As Long
    picCurve.FOREColor = RGB(128, 128, 128)
    For I = 0 To 255 Step 64
        picCurve.Line (I, 0)-(I, 255)
        picCurve.Line (0, I)-(255, I)
    Next I
    'Now draw the knots
    picCurve.FOREColor = RGB(255, 0, 0)
    For I = 1 To nPoints
        'If this is the currently selected knot, fill it in with red
        If I = selPoint Then
            picCurve.FillStyle = 0
            picCurve.FillColor = RGB(255, 0, 0)
        End If
        picCurve.Circle (iX(I), iY(I)), 4, RGB(255, 0, 0)
        If I = selPoint Then
            picCurve.FillStyle = 1
            picCurve.FillColor = RGB(0, 0, 0)
        End If
    Next I
    picCurve.FOREColor = RGB(0, 0, 0)
    'Clear the results array and reset the max/min variables
    For I = -1 To 256
        results(I) = -1
    Next I
    MinX = 256
    MaxX = -1
    
    'Now run a loop through the knots, calculating spline values as we go
    Call SetPandU
    Dim XPos As Long, YPos As Single
    For I = 1 To nPoints - 1
        For XPos = iX(I) To iX(I + 1)
            YPos = getCurvePoint(I, XPos)
            If XPos < MinX Then MinX = XPos
            If XPos > MaxX Then MaxX = XPos
            If YPos > 255 Then YPos = 254       'Force values to be in the 1-254 range (0-255 also
            If YPos < 0 Then YPos = 1           ' works, but is harder to see on the picture box)
            results(XPos) = YPos
        Next XPos
    Next I
    
    'Based on the maximum and minimum, calculate preceding and trailing y-values
    For I = -1 To MinX - 1
        results(I) = results(MinX)
    Next I
    For I = 256 To MaxX + 1 Step -1
        results(I) = results(MaxX)
    Next I
    
    'Draw the finished spline
    For I = 0 To 255
        picCurve.Line (I, results(I))-(I - 1, results(I - 1))
    Next I
    picCurve.Refresh
    
    'Last, but certainly not least, draw the curves-adjusted image
    drawCurves picBack, picMain
    
End Function

'Original required spline function:
Private Function getCurvePoint(ByVal I As Long, ByVal V As Single) As Single
    Dim t As Single
    'derived curve equation (which uses p's and u's for coefficients)
    t = (V - iX(I)) / U(I)
    getCurvePoint = t * iY(I + 1) + (1 - t) * iY(I) + U(I) * U(I) * (f(t) * p(I + 1) + f(1 - t) * p(I)) / 6#
End Function

'Original required spline function:
Private Function f(X As Single) As Single
        f = X * X * X - X
End Function

'Original required spline function:
Private Sub SetPandU()
    Dim I As Integer
    Dim d() As Single
    Dim w() As Single
    ReDim d(nPoints) As Single
    ReDim w(nPoints) As Single
    For I = 2 To nPoints - 1
        d(I) = 2 * (iX(I + 1) - iX(I - 1))
    Next
    For I = 1 To nPoints - 1
        U(I) = iX(I + 1) - iX(I)
    Next
    For I = 2 To nPoints - 1
        w(I) = 6# * ((iY(I + 1) - iY(I)) / U(I) - (iY(I) - iY(I - 1)) / U(I - 1))
    Next
    For I = 2 To nPoints - 2
        w(I + 1) = w(I + 1) - w(I) * U(I) / d(I)
        d(I + 1) = d(I + 1) - U(I) * U(I) / d(I)
    Next
    p(1) = 0#
    For I = nPoints - 1 To 2 Step -1
        p(I) = (w(I) - U(I) * p(I + 1)) / d(I)
    Next
    p(nPoints) = 0#
End Sub

Public Sub ResetScroll(ByVal Width As Long, ByVal Height As Long)
    On Error GoTo ErrorHandle
    picData.Visible = False
    picData.Top = 0
    picData.Left = 0
    picData.Width = Width
    picData.Height = Height
    With picClip
         If Width < .ScaleWidth Then picData.Left = (.ScaleWidth - Width) \ 2
         If Height < .ScaleHeight Then picData.Top = (.ScaleHeight - Height) \ 2
    End With

Next1:

ErrorHandle:
    picData.Visible = True
End Sub

Private Sub clsProcess_Complete(ByVal lTimeMs As Long)
    ProgressBar1.Value = 0
    ProgressBar1.Visible = False
End Sub

Private Sub clsProcess_InitProgress(ByVal Max As Long)
    With ProgressBar1
         .Min = 0
         .Max = Max
         .Value = 0
    End With
End Sub

Private Sub clsProcess_Progress(ByVal lPosition As Long)
ProgressBar1.Visible = True
ProgressBar1.Value = lPosition
End Sub
Private Sub FILHIDDEN_PathChange()
    SPATH = filHidden.Path
    If Len(SPATH) > 0 Then SPATH = SPATH & IIf(Right$(SPATH, 1) <> "\", "\", "")
    Call Browse
End Sub

Private Sub Form_Activate()
Call UnHook
H_DOS = 0
gHW = Me.hwnd '鼠标控件
Call Hook '唤醒鼠标滑轮API

End Sub

Private Sub Form_Load()
If LONELY_MODE = False Then
If frmma.Left >= Me.Width / 2 Then Me.Move frmma.Left - Me.Width, frmma.Top Else Me.Move frmma.Left + frmma.Width, frmma.Top
Else
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Load Frmm
End If

On Error Resume Next
CHECK1.Value = GetInitEntry("QCODE", "REALTIME", 0)
Set clsProcess = New cDIBProcess
IS_MV_ON = True

LB(0).SETTXT "打    开"
LB(1).SETTXT "保    存"
LB(2).SETTXT "旋    转"
LB(3).SETTXT "生成二维码"
LB(4).SETTXT "保    存"
Pic_Browse

ICT(5).SETTXT "保    存"
ICT(6).SETTXT "恢复"
ICT(7).SETTXT "保    存"
ICT(8).SETTXT "打开"

ICM.HASTIP = False
ICM.HASLINE = False
ICM.SETTXT "工具"
ICM.SETCOLOR vbBlack, &H554E4

filHidden.Path = App.Path
FILE_T = App.Path & "\THUMBS\THUMB.Bmp"

Dim I As Long
cmb1(0).AddItem "Automatic"
For I = 1 To 40
 cmb1(0).AddItem CStr(I)
Next I

For I = 0 To IW.Count - 1
IW(I).HASLINE = False
IW(I).SETCOLOR vbBlack, &H899F1E
IW(I).SETTXTCOLOR vbWhite, vbWhite
IW(I).SETTXT ""
Next
CHECK3.SETCOLOR PTXTPIC.BackColor, vbWhite
CHECK2.SETCOLOR PTXTPIC.BackColor, vbWhite
CHECK1.SETCOLOR PICCODE.BackColor, vbWhite

CHECK2.SETTXT "白色底色"
CHECK3.SETTXT "彩色字符"
CHECK1.SETTXT "即时预览二维码"

    picFrame.Visible = True
    picSlide.Visible = False
    
CHECK2.Value = 1
CHECK3.Value = 1

CHECK2.M_STYLE = 2
CHECK3.M_STYLE = 2
CHECK1.M_STYLE = 2

IW(0).SETTIP "旋转图片"
IW(1).SETTIP "制作二维码"
IW(2).SETTIP "图像滤镜"
IW(3).SETTIP "曝光度调节"
IW(4).SETTIP "本地图库"
IW(5).SETTIP "字符画"
IW(5).SETTXT "字符画"
IW(5).SETTXTCOLOR vbWhite, vbWhite
IW(5).HASTIP = False

IW(5).SETFONT "微软雅黑", 18, True, 18, True
For I = 0 To ICZ.Count - 1
ICZ(I).SETCOLOR PTXTPIC.BackColor, &HDAA52D, vbWhite
Next
ICZ(0).SETTXT "打开"
ICZ(1).SETTXT "执行"
ICZ(2).SETTXT "复制"
ICZ(3).SETTXT "保存"
ICZ(4).SETTXT "浏览"
ICZ(4).SETCOLOR &H404040, &HB9C127, vbWhite

IW(1).SETPNG App.Path & "\SKIN\QC.PNG", 20, 20
IW(0).SETPNG App.Path & "\SKIN\XUANZ.PNG", 20, 20
IW(2).SETPNG App.Path & "\SKIN\LVJ.PNG", 20, 20
IW(3).SETPNG App.Path & "\SKIN\BAOG.PNG", 20, 20
IW(4).SETPNG App.Path & "\SKIN\VIEW.PNG", 20, 20

cmb1(0).ListIndex = 0
cmb1(1).AddItem "L - 7%"
cmb1(1).AddItem "M - 15%"
cmb1(1).AddItem "Q - 25%"
cmb1(1).AddItem "H - 30%"
cmb1(1).ListIndex = 1
cmb1(2).AddItem "Automatic"

For I = 0 To 7
 cmb1(2).AddItem CStr(I)
Next I

For I = 0 To PIC_DEMO.Count - 1
PIC_DEMO(I).AnyWhere = 1
PIC_DEMO(I).SETIMG PIC_T
Next

PICVIEW.Move 0, 0
PO(2).Move 0, PO(0).Height, PICVIEW.ScaleWidth, PICVIEW.ScaleHeight - PO(0).Height
cmb1(2).ListIndex = 0
cmb1(3).AddItem "ANSI"
cmb1(3).AddItem "UTF-8"
cmb1(3).ListIndex = 1

Call CreatQCode(TXTCODE.Text)

PIC_DEMO(0).SETTXT "增加亮度"
PIC_DEMO(1).SETTXT "减少亮度"
PIC_DEMO(2).SETTXT "模糊"
PIC_DEMO(3).SETTXT "锐化"
PIC_DEMO(4).SETTXT "做旧"
PIC_DEMO(5).SETTXT "暖色调"
PIC_DEMO(6).SETTXT "冷色调"
PIC_DEMO(7).SETTXT "黑白画面"
PIC_DEMO(8).SETTXT "漫画风格"
PIC_DEMO(9).SETTXT "漫反射"

    isMouseDown = False
    selPoint = -1
    MinX = 256
    MaxX = -1
    nPoints = 3
    ReDim iX(nPoints) As Single
    ReDim iY(nPoints) As Single
    ReDim p(nPoints) As Single
    ReDim U(nPoints) As Single
    For I = 1 To nPoints
        iX(I) = (I - 1) * (256 / (nPoints - 1))
        iY(I) = 255 - iX(I)
    Next I
    drawCubicSpline

    cx = PDEMO.ScaleWidth
    cy = PDEMO.ScaleHeight
    Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
    'PIC_MAIN.Line (0, 0)-(PIC_MAIN.ScaleWidth - 1, PIC_MAIN.ScaleHeight - 1),COLOR_HIGH
    Call PaintPng(App.Path & "\SKIN\P_T.PNG", Me.hdc, 8, 8)

    Call SeekMe(Me)
    Me.Show
    filHidden.Path = GetInitEntry("PIC_EDIT", "PATH", App.Path & "\SKIN\BK")
    If filHidden.ListCount = 0 Then
    picFrame.Cls
    Call PaintPng(App.Path & "\SKIN\NO_PIC.PNG", picFrame.hdc, (picFrame.ScaleWidth - 250) / 2, (picFrame.ScaleHeight - 75) / 2)
    Else
    picFrame.Cls
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub Form_Resize()
    Call ResetScroll(picData.Width, picData.Height)
End Sub

Private Sub Form_Terminate()
Set frmGraphic = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim filename As String, lIdx As Long
lRet = SetInitEntry("Qcode", "RealTime", CHECK1.Value)
filename = App.Path & "\THUMBS\THUMBS.Bmp"
fso.DeleteFile filename
For lIdx = 1 To OPTTHUMB.Count - 1
Unload OPTTHUMB(lIdx)
Next lIdx
Set clsDIB = Nothing
Set clsProcess = Nothing
If LONELY_MODE = True Then End
End Sub
Sub 打开(filename As String)
On Error Resume Next
Dim fn As String
If filename = "" Or Dir(filename) = "" Then Exit Sub
    Set clsDIB = New CLSPICDIBS
    Select Case UCase(Right(filename, 3))
    Case "BMP", "JPG", "GIF"
    picData.PICTURE = LoadPicture(filename)
    End Select
    fn = App.Path & "\THUMBS\THUMBS.Bmp"
    Call SavePicture(picData.image, fn)
    picBack.PICTURE = LoadPicture(fn)
    With clsDIB
         .CreateFromFile filename
         Call ResetScroll(.Width, .Height)
         .PaintPicture picData.hdc
         picData.Refresh
         SPATH = GetPathFromFileName(filename, "\")
         filHidden.Path = SPATH
    End With
        Dim fDraw As New FastDrawing
        Dim ImageData() As Byte
        Dim iWidth As Long, iHeight As Long
        iWidth = fDraw.GetImageWidth(picBack)
        iHeight = fDraw.GetImageHeight(picBack)
        
        picMain.Move (PICSUN.ScaleWidth - iWidth) / 2, (PICSUN.ScaleHeight - iHeight) / 2, iWidth, iHeight
        fDraw.GetImageData2D picBack, ImageData()
        fDraw.SetImageData2D picMain, iWidth, iHeight, ImageData()
        If iWidth >= 600 Or iHeight >= 600 Then Exit Sub
        'Call DEMO_IT(fn)
End Sub
Sub DEMO_IT(filename As String)
On Error Resume Next
        PDEMO.PICTURE = LoadPicture(filename)
        PDEMO.Tag = filename
        PDEMO.Move (PO(3).ScaleWidth - PDEMO.Width) / 2, (PO(3).ScaleHeight - PDEMO.Height) / 2
        cx = PDEMO.ScaleWidth
        cy = PDEMO.ScaleHeight
        PIC_S.Move 200, 200, cx, cy
        PIC_S.PaintPicture PDEMO.PICTURE, 0, 0, cx, cy
        PIC_T.Move 0, 0, cy / 3, cy / 3
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        Call 提高亮度(PIC_T)
        PIC_DEMO(0).SETIMG PIC_T
        AUTO_T = 0
        TIMAUTO.Enabled = True
End Sub
Sub View_It(filename As String)
On Error Resume Next
PICVIEW.Visible = True
If Dir(filename) = "" Then Exit Sub
LB_FN.Caption = filename
Call Pic_Browse
PICSEE.PICTURE = LoadPicture(filename)
Select Case LCase(Right(filename, 3))
Case "bmp", "jpg"
pic_gif.LoadAnimatedGIF_File ""
PICSEE.Visible = True
pic_gif.Visible = False
LB_CT.Caption = "1/1"
Case "gif"
PICSEE.Visible = False
pic_gif.Visible = True
pic_gif.LoadAnimatedGIF_File (filename)
Case "png"
pic_gif.LoadAnimatedGIF_File ""
PICSEE.Visible = True
pic_gif.Visible = False
LB_CT.Caption = "1/1"
Call OPENISPNG(PICSEE, filename)
End Select
PICSEE.Move (PICVIEW.ScaleWidth - PICSEE.Width) / 2, (PICVIEW.ScaleHeight - PICSEE.Height) / 2
pic_gif.Move PICSEE.Left, PICSEE.Top, PICSEE.Width, PICSEE.Height
End Sub
Public Sub PrepareImg(pic As PictureBox)
On Error Resume Next
    ReDim larrCol(2, cx, cy)
    For I = 0 To cx
        For j = 0 To cy
            tmpCol = GetPixel(pic.hdc, I, j)
            r = tmpCol Mod 256
            G = (tmpCol / 256) Mod 256
            b = tmpCol / 256 / 256
            larrCol(0, I, j) = r
            larrCol(1, I, j) = G
            larrCol(2, I, j) = b
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    PG1.Value = 0
End Sub
Private Sub 旋转()
    Load frmRotate
    frmRotate.Show vbModal
    If clsRotateDIB Is Nothing Then
       Set clsRotateDIB = New CLSPICDIBS
    Else
       clsRotateDIB.ClearUp
    End If
    With clsDIB
         clsRotateDIB.Create .Width, .Height
         .PaintPicture clsRotateDIB.hdc
    End With
    clsProcess.RotateDIB clsRotateDIB, sngAngle, lngBackColor
    With clsRotateDIB
         Call ResetScroll(.Width, .Height)
         .PaintPicture picData.hdc
    End With
    picData.Refresh
End Sub

Private Sub 保存()
Dim sFile As String
    On Error GoTo ErrorHandle
    sFile = ShowSave(Me.hwnd, "*.BMP位图文件" & Chr$(0) & "*.Bmp", "保存文件")
    
    If clsRotateDIB Is Nothing Then
       clsDIB.SaveBitmap sFile
    Else
       clsRotateDIB.SaveBitmap sFile
    End If
ErrorHandle:
End Sub

Private Sub ICM_Click()
PIC_MAIN.Visible = Not PIC_MAIN.Visible
PIC_MAIN.ZOrder 0
End Sub

Private Sub ICT_CLICK(Index As Integer)
On Error Resume Next
Dim filename As String
Select Case Index
Case 5
Call frmma.保存一下(picMain)
Case 6
PDEMO.PICTURE = LoadPicture(PDEMO.Tag)
Case 7
Call frmma.保存一下(PDEMO)
Case 8
filename = ShowOpen(Me.hwnd, "图像文件(*.Bmp *.gif *.jpg )" & Chr(0) & "*.Bmp;*.gif;*.jpg", "打开图像")
If filename = "" Then Exit Sub
Call DEMO_IT(filename)
End Select

End Sub

Private Sub ICZ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Dim filename As String, SB As String
Select Case Index
Case 0
filename = ShowOpen(Me.hwnd, "*.Bmp;*.JPG;*.GIF" & Chr(0) & "*.Bmp;*.JPG;*.GIF", "打开")
If filename <> "" And MMAIN.PathFileExists(filename) <> 0 Then Call MyOpen(filename)
Case 1
Call 字符画
Case 2
COPY_ZF
Case 3
filename = ShowSave(Me.hwnd, "Bmp" & Chr(0) & "*.Bmp" & Chr(0) & "JEPG" & Chr(0) & "*.JPG", "保存")
If filename = "" Then Exit Sub
SB = UCase(Right(filename, 3))
Select Case SB
Case "BMP"
Call SavePicture(PicP.image, filename)
Case "JPG"
Call PictureBoxSaveJPG(PicP.image, filename, 100)
End Select
Case 4
Call 浏览
End Select
End Sub

Private Sub IW_Click(Index As Integer)
Select Case Index
Case 0
Me.pic_turn
Case 1
Me.pic_code
Case 2
Me.Pic_Talking
Case 3
Me.pic_sun
Case 4
Call Pic_Browse
Case 5
Call Pic_TXT
End Select
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub



Private Sub LB_Click(Index As Integer)
Dim filename As String
Select Case Index
Case 0
filename = ShowOpen(Me.hwnd, "图像文件(*.Bmp *.gif *.jpg)" & Chr(0) & "*.Bmp;*.gif;*.jpg", "打开图像")
If filename = "" Or PathFileExists(filename) = 0 Then Exit Sub
Call 打开(filename)
Case 1
Call 保存
Case 2
Call 旋转
Case 3
Call CreatQCode(TXTCODE.Text)
Case 4
Call frmma.保存一下(PCODE)
End Select
End Sub
Sub CreatQCode(Text As String)
'test
Dim b2() As Byte
Dim I As Long, m As Long
'///
For I = 0 To cmb1.UBound
 If cmb1(I).ListIndex < 0 Then Exit Sub
Next I
'///
Select Case cmb1(3).ListIndex
Case 1
 m = Len(Text)
 I = m * 3 + 64
 ReDim b2(I)
 m = WideCharToMultiByte(CP_UTF8, 0, ByVal StrPtr(Text), m, b2(0), I, ByVal 0, ByVal 0)
Case Else
 Text = StrConv(Text, vbFromUnicode)
 b2 = Text
 m = LenB(Text)
End Select
Set Image1.PICTURE = obj.Encode(b2, m, cmb1(0).ListIndex, cmb1(1).ListIndex + 1, cmb1(2).ListIndex - 1)
PCODE.PaintPicture Image1.PICTURE, 0, 0, PCODE.ScaleWidth, PCODE.ScaleHeight

Call MMAIN.PictureBoxSaveJPG(PCODE.image, App.Path & "\MEDIA\Paint\QCODE.JPG")
End Sub

Private Sub LB_CT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LB_FN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LBCO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub optThumb_Click(Index As Integer)
SPATH = filHidden.Path
If Len(SPATH) > 0 Then SPATH = SPATH & IIf(Right$(SPATH, 1) <> "\", "\", "")
Select_Pic = SPATH & OPTTHUMB(Index).Tag
End Sub

Private Sub optThumb_DBLCLICK(Index As Integer)
Select_Pic = SPATH & OPTTHUMB(Index).Tag
Call View_It(Select_Pic)
End Sub

Private Sub optThumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select_Pic = SPATH & OPTTHUMB(Index).Tag
If Button = 2 Then Me.PopupMenu Frmm.图像处理
End Sub


Private Sub optThumb_MOUSEMOVE(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub



Private Sub PBK_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV_ON = False Then
IS_MV_ON = True
PBK(Index).Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PBK(Index).hdc, 0, 0)
End If
End Sub

Private Sub PBK_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 0
PICVIEW.Visible = False
Case 1
Call Pic_Browse
Case 2
Call Pic_Browse
Case 3
Call Pic_Browse
Case 4
Call Pic_Browse
Case 5
Call Pic_Browse
End Select
End Sub

Private Sub PDEMO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage PDEMO.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub PDEMO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PIC_DEMO_Click(Index As Integer)
Select Case Index
Case 0
Call 提高亮度(PDEMO)
Case 1
Call 变暗(PDEMO)
Case 2
Call 模糊(PDEMO)
Case 3
Call 锐化(PDEMO)
Case 4
Call 做旧(PDEMO)
Case 5
Call 暖色调(PDEMO)
Case 6
Call 冷光(PDEMO)
Case 7
Call 黑白(PDEMO)
Case 8
Call 漫画风格(PDEMO)
Case 9
Call 漫反射(PDEMO)
End Select
End Sub


Private Sub pic_gif_FrameChanged(ByVal FrameIndex As Long, viaTimer As Boolean)
LB_CT.Caption = FrameIndex
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage PICSEE.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub picBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub picBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICBROWSER_Resize()
Dim X       As Long
Dim Y       As Long
Dim lIdx    As Long
Dim lCols   As Long
        
            vsbSlide.Move picFrame.ScaleWidth - vsbSlide.Width, 0, vsbSlide.Width, picFrame.ScaleHeight
            lCols = Int((picFrame.ScaleWidth - vsbSlide.Width) / OPTTHUMB(0).Width)
            For lIdx = 0 To OPTTHUMB.Count - 1
                X = (lIdx Mod lCols) * OPTTHUMB(0).Width
                Y = Int(lIdx / lCols) * OPTTHUMB(0).Height
                OPTTHUMB(lIdx).Move X, Y
            Next lIdx
            picSlide.Width = lCols * OPTTHUMB(0).Width
            picSlide.Height = Int(OPTTHUMB.Count / lCols) * OPTTHUMB(0).Height
            If Int(OPTTHUMB.Count / lCols) < (OPTTHUMB.Count / lCols) Then
                picSlide.Height = picSlide.Height + OPTTHUMB(0).Height
            End If
            vsbSlide.Value = 0
            vsbSlide.Max = picSlide.Height - picFrame.ScaleHeight
            If vsbSlide.Max < 0 Then
                vsbSlide.Max = 0
                vsbSlide.Enabled = False
            Else
                vsbSlide.Enabled = True
                vsbSlide.SmallChange = OPTTHUMB(0).Height
                vsbSlide.LargeChange = picFrame.ScaleHeight
            End If
    
End Sub

Private Sub picClip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub


Private Sub picClip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICCODE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PICCODE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub picData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage picData.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub picData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub picFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub picFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage picMain.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PicP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICSEE_DblClick()
PICVIEW.Visible = False
End Sub

Private Sub PICSEE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage PICSEE.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub picSlide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub picSlide_Resize()
'If picSlide.Height > picFrame.ScaleHeight Then vsbSlide.Visible = True Else vsbSlide.Visible = False
End Sub


Private Sub PICSUN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PICSUN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICTALK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PICTALK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PICVIEW_DblClick()
PICVIEW.Visible = False
End Sub

Private Sub PICVIEW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PO_DblClick(Index As Integer)
Select Case Index
Case 2
PICVIEW.Visible = False
End Select
End Sub

Private Sub PO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub PTXTPIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PTXTPIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MOVENOW
End Sub

Private Sub PTXTPIC_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strpath As String
If Data.files.Count = 0 Then Exit Sub
strpath = Data.files(1)
Select Case LCase$(Right$(Data.files(1), 3))
Case "bmp", "gif", "jpg"
Call MyOpen(strpath)
End Select
End Sub

Private Sub TIMAUTO_Timer()
On Error Resume Next
AUTO_T = AUTO_T + 1
If AUTO_T = 1 Then

ElseIf AUTO_T = 2 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
        Call 变暗(PIC_T)
PIC_DEMO(1).SETIMG PIC_T
ElseIf AUTO_T = 3 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 模糊(PIC_T)
PIC_DEMO(2).SETIMG PIC_T
ElseIf AUTO_T = 4 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 锐化(PIC_T)
PIC_DEMO(3).SETIMG PIC_T
ElseIf AUTO_T = 5 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 做旧(PIC_T)
PIC_DEMO(4).SETIMG PIC_T
ElseIf AUTO_T = 6 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 暖色调(PIC_T)
PIC_DEMO(5).SETIMG PIC_T
ElseIf AUTO_T = 7 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 冷光(PIC_T)
PIC_DEMO(6).SETIMG PIC_T
ElseIf AUTO_T = 8 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 黑白(PIC_T)
PIC_DEMO(7).SETIMG PIC_T
ElseIf AUTO_T = 9 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 漫画风格(PIC_T)
PIC_DEMO(8).SETIMG PIC_T
ElseIf AUTO_T = 10 Then
        PIC_T.PaintPicture PIC_S.image, 0, 0, cx / 3, cy / 3, 0, 0, cx, cy
        Call SavePicture(PIC_T.image, FILE_T)
        PIC_T.PICTURE = LoadPicture(FILE_T)
        
Call 漫反射(PIC_T)
PIC_DEMO(9).SETIMG PIC_T
End If
If AUTO_T > 10 Then TIMAUTO.Enabled = False
End Sub

Private Sub TXTCODE_Change()
If CHECK1.Value = 1 Then Call CreatQCode(TXTCODE.Text)
End Sub

Private Sub TXTCODE_GotFocus()
On Error Resume Next
TXTCODE.SelStart = 0
TXTCODE.SelLength = Len(TXTCODE.Text)
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
If X3.Visible = False Then Me.Hide
End Sub
Sub Pic_TXT()
PICBROWSER.Visible = False
PICCODE.Visible = False
PICSUN.Visible = False
PICTALK.Visible = False
picBottom.Visible = False
PTXTPIC.Visible = True
End Sub
Sub Pic_Talking()
PICBROWSER.Visible = False
PICCODE.Visible = False
PICSUN.Visible = False
PICTALK.Visible = True
picBottom.Visible = False
PTXTPIC.Visible = False
End Sub
Sub Pic_Browse()
PICBROWSER.Visible = True
PICCODE.Visible = False
PICSUN.Visible = False
PICTALK.Visible = False
picBottom.Visible = False
PTXTPIC.Visible = False
End Sub
Sub pic_sun()
PICBROWSER.Visible = False
PICCODE.Visible = False
PICSUN.Visible = True
PICTALK.Visible = False
picBottom.Visible = False
PTXTPIC.Visible = False
End Sub
Sub pic_turn()
PICBROWSER.Visible = False
PICCODE.Visible = False
PICSUN.Visible = False
PICTALK.Visible = False
picBottom.Visible = True
PTXTPIC.Visible = False
End Sub
Sub pic_code()
PICBROWSER.Visible = False
PICCODE.Visible = True
PICSUN.Visible = False
PICTALK.Visible = False
picBottom.Visible = False
PTXTPIC.Visible = False
End Sub

Private Sub vsbSlide_Change()
    picSlide.Top = -vsbSlide.Value
End Sub


Private Sub vsbSlide_Scroll()

    vsbSlide_Change

End Sub

Private Sub Browse(Optional ByVal bDontShowBrowser As Boolean)
Dim lRet    As Long
    If Len(SPATH) > 0 Then
        On Error Resume Next
        If filHidden.Path = SPATH Then Exit Sub
        filHidden.Path = SPATH
        Call CreateThumbs
    End If
End Sub
Private Sub CreateThumbs()
Dim iMaxLen As Integer
Dim lIdx    As Long
Dim lPicCnt As Long
Dim lFilCnt As Long
Dim SPATH   As String
Dim sText   As String
    filHidden.Refresh
    picSlide.Move 0, 0, OPTTHUMB(0).Width, OPTTHUMB(0).Height
    picSlide.Visible = False
    While OPTTHUMB.Count > 1
        Unload OPTTHUMB(OPTTHUMB.Count - 1)
    Wend
    DoEvents
    SPATH = filHidden.Path
    SPATH = SPATH & IIf(Right$(SPATH, 1) <> "\", "\", "")
    lFilCnt = filHidden.ListCount
    
    If lFilCnt = 0 Then
    picFrame.Cls
    Call PaintPng(App.Path & "\SKIN\NO_PIC.PNG", picFrame.hdc, (picFrame.ScaleWidth - 250) / 2, (picFrame.ScaleHeight - 75) / 2)
    Else
    picFrame.Cls
    End If
    
    If Len(SPATH) = 0 Then Exit Sub
        Call StartProgress
        For lIdx = 0 To filHidden.ListCount - 1
            Call UpdateProgress((CSng(lIdx + 1) / CSng(lFilCnt)) * 100, filHidden.List(lIdx))
            ERR.Clear
            If ERR.Number = 0 Then
                If lPicCnt > 0 Then
                    Load OPTTHUMB(lPicCnt)
                    Set OPTTHUMB(lPicCnt).Container = picSlide
                End If
                OPTTHUMB(lPicCnt).Tag = filHidden.List(lIdx)
                OPTTHUMB(lPicCnt).AUTOSIZE = False
                OPTTHUMB(lPicCnt).SETCOLOR COLOR_NOR, COLOR_HIGH
                OPTTHUMB(lPicCnt).SETTXTCOLOR vbWhite, vbWhite
                OPTTHUMB(lPicCnt).IS_PIC = True
                OPTTHUMB(lPicCnt).SETPIC filHidden.Path & "\" & OPTTHUMB(lPicCnt).Tag
                'OPTTHUMB(lPicCnt).SHOWTOOL = True
                sText = filHidden.List(lIdx)
                iMaxLen = OPTTHUMB(lPicCnt).Width - 15
                If picSlide.TextWidth(sText) > iMaxLen Then iMaxLen = iMaxLen - picSlide.TextWidth("...")
                While picSlide.TextWidth(sText) > iMaxLen
                    sText = Left$(sText, Len(sText) - 1)
                Wend
                If iMaxLen < OPTTHUMB(lPicCnt).Width - 15 Then sText = sText & "..."
                OPTTHUMB(lPicCnt).SETTIP sText
                OPTTHUMB(lPicCnt).MYTIT = sText
                'OPTTHUMB(lPicCnt).HAS_TXT = False
                OPTTHUMB(lPicCnt).HASLINE = False
                OPTTHUMB(lPicCnt).Visible = True
                lPicCnt = lPicCnt + 1
            End If
        Next lIdx
        LBCO.Caption = Format(filHidden.ListCount, "000")
        picProgress.Visible = False
        Call PICBROWSER_Resize
        picSlide.Visible = True
End Sub

Private Sub StartProgress()

    picProgress.Cls
    With picProgressSlide
        .Cls
        .Move 0, 0, 1, picProgress.ScaleHeight
    End With
    
    picProgress.Visible = True
    
End Sub

Private Sub UpdateProgress(ByVal iPercent As Integer, ByVal sCaption As String)

Dim lTextTop    As Long

    picProgress.Cls
    picProgressSlide.Cls
    picProgressSlide.Width = picProgress.ScaleWidth * (CSng(iPercent) / 100!)
    lTextTop = (picProgress.ScaleHeight - picProgress.TextHeight(sCaption)) / 2
    picProgress.CurrentX = 3
    picProgress.CurrentY = lTextTop
    'picProgress.Print sCaption
    picProgressSlide.CurrentX = 3
    picProgressSlide.CurrentY = lTextTop
    'picProgressSlide.Print sCaption
    DoEvents
    
End Sub
Private Sub picCurve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'No point selected yet
    selPoint = -1
    
    'Search to see if the user has clicked on (or very near) an existing point
    Dim Found As Long
    Found = checkClick(X, Y)
    
    'If the user has selected an existing point, mark it
    If Found > -1 Then
        selPoint = Found
    Else
        'No match was found, so create a new point here if:
        '  1) This x-coordinate isn't already occupied
        Dim I As Long
        For I = 1 To nPoints
            'The user has clicked on an already occupied x-coordinate. Our spline formula doesn't
            'allow two knots to have the same x-value, so instead of creating a new knot just
            'select the knot already occupying this coordinate
            If X = iX(I) Then
                selPoint = I
                picCurve.MousePointer = 5
                isMouseDown = True
                Exit Sub
            End If
        Next I
        
        '  2) We haven't reached the maximum allowed limit yet
        If nPoints < maxNPoints Then
            'Increase the total number of points and fix all our arrays
            nPoints = nPoints + 1
            ReDim Preserve iX(nPoints) As Single
            ReDim Preserve iY(nPoints) As Single
            ReDim Preserve p(nPoints) As Single
            ReDim Preserve U(nPoints) As Single
            'Figure out which points surround this location
            Dim nextX As Long
            nextX = nPoints
            For I = 1 To nPoints
                If iX(I) > X Then
                    nextX = I
                    Exit For
                End If
            Next I
                        
            'Shift all points after this to the right
            For I = nPoints - 1 To nextX Step -1
                iX(I + 1) = iX(I)
                iY(I + 1) = iY(I)
            Next I
            iX(nextX) = X
            iY(nextX) = Y
            
            'Draw the new spline, change the mousepointer to the move pointer, select this point
            drawCubicSpline
            picCurve.MousePointer = 5
            selPoint = nextX
            
        End If
    End If
    
    'We mark the mouse state here for use in the MouseMove sub
    isMouseDown = True
End Sub

'Simple distance formula here - we use this to calculate if the user has clicked on (or near) a knot
Private Function pDistance(ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Single
    pDistance = Sqr((X1 - X2) ^ 2 + (y1 - y2) ^ 2)
End Function

'MouseMove allows the user to interactively adjust existing knots and add new knots
Private Sub picCurve_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MOVENOW
    'Button down AND a point is current selected
    If isMouseDown = True And selPoint > -1 Then
        'The first knot has to be checked specially (no point before it)
        If selPoint = 0 Then
            If (X >= 0) And (X < iX(selPoint + 1)) Then iX(selPoint) = X
            If (Y >= 0) And (Y <= 255) Then iY(selPoint) = Y
            drawCubicSpline
            Exit Sub
        End If
        'The last knot also has to be checked specially (no point after it)
        If selPoint = nPoints Then
            If (X > iX(selPoint - 1)) And (X <= 255) Then iX(selPoint) = X
            If (Y >= 0) And (Y <= 255) Then iY(selPoint) = Y
            drawCubicSpline
            Exit Sub
        End If
        'For all middle knots, don't allow them to be moved past neighboring knots
        If (X > iX(selPoint - 1)) And (X < iX(selPoint + 1)) Then iX(selPoint) = X
        If (Y >= 0) And (Y <= 255) Then iY(selPoint) = Y
    End If
    drawCubicSpline
    
    'Button up
    If isMouseDown = False Then
        'If the user is close to a knot, change the mousepointer to 'move'
        Dim Found As Long
        Found = checkClick(X, Y)
        If Found > -1 Then
            picCurve.MousePointer = 5
        Else
            picCurve.MousePointer = 0
        End If
    End If
    
End Sub

'When the mouse is lifted, reset the mousestate boolean and the mouse pointer
Private Sub picCurve_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMouseDown = False
    picCurve.MousePointer = 0
End Sub

'Simple distance routine to see if a location on the picture box is near an existing knot
Private Function checkClick(ByVal X As Long, ByVal Y As Long) As Long
    Dim Dist As Single
    Dim I As Long
    For I = 1 To nPoints
        Dist = pDistance(X, Y, iX(I), iY(I))
        'If we're close to an existing point, return the index of that point
        If Dist < mouseAccuracy Then
            checkClick = I
            Exit Function
        End If
    Next I
    'Returning -1 says we're not close to an existing point (so try to create a new one!)
    checkClick = -1
End Function

Public Sub drawCurves(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim ImageData() As Byte
    
    'Coordinate variables
    Dim X As Long, Y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(picBack)
    iHeight = fDraw.GetImageHeight(picBack)
    fDraw.GetImageData2D picBack, ImageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, G As Long, b As Long, L As Long

    'Look-up table calculation for new gamma values
    Dim newGamma(0 To 255) As Byte
    Dim tmpGamma As Double
    For X = 0 To 255
        tmpGamma = CDbl(X) / 255
        If results(X) <= (256 - X) Then
            tmpGamma = tmpGamma ^ (1 / ((256 - X) / (results(X) + 1)))
        Else
            tmpGamma = tmpGamma ^ ((1 / ((256 - X) / (results(X) + 1))) ^ 1.5)
        End If
        tmpGamma = tmpGamma * 255
        If tmpGamma > 255 Then
            tmpGamma = 255
        ElseIf tmpGamma < 0 Then
            tmpGamma = 0
        End If
        newGamma(X) = tmpGamma
    Next X
    
    'Now run a quick loop through the image, adjusting pixel values with the look-up tables
    Dim QuickX As Long
    For X = 0 To iWidth - 1
        QuickX = X * 3
    For Y = 0 To iHeight - 1
        'Grab red, green, and blue
        r = ImageData(QuickX + 2, Y)
        G = ImageData(QuickX + 1, Y)
        b = ImageData(QuickX, Y)
        'Correct them all
        ImageData(QuickX + 2, Y) = newGamma(r)
        ImageData(QuickX + 1, Y) = newGamma(G)
        ImageData(QuickX, Y) = newGamma(b)
    Next Y
    Next X
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D picMain, iWidth, iHeight, ImageData()
End Sub

Sub 浮雕(pic As PictureBox)
    On Error Resume Next
    Call PrepareImg(pic)
    PG1.Visible = True
    For I = 0 To cx - 1
        For j = 0 To cy - 1
            r = Abs(larrCol(0, I, j) - larrCol(0, I + 1, j + 1) + 128)
            G = Abs(larrCol(1, I, j) - larrCol(1, I + 1, j + 1) + 128)
            b = Abs(larrCol(2, I, j) - larrCol(2, I + 1, j + 1) + 128)
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 锐化(pic As PictureBox)
    On Error Resume Next
    Call PrepareImg(pic)
    PG1.Visible = True
    For I = 1 To cx
        For j = 1 To cy
            r = larrCol(0, I, j) + 0.5 * (larrCol(0, I, j) - larrCol(0, I - 1, j - 1))
            G = larrCol(1, I, j) + 0.5 * (larrCol(1, I, j) - larrCol(1, I - 1, j - 1))
            b = larrCol(2, I, j) + 0.5 * (larrCol(2, I, j) - larrCol(2, I - 1, j - 1))
            
            If r > 255 Then r = 255
            If r < 0 Then r = 0
            If G > 255 Then G = 255
            If G < 0 Then G = 0
            If b > 255 Then b = 255
            If b < 0 Then b = 0

            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 漫反射(pic As PictureBox)
    Dim nP1 As Integer, nP2 As Integer, nP3 As Integer
    On Error Resume Next
    Call PrepareImg(pic)
    PG1.Visible = True
    For I = 2 To cx - 3
        For j = 2 To cy - 3
            nP1 = Int(Rnd * 5 - 2)
            nP2 = Int(Rnd * 5 - 2)
            nP3 = Int(Rnd * 5 - 2)
            r = Abs(larrCol(0, I, j + nP1))
            G = Abs(larrCol(1, I + nP2, j))
            b = Abs(larrCol(2, I + nP3, j + nP3))
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 提高亮度(pic As PictureBox)
    Dim C As Long
    PG1.Visible = True
    On Error Resume Next
    Call PrepareImg(pic)
    For I = 0 To cx
    DoEvents
        For j = 0 To cy
        DoEvents
            C = Abs((larrCol(0, I, j) + larrCol(1, I, j) + larrCol(2, I, j)) \ 3)
            r = Abs(larrCol(0, I, j) + C)
            G = Abs(larrCol(1, I, j) + C)
            b = Abs(larrCol(2, I, j) + C)
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 冷光(pic As PictureBox)
    Dim TColI As Long
    On Error Resume Next
    PG1.Visible = True
    For I = 0 To cx
        For j = 0 To cy
            TColI = GetPixel(pic.hdc, I, j)
            r = TColI Mod 256
            G = (TColI \ 256) Mod 256
            b = TColI \ 256 \ 256
            r = Abs((r - G - b) * 1.5)
            G = Abs((G - b - r) * 1.5)
            b = Abs((b - r - G) * 1.5)
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    PG1.Value = 0
    PG1.Visible = False
    pic.Refresh
End Sub
Sub 变暗(pic As PictureBox)
    On Error Resume Next
    PG1.Visible = True
    Call PrepareImg(pic)
    For I = 0 To cx
        For j = 0 To cy
            r = Abs(larrCol(0, I, j) - 64)
            G = Abs(larrCol(1, I, j) - 64)
            b = Abs(larrCol(2, I, j) - 64)
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 暖色调(pic As PictureBox)
    Dim bNo As Boolean
    Dim TColW As Long
    PG1.Visible = True
    On Error Resume Next
    For I = 0 To cx
        For j = 0 To cy
        DoEvents
            TColW = GetPixel(pic.hdc, I, j)
            r = TColW Mod 256
            G = (TColW \ 256) Mod 256
            b = TColW \ 256 \ 256
            
            r = Abs(((r ^ 2) / ((b + G) + 10)) * 128)
            b = Abs(((b ^ 2) / ((G + r) + 10)) * 128)
            G = Abs(((G ^ 2) / ((r + b) + 10)) * 128)
nOK:
                If r > 32767 Then
                    r = r - 32767
                ElseIf G > 32767 Then
                    G = G - 32767
                ElseIf b > 32767 Then
                    b = b - 32767
                End If
                If r > 32767 Or G > 32767 Or b > 32767 Then
                    GoTo nOK
                End If
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        DoEvents
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    PG1.Value = 0
    PG1.Visible = False
    pic.Refresh
End Sub
Sub 陌生的(pic As PictureBox)
    On Error Resume Next
    PG1.Visible = True
    Call PrepareImg(pic)
    For I = 0 To cx
        For j = 0 To cy
            If (larrCol(1, I, j) = 0) Or (larrCol(2, I, j) = 0) Then
                larrCol(1, I, j) = 1
                larrCol(2, I, j) = 1
            End If
            r = Abs(SIN(Atn(larrCol(1, I, j) / larrCol(2, I, j))) * 125 + 20)
            G = Abs(SIN(Atn(larrCol(0, I, j) / larrCol(2, I, j))) * 125 + 20)
            b = Abs(SIN(Atn(larrCol(0, I, j) / larrCol(1, I, j))) * 125 + 20)
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 水绿色(pic As PictureBox)
    Dim tColQ As Long
    
    On Error Resume Next
    PG1.Visible = True
    For I = 0 To cx
        For j = 0 To cy
            tColQ = GetPixel(pic.hdc, I, j)
            r = tColQ Mod 256
            G = (tColQ \ 256) Mod 256
            b = tColQ \ 256 \ 256
            r = (G - b) ^ 2 / 125
            G = (r - b) ^ 2 / 125
            b = (r - G) ^ 2 / 125
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 夜晚(pic As PictureBox)
    On Error Resume Next
    PG1.Visible = True
    Call PrepareImg(pic)
    For I = 0 To cx
        For j = 0 To cy
            r = Abs((larrCol(0, I, j) * larrCol(0, I, j)) / 256)
            G = Abs((larrCol(1, I, j) * larrCol(1, I, j)) / 256)
            b = Abs((larrCol(2, I, j) * larrCol(2, I, j)) / 256)
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 做旧(pic As PictureBox)
    Dim TColA
    
    On Error Resume Next
    PG1.Visible = True
    For I = 0 To cx
        For j = 0 To cy
            TColA = GetPixel(pic.hdc, I, j)
            r = TColA Mod 256
            G = (TColA \ 256) Mod 256
            b = TColA \ 256 \ 256
            r = Abs((G * b) / 256)
            G = Abs((b * r) / 256)
            b = Abs((r * G) / 256)
            SetPixel pic.hdc, I, j, RGB(r, G, b)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub

Sub 模糊(pic As PictureBox)
    On Error Resume Next
    Call PrepareImg(pic)
    PG1.Visible = True
    For I = 1 To cx - 1
        For j = 1 To cy - 1
            r = Abs(larrCol(0, I - 1, j - 1) + larrCol(0, I, j - 1) + larrCol(0, I + 1, j - 1) + larrCol(0, I - 1, j) + larrCol(0, I, j) + larrCol(0, I + 1, j) + larrCol(0, I - 1, j + 1) + larrCol(0, I, j + 1) + larrCol(0, I + 1, j + 1))
            G = Abs(larrCol(1, I - 1, j - 1) + larrCol(1, I, j - 1) + larrCol(1, I + 1, j - 1) + larrCol(1, I - 1, j) + larrCol(1, I, j) + larrCol(1, I + 1, j) + larrCol(1, I - 1, j + 1) + larrCol(1, I, j + 1) + larrCol(1, I + 1, j + 1))
            b = Abs(larrCol(2, I - 1, j - 1) + larrCol(2, I, j - 1) + larrCol(2, I + 1, j - 1) + larrCol(2, I - 1, j) + larrCol(2, I, j) + larrCol(2, I + 1, j) + larrCol(2, I - 1, j + 1) + larrCol(2, I, j + 1) + larrCol(2, I + 1, j + 1))
            SetPixel pic.hdc, I, j, RGB(r / 10, G / 10, b / 10)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 灰化(pic As PictureBox)
    Dim C As Integer
    On Error Resume Next
    PG1.Visible = True
    Call PrepareImg(pic)
    For I = 0 To cx
        For j = 0 To cy
            C = larrCol(0, I, j) * 0.3 + larrCol(1, I, j) * 0.59 + larrCol(2, I, j) * 0.11
            SetPixel pic.hdc, I, j, RGB(C, C, C)
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    pic.Refresh
    PG1.Visible = False
    PG1.Value = 0
End Sub
Sub 漫画风格(pic As PictureBox)
    Dim Col As Long
    On Error Resume Next
    PG1.Visible = True
    For I = 0 To pic.Width
        For j = 0 To pic.Height
        DoEvents
            Col = GetPixel(pic.hdc, I, j)
            r = Abs(Col Mod 256)
            G = Abs((Col \ 256) Mod 256)
            b = Abs(Col \ 256 \ 256)
            r = Abs(r * (G - b + G + r)) / 256
            G = Abs(r * (b - G + b + r)) / 256
            b = Abs(G * (b - G + b + r)) / 256
            Col = RGB(r, G, b)
            r = Abs(Col Mod 256)
            G = Abs((Col \ 256) Mod 256)
            b = Abs(Col \ 256 \ 256)
            r = (r + G + b) / 3
            Col = RGB(r, r, r)
            SetPixel pic.hdc, I, j, Col
        Next j
        DoEvents
        PG1.Value = I * 100 \ (pic.Width - 1)
    Next I
    PG1.Value = 0
    PG1.Visible = False
    pic.Refresh
    
End Sub
Sub 黑白(pic As PictureBox)
    Dim Col As Long
    On Error Resume Next
    PG1.Visible = True
    For I = 0 To pic.Width
        For j = 0 To pic.Height
        DoEvents
            Col = GetPixel(pic.hdc, I, j)
            r = Col Mod 256
            G = (Col Mod 256) \ 256
            b = Col \ 256 \ 256
            
            If r < 200 And G < 200 And b < 200 Then
                Col = vbBlack
            Else
                Col = vbWhite
            End If
            SetPixel pic.hdc, I, j, Col
        Next j
        DoEvents
        PG1.Value = I * 100 \ (pic.Width - 1)
    Next I
    PG1.Value = 0
    PG1.Visible = False
    pic.Refresh
End Sub
Sub 增加噪点(pic As PictureBox)
On Error Resume Next
    Dim tColR1 As Long, tColR2 As Long, tColR3 As Long, tColR4 As Long, tColR5 As Long
    On Error Resume Next
    PG1.Visible = True
    For I = 0 To cx
        For j = 0 To cy
            tColR1 = GetPixel(pic.hdc, I, j)
            tColR2 = GetPixel(pic.hdc, I + 1, j)
            tColR3 = GetPixel(pic.hdc, I - 1, j)
            tColR4 = GetPixel(pic.hdc, I, j + 1)
            tColR5 = GetPixel(pic.hdc, I, j - 1)
            SetPixel pic.hdc, I, j, (Abs(tColR1) - (Abs(tColR2 + tColR3 + tColR4 + tColR5) / 4))
        Next j
        PG1.Value = I * 100 \ (cx - 1)
    Next I
    PG1.Value = 0
    PG1.Visible = False
    pic.Refresh
End Sub
Sub 浏览()
Dim BFPATH As String
BFPATH = BrowseFolder("浏览文件夹", Me)
If BFPATH = "" Then Exit Sub
filHidden.Path = BFPATH
If filHidden.ListCount > 0 Then lRet = SetInitEntry("PIC_EDIT", "PATH", BFPATH)
End Sub
Sub MOVENOW()
X1.Visible = True
X2.Visible = False
X3.Visible = False
If PIC_MAIN.Visible = True Then PIC_MAIN.Visible = False
If IS_MV_ON = True Then
IS_MV_ON = False
Dim I As Integer
For I = 0 To PBK.Count - 1
PBK(I).Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK(I).hdc, 0, 0)
Next
End If
End Sub
'字符画
Sub 字符画()
    Dim MyJianQB As String  '用于暂时存放Text1，以后可用剪贴板
    Dim MyJingD As Integer  '用于精度
    Dim MyRGBRe As String     '当前RGB所替换的字符
    Dim MyColor As Long   '当前获取颜色
    Dim MyR As Long   '调整R色
    Dim MyG As Long   '调整G色
    Dim MyB As Long   '调整B色
    Dim MyRGB As Long   '当前调整颜色
    If CHECK2.Value = 1 Then
        PicP.BackColor = &H80000009  '无底色
    Else
        PicP.BackColor = &H8000000F '有底色
    End If
    picProgress.Visible = True
    PicP.Cls
    PicP.FOREColor = &H0&  '字体暂为黑色
    MyJingD = Trim(Text4.Text) '取其精度
    MyWH = pic.image.Width / pic.image.Height '原图宽与高之比例
    pic.Scale (0, 0)-(MyJingD, MyJingD / MyWH)
    PicP.Scale (0, 0)-(MyJingD, MyJingD / MyWH)
    For I = 0 To MyJingD / MyWH - 1
        For j = 0 To MyJingD - 1
            MyColor = pic.POINT(j, I)
            MyR = ((MyColor Mod 256) * 2 \ 256) * 255  '用“MyColor Mod 256”公式取其R色，当小于128时调整为0，大于等于128时为255
            MyG = (((MyColor \ 256) Mod 256) * 2 \ 256) * 255
            MyB = ((MyColor \ 65536) * 2 \ 256) * 255
            
            MyRGB = RGB(MyR, MyG, MyB)   '把当前所获取颜色调整为八种基本色
            Select Case MyRGB
                
                Case RGB(0, 0, 0) '黑
                    MyRGBRe = Text2.Text '"M"
                Case RGB(255, 0, 0) '红
                    MyRGBRe = Text3.Text '"A"
                Case RGB(0, 255, 0) 'G为绿
                    MyRGBRe = Text5.Text '"#"
                Case RGB(255, 255, 0) '黄
                    MyRGBRe = Text6.Text '"9"
                Case RGB(0, 0, 255) '蓝
                    MyRGBRe = Text7.Text '"l"
                Case RGB(255, 0, 255) '紫红
                    MyRGBRe = Text8.Text '"o"
                Case RGB(0, 255, 255) '青蓝
                    MyRGBRe = Text9.Text '":"
                Case RGB(255, 255, 255) '白
                    MyRGBRe = Text10.Text '"'"
                Case Else
                
            End Select
            
            PicP.CurrentX = j
            PicP.CurrentY = I
            If CHECK3.Value = 1 Then PicP.FOREColor = MyColor  '设字符为原色
            PicP.Print MyRGBRe
            On Error Resume Next
            picProgressSlide.Width = Me.ScaleWidth / (I * MyJingD + j) * 100
            MyJianQB = MyJianQB & MyRGBRe
        Next j
        MyJianQB = MyJianQB & vbCrLf
    Next I
    
    Text1.Text = MyJianQB
    pic.Visible = False
    PicP.Visible = True
    picProgress.Visible = False
End Sub

Private Sub COPY_ZF()
    Clipboard.Clear
    Clipboard.SetText Text1.Text
    If Text1.Text = "" Then
    Call SHOWWRONG("剪贴板无内容", 2)
    Exit Sub
    End If
    Call SHOWWRONG("已复制到剪贴板,但非彩色", 1)
End Sub

Private Sub Text4_LostFocus()
    On Error GoTo ErrHandle
    pic.Visible = True

    If Trim(Text4.Text) < 20 Or (Text4.Text) > 151 Then
        Call SHOWWRONG("数据必须在20至150之间!!!", 0)
        Text4.Text = 50
        Text4.SetFocus
        pic.Visible = True
    End If
    
    Exit Sub
ErrHandle:
    Call SHOWWRONG("必须是数据!!!", 0)
    Text4.Text = "50"
    Text4.SetFocus
    pic.Visible = True

End Sub

Sub MyOpen(filename As String)
    pic.PICTURE = LoadPicture(filename)
    pic.AUTOSIZE = True
    pic.AutoRedraw = True
    pic.Visible = True
    MyWH = pic.image.Width / pic.image.Height
    If MyWH > Screen.Width / Screen.Height Then
        PicP.Width = Screen.Width
        PicP.Height = Screen.Width / MyWH
    Else
        PicP.Height = Screen.Height
        PicP.Width = Screen.Height * MyWH
    End If
    PicP.Visible = False
End Sub

