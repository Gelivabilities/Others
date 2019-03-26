VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFeatures 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "控件技术演示"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6975
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PicProps 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   0
      Left            =   360
      ScaleHeight     =   3375
      ScaleWidth      =   6255
      TabIndex        =   26
      Top             =   1800
      Width           =   6255
      Begin VB.CheckBox Check4 
         Caption         =   "Hand Pointer"
         Height          =   255
         Left            =   3720
         TabIndex        =   52
         Top             =   3000
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Frame FraMode 
         Caption         =   "按钮模式"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3720
         TabIndex        =   48
         Top             =   1680
         Width           =   2415
         Begin VB.OptionButton optMode 
            Caption         =   "一般按钮"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optMode 
            Caption         =   "多选按钮模式"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   50
            Top             =   750
            Width           =   2055
         End
         Begin VB.OptionButton optMode 
            Caption         =   "检查框模式"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   495
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "文字 && 对齐"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   3495
         Begin VB.ComboBox cboTextEffects 
            Height          =   315
            ItemData        =   "frmFeatures.frx":0000
            Left            =   1440
            List            =   "frmFeatures.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1440
            Width           =   1815
         End
         Begin VB.ComboBox cboPicAlign 
            Height          =   315
            ItemData        =   "frmFeatures.frx":004E
            Left            =   1440
            List            =   "frmFeatures.frx":006D
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1030
            Width           =   1815
         End
         Begin VB.ComboBox cboTextAlign 
            Height          =   315
            ItemData        =   "frmFeatures.frx":00E5
            Left            =   1440
            List            =   "frmFeatures.frx":00F2
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   620
            Width           =   1815
         End
         Begin VB.TextBox txtCaption 
            Height          =   315
            Left            =   1440
            TabIndex        =   40
            Text            =   "Button B"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "特殊特效"
            Height          =   180
            Left            =   240
            TabIndex        =   47
            Top             =   1500
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "图片对齐"
            Height          =   180
            Left            =   240
            TabIndex        =   45
            Top             =   1110
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "标题对齐"
            Height          =   180
            Left            =   240
            TabIndex        =   43
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "标题"
            Height          =   180
            Left            =   240
            TabIndex        =   42
            Top             =   315
            Width           =   360
         End
      End
      Begin VB.Frame FraPictures 
         Caption         =   "图片"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3720
         TabIndex        =   32
         Top             =   120
         Width           =   2415
         Begin prjButton.jcbutton cmdLoadPicture 
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   33
            Top             =   360
            Width           =   375
            _extentx        =   661
            _extenty        =   450
            buttonstyle     =   2
            font            =   "frmFeatures.frx":010B
            backcolor       =   15523806
            caption         =   "..."
            pictureeffectonover=   0
            pictureeffectondown=   0
            captioneffects  =   0
            tooltipbackcolor=   0
            colorscheme     =   2
         End
         Begin prjButton.jcbutton cmdLoadPicture 
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   35
            Top             =   720
            Width           =   375
            _extentx        =   661
            _extenty        =   450
            buttonstyle     =   2
            font            =   "frmFeatures.frx":0133
            backcolor       =   15523806
            caption         =   "..."
            pictureeffectonover=   0
            pictureeffectondown=   0
            captioneffects  =   0
            tooltipbackcolor=   0
            colorscheme     =   2
         End
         Begin prjButton.jcbutton cmdLoadPicture 
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   37
            Top             =   1080
            Width           =   375
            _extentx        =   661
            _extenty        =   450
            buttonstyle     =   2
            font            =   "frmFeatures.frx":015B
            backcolor       =   15523806
            caption         =   "..."
            pictureeffectonover=   0
            pictureeffectondown=   0
            captioneffects  =   0
            tooltipbackcolor=   0
            colorscheme     =   2
         End
         Begin VB.Label Label2 
            Caption         =   "按下图片"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   38
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "划过图片"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   36
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "正常图片"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame FraAppearance 
         Caption         =   "界面"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   3495
         Begin VB.ComboBox cboStyle 
            Height          =   315
            ItemData        =   "frmFeatures.frx":0183
            Left            =   1440
            List            =   "frmFeatures.frx":01B1
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox cboColor 
            Height          =   315
            ItemData        =   "frmFeatures.frx":0253
            Left            =   1440
            List            =   "frmFeatures.frx":0260
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "按钮样式"
            Height          =   180
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "颜色方案"
            Height          =   180
            Left            =   240
            TabIndex        =   30
            Top             =   690
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox PicProps 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   1
      Left            =   360
      ScaleHeight     =   3375
      ScaleWidth      =   6255
      TabIndex        =   5
      Top             =   1800
      Width           =   6255
      Begin VB.Frame Frame8 
         Caption         =   "Disable Picture Effects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   5895
         Begin prjButton.jcbutton cmdToggleEnable 
            Height          =   320
            Left            =   4440
            TabIndex        =   23
            Top             =   275
            Width           =   1215
            _extentx        =   2143
            _extenty        =   556
            buttonstyle     =   6
            font            =   "frmFeatures.frx":027E
            caption         =   "Toggle Enable"
            mode            =   1
            pictureeffectonover=   0
            pictureeffectondown=   0
            captioneffects  =   0
            tooltipbackcolor=   0
         End
         Begin VB.OptionButton optDisPic 
            Caption         =   "Grayed (Picture Opacity dependent)"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   22
            Top             =   360
            Width           =   2850
         End
         Begin VB.OptionButton optDisPic 
            Caption         =   "Blended"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Picture Opacity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2880
         TabIndex        =   15
         Top             =   1080
         Width           =   3135
         Begin VB.HScrollBar HScrollOpacity 
            Height          =   255
            LargeChange     =   50
            Left            =   960
            Max             =   255
            SmallChange     =   10
            TabIndex        =   19
            Top             =   320
            Value           =   210
            Width           =   2055
         End
         Begin VB.HScrollBar HScrollOpacityOver 
            Height          =   255
            LargeChange     =   50
            Left            =   960
            Max             =   255
            SmallChange     =   10
            TabIndex        =   18
            Top             =   720
            Value           =   255
            Width           =   2055
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "On Over"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "正常"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Other Effects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   2880
         TabIndex        =   8
         Top             =   120
         Width           =   3135
         Begin VB.CheckBox chkPicShadow 
            Caption         =   "Picture Shadow"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkPicPush 
            Caption         =   "Picture Push On Hover"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Picture Down Effects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2415
         Begin VB.OptionButton optPicDownEff 
            Caption         =   "Darker"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optPicDownEff 
            Caption         =   "Lighter"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optPicDownEff 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Picture Over Effects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2415
         Begin VB.OptionButton optPicOverEff 
            Caption         =   "Darker"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton optPicOverEff 
            Caption         =   "Lighter"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optPicOverEff 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox PicProps 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   2
      Left            =   360
      ScaleHeight     =   3255
      ScaleWidth      =   6255
      TabIndex        =   71
      Top             =   1800
      Width           =   6255
      Begin VB.Frame Frame9 
         Caption         =   "DropDown Symbols"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   3480
         TabIndex        =   75
         Top             =   960
         Width           =   2175
         Begin VB.OptionButton optSymbol 
            Caption         =   "Right"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   88
            Top             =   1440
            Width           =   855
         End
         Begin VB.OptionButton optSymbol 
            Caption         =   "Down"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   87
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optSymbol 
            Caption         =   "Up"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   86
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton optSymbol 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   84
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   12
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   90
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   12
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   89
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   12
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   85
            Top             =   1440
            Width           =   255
         End
      End
      Begin VB.CheckBox chkDropDownSep 
         Caption         =   "DropDown Separator"
         Height          =   255
         Left            =   3480
         TabIndex        =   74
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chkDropDownEnabled 
         Caption         =   "DropDown Enabled"
         Height          =   255
         Left            =   3480
         TabIndex        =   73
         Top             =   120
         Width           =   2175
      End
      Begin VB.Frame fraProps 
         Caption         =   "Dropdown Alignments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   3135
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "BottomRight Align"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   83
            Tag             =   "7"
            Top             =   2400
            Width           =   1935
         End
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "TopRight Align"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   82
            Tag             =   "6"
            Top             =   2106
            Width           =   1815
         End
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "BottomLeft Align"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   81
            Tag             =   "5"
            Top             =   1815
            Width           =   1815
         End
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "TopLeft Align"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   80
            Tag             =   "4"
            Top             =   1524
            Width           =   1455
         End
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "Right Align"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   79
            Tag             =   "3"
            Top             =   1233
            Width           =   1335
         End
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "Left Align"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   78
            Tag             =   "2"
            Top             =   942
            Width           =   1575
         End
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "Top Align"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   77
            Tag             =   "1"
            Top             =   651
            Width           =   1695
         End
         Begin VB.OptionButton optMenuAlign 
            Caption         =   "Bottom Align"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   76
            Tag             =   "0"
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox PicProps 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   3
      Left            =   360
      ScaleHeight     =   3375
      ScaleWidth      =   6255
      TabIndex        =   53
      Top             =   1800
      Width           =   6255
      Begin VB.Frame Frame1 
         Caption         =   "Tooltip Styles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   2535
         Begin VB.OptionButton optTooltipStyle 
            Caption         =   "Tooltip Standard"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optTooltipStyle 
            Caption         =   "Tooltip Balloon"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   69
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame fraTooltip 
         Caption         =   "Tooltip Icons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   2520
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1110
            Index           =   0
            Left            =   45
            ScaleHeight     =   1110
            ScaleWidth      =   2265
            TabIndex        =   63
            Top             =   200
            Width           =   2265
            Begin VB.OptionButton optTooltipIcon 
               Caption         =   "Icon None"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   67
               Top             =   90
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optTooltipIcon 
               Caption         =   "Icon Info"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   66
               Top             =   590
               Width           =   975
            End
            Begin VB.OptionButton optTooltipIcon 
               Caption         =   "Icon Warning"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   65
               Top             =   340
               Width           =   1335
            End
            Begin VB.OptionButton optTooltipIcon 
               Caption         =   "Icon Error"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   64
               Top             =   840
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tooltip Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   3000
         TabIndex        =   54
         Top             =   120
         Width           =   3015
         Begin VB.CheckBox chkRTL 
            Caption         =   "Right To Left tooltips"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   2040
            Width           =   2055
         End
         Begin VB.PictureBox picTooltipColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   465
            TabIndex        =   57
            ToolTipText     =   "Select Tooltip BackColor"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtToolText 
            Height          =   285
            Left            =   240
            TabIndex        =   56
            Text            =   "Tooltip Text"
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtToolTitle 
            Height          =   285
            Left            =   240
            TabIndex        =   55
            Text            =   "Tooltip Text"
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提示背景色"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   61
            Top             =   1680
            Width           =   1260
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提示信息"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   60
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提示标题"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   59
            Top             =   360
            Width           =   825
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6615
      TabIndex        =   91
      Top             =   120
      Width           =   6615
      Begin prjButton.jcbutton cmdJCTest 
         Height          =   735
         Index           =   0
         Left            =   495
         TabIndex        =   92
         Top             =   120
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1296
         buttonstyle     =   3
         font            =   "frmFeatures.frx":02A6
         backcolor       =   14935011
         caption         =   "示例按钮 A"
         handpointer     =   -1  'True
         picturenormal   =   "frmFeatures.frx":02CE
         pictureeffectonover=   0
         pictureeffectondown=   0
         pictureopacity  =   210
         captioneffects  =   0
         colorscheme     =   2
      End
      Begin prjButton.jcbutton cmdJCTest 
         Height          =   735
         Index           =   1
         Left            =   2520
         TabIndex        =   93
         Top             =   120
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1296
         buttonstyle     =   2
         font            =   "frmFeatures.frx":0C20
         backcolor       =   15199212
         caption         =   "示例按钮 B"
         handpointer     =   -1  'True
         picturenormal   =   "frmFeatures.frx":0C48
         picturealign    =   2
         pictureeffectonover=   0
         pictureeffectondown=   0
         pictureopacity  =   210
         captioneffects  =   0
         maskcolor       =   16777215
         tooltip         =   "Tooltip Text"
         tooltiptitle    =   "Tooltip Title"
      End
      Begin prjButton.jcbutton cmdJCTest 
         Height          =   735
         Index           =   2
         Left            =   4560
         TabIndex        =   94
         Top             =   120
         Width           =   1815
         _extentx        =   3201
         _extenty        =   1296
         buttonstyle     =   8
         font            =   "frmFeatures.frx":159A
         backcolor       =   16765357
         caption         =   "示例按钮 C"
         handpointer     =   -1  'True
         picturenormal   =   "frmFeatures.frx":15C2
         pictureeffectonover=   0
         pictureeffectondown=   0
         pictureopacity  =   210
         captioneffects  =   0
      End
   End
   Begin prjButton.jcbutton cmdProperties 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   1290
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      buttonstyle     =   8
      font            =   "frmFeatures.frx":1F14
      backcolor       =   16765357
      caption         =   "按下"
      mode            =   2
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   1800
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjButton.jcbutton cmdProperties 
      Height          =   420
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1245
      Width           =   1335
      _extentx        =   2355
      _extenty        =   741
      buttonstyle     =   8
      font            =   "frmFeatures.frx":1F3C
      backcolor       =   16765357
      caption         =   "公用"
      mode            =   2
      value           =   -1  'True
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin prjButton.jcbutton cmdProperties 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   1290
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      buttonstyle     =   8
      font            =   "frmFeatures.frx":1F64
      backcolor       =   16765357
      caption         =   "图片特效"
      mode            =   2
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin prjButton.jcbutton cmdProperties 
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   2
      Top             =   1290
      Width           =   1215
      _extentx        =   2143
      _extenty        =   661
      buttonstyle     =   8
      font            =   "frmFeatures.frx":1F8C
      backcolor       =   16765357
      caption         =   "提示信息"
      mode            =   2
      pictureeffectonover=   0
      pictureeffectondown=   0
      captioneffects  =   0
      tooltipbackcolor=   0
   End
   Begin VB.Frame Frame3 
      Height          =   3735
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Menu MenuDemo 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu m1 
         Caption         =   "DropDown Menu"
      End
      Begin VB.Menu m2 
         Caption         =   "with a dropdown Symbol"
      End
      Begin VB.Menu m3 
         Caption         =   "and a Dropdown Separator"
      End
   End
End
Attribute VB_Name = "frmFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Option Explicit

Private i As Long

Private Sub cboColor_Click()
    cmdJCTest(1).ColorScheme = cboColor.ListIndex
End Sub

Private Sub cboPicAlign_Click()
    cmdJCTest(1).PictureAlign = cboPicAlign.ListIndex
End Sub

Private Sub cboStyle_Click()
    cmdJCTest(1).ButtonStyle = cboStyle.ListIndex
    If cmdJCTest(1).ButtonStyle = eOfficeXP Then
        chkPicPush.Value = vbChecked
    Else
        chkPicPush.Value = vbUnchecked
    End If
    cmdJCTest(1).ColorScheme = cboColor.ListIndex
End Sub


Private Sub cboTextAlign_Click()
    cmdJCTest(1).CaptionAlign = cboTextAlign.ListIndex
End Sub

Private Sub cboTextEffects_Click()
    For i = 0 To 2
        cmdJCTest(i).CaptionEffects = cboTextEffects.ListIndex
    Next i
End Sub

Private Sub chkMenuSep_Click()
    For i = 0 To 2
        cmdJCTest(i).DropDownSeparator = Not cmdJCTest(i).DropDownSeparator
    Next i
End Sub

Private Sub Check4_Click()
    For i = 0 To 2
        cmdJCTest(i).HandPointer = Not cmdJCTest(i).HandPointer
    Next i
End Sub

Private Sub chkDropDownEnabled_Click()

Dim t As Long
    
    ' --Get the option button value
    For i = 0 To 7
        If optMenuAlign(i).Value = True Then
            t = i
        End If
    Next i
    
    'Set Menu
    cmdJCTest(1).SetPopupMenu MenuDemo, t
    If chkDropDownEnabled.Value = vbUnchecked Then
        cmdJCTest(1).UnsetPopupMenu
    End If
    
End Sub

Private Sub chkDropDownSep_Click()
    For i = 0 To 2
        cmdJCTest(i).DropDownSeparator = Not cmdJCTest(i).DropDownSeparator
    Next i
End Sub

Private Sub chkEnabled_Click()

Dim d As Long
    For d = 0 To 2
        cmdJCTest(d).Enabled = Not cmdJCTest(d).Enabled
    Next d
    
End Sub

Private Sub chkPicPush_Click()
    For i = 0 To 2
        cmdJCTest(i).PicturePushOnHover = Not cmdJCTest(i).PicturePushOnHover
    Next i
End Sub

Private Sub chkPicShadow_Click()
    For i = 0 To 2
        cmdJCTest(i).PictureShadow = Not cmdJCTest(i).PictureShadow
    Next i
End Sub

Private Sub chkRTL_Click()
    cmdJCTest(1).RightToLeft = Not cmdJCTest(1).RightToLeft
End Sub

Private Sub cmdJCTest_Click(Index As Integer)
    
    If Index = 1 Then
        If chkDropDownEnabled.Value = vbChecked Then
            cmdJCTest(1).SetPopupMenu MenuDemo, 0
        End If
    End If
    
End Sub

Private Sub cmdPicNormal_Click(Index As Integer)
    
End Sub

Private Sub cmdLoadPicture_Click(Index As Integer)
On Error GoTo h:

    With cdl
        .Filter = "All Picture Files |*.bmp;*.jpg;*.gif;*.ico; *.wmf;|All Files (*.*)|*.*"
        .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
        .ShowOpen
        
        Select Case Index
        Case 0
            Set cmdJCTest(1).PictureNormal = LoadPicture(.FileName)
        Case 1
            Set cmdJCTest(1).PictureHot = LoadPicture(.FileName)
        Case 2
            Set cmdJCTest(1).PictureDown = LoadPicture(.FileName)
        End Select
    End With
 
h:
If Err.Number = 481 Then MsgBox "Invalid Picture", vbCritical + vbOKOnly, "Unsupported Picture"
End Sub

Private Sub cmdProperties_Click(Index As Integer)
    
    For i = 0 To 3
        cmdProperties(i).Font.Bold = False
        cmdProperties(i).Top = 86
        cmdProperties(i).Height = 25
        PicProps(i).Visible = False
    Next i
    
    PicProps(Index).Visible = True
    cmdProperties(Index).Font.Bold = True
    cmdProperties(Index).Top = 83
    cmdProperties(Index).Height = 28
    
    ' --May be disbled some times and user forget from where to enable!!!!!
    For i = 0 To 2
        If Not cmdJCTest(i).Enabled Then
            cmdJCTest(i).Enabled = True
        End If
    Next i
    
End Sub

Private Sub cmdToggleEnable_Click()
    For i = 0 To 2
        cmdJCTest(i).Enabled = Not cmdToggleEnable.Value
    Next i
    
    ' --I don't know why I have to refresh, but without refresh,
    ' --it was not toggling until mouse leaves from the buton (try!)
    ' --May be due to checkbox mode (forget it!)
    Refresh
    
End Sub

Private Sub Form_Load()
    
Dim combo As Control

    For Each combo In frmFeatures.Controls
        If TypeOf combo Is ComboBox Then
            combo.ListIndex = 0
        End If
    Next combo
    
    cboStyle.ListIndex = 2
    cboTextAlign.ListIndex = 1
    cboPicAlign.ListIndex = 6
    cmdJCTest(1).TooltipTitle = txtToolTitle().Text
    cmdJCTest(1).ToolTipText = txtToolText().Text

End Sub



Private Sub Form_Unload(Cancel As Integer)
    frmButtonDemo.Show
End Sub

Private Sub HScrollOpacity_Change()
    For i = 0 To 2
        cmdJCTest(i).PictureOpacity = HScrollOpacity.Value
    Next i
End Sub

Private Sub HScrollOpacityOver_Change()
    For i = 0 To 2
        cmdJCTest(i).PictureOpacityOnOver = HScrollOpacityOver.Value
    Next i
End Sub


Private Sub optDisPic_Click(Index As Integer)
    For i = 0 To 2
        cmdJCTest(i).DisabledPictureMode = Index
    Next i
End Sub

Private Sub optMenuAlign_Click(Index As Integer)
    If chkDropDownEnabled.Value = vbChecked Then
        cmdJCTest(1).SetPopupMenu MenuDemo, optMenuAlign(Index).Tag
    End If
End Sub

Private Sub optMode_Click(Index As Integer)
    For i = 0 To 2
        cmdJCTest(i).Value = False
        cmdJCTest(i).Mode = Index
    Next i
End Sub

Private Sub optPicDownEff_Click(Index As Integer)
    For i = 0 To 2
        cmdJCTest(i).PictureEffectOnDown = Index
    Next i
End Sub

Private Sub optPicOverEff_Click(Index As Integer)
    For i = 0 To 2
        cmdJCTest(i).PictureEffectOnOver = Index
    Next i
End Sub

Private Sub optSymbol_Click(Index As Integer)
    For i = 0 To 2
        cmdJCTest(i).DropDownSymbol = Index
    Next i
End Sub

Private Sub optTooltipIcon_Click(Index As Integer)
    cmdJCTest(1).ToolTipIcon = Index
End Sub

Private Sub optTooltipStyle_Click(Index As Integer)
    cmdJCTest(1).ToolTipType = Index
End Sub

Private Sub picTooltipColor_Click()
    cdl.ShowColor
    picTooltipColor.BackColor = cdl.Color
    cmdJCTest(1).TooltipBackColor = picTooltipColor.BackColor
End Sub

Private Sub txtCaption_Change()
    cmdJCTest(1).Caption = txtCaption.Text
End Sub

Private Sub txtToolText_Change()
    cmdJCTest(1).ToolTip = txtToolText.Text
End Sub

Private Sub txtToolTitle_Change()
    cmdJCTest(1).TooltipTitle = txtToolTitle.Text
End Sub

