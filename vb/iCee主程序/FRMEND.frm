VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FRMEND 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "任务管理器"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10545
   Icon            =   "FRMEND.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PMAIN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   480
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   633
      TabIndex        =   18
      Top             =   1320
      Width           =   9495
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   3735
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   0
         Width           =   8850
         _ExtentX        =   12515
         _ExtentY        =   4471
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   3495
            Left            =   2040
            ScaleHeight     =   3495
            ScaleWidth      =   6615
            TabIndex        =   27
            Top             =   120
            Width           =   6615
            Begin VB.Label lblRegisteredUser 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   49
               Top             =   2640
               Width           =   90
            End
            Begin VB.Label lblRegisteredOrganization 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   48
               Top             =   2940
               Width           =   90
            End
            Begin VB.Label lblProductID 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1380
               TabIndex        =   47
               Top             =   3240
               Width           =   90
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "注册用户:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   8
               Left            =   405
               TabIndex        =   46
               Top             =   2640
               Width           =   810
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "组织:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   9
               Left            =   765
               TabIndex        =   45
               Top             =   2940
               Width           =   450
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "产品序号:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   44
               Top             =   3240
               Width           =   1095
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "更新:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   4
               Left            =   840
               TabIndex        =   43
               Top             =   2280
               Width           =   450
            End
            Begin VB.Label lblOSUpdate 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   42
               Top             =   2280
               Width           =   90
            End
            Begin VB.Label lblOSPlatform 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   41
               Top             =   1680
               Width           =   90
            End
            Begin VB.Label lblOSVersion 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   40
               Top             =   1980
               Width           =   90
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "操作系统:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   2
               Left            =   480
               TabIndex        =   39
               Top             =   1680
               Width           =   810
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "操作系统版本:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   3
               Left            =   120
               TabIndex        =   38
               Top             =   1980
               Width           =   1170
            End
            Begin VB.Label lblProcessorMake 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   37
               Top             =   735
               Width           =   90
            End
            Begin VB.Label lblProcessorModel 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   36
               Top             =   1035
               Width           =   90
            End
            Begin VB.Label lblProcessorSpeed 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   35
               Top             =   1335
               Width           =   90
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "CPU制作商:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   5
               Left            =   360
               TabIndex        =   34
               Top             =   720
               Width           =   900
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "型号:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   6
               Left            =   810
               TabIndex        =   33
               Top             =   1020
               Width           =   450
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "速率:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   7
               Left            =   810
               TabIndex        =   32
               Top             =   1335
               Width           =   450
            End
            Begin VB.Label LA 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "用户名:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   1
               Left            =   570
               TabIndex        =   31
               Top             =   420
               Width           =   630
            End
            Begin VB.Label LA 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "计算机名:"
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Index           =   0
               Left            =   465
               TabIndex        =   30
               Top             =   120
               Width           =   810
            End
            Begin VB.Label lblComputerName 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   29
               Top             =   120
               Width           =   90
            End
            Begin VB.Label lblUserName 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   180
               Left            =   1395
               TabIndex        =   28
               Top             =   420
               Width           =   90
            End
         End
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   3840
         Width           =   1650
         _ExtentX        =   13785
         _ExtentY        =   4471
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   2
         Left            =   2160
         TabIndex        =   21
         Top             =   3840
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   3
         Left            =   3960
         TabIndex        =   22
         Top             =   3840
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   4
         Left            =   5760
         TabIndex        =   23
         Top             =   3840
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   5
         Left            =   360
         TabIndex        =   25
         Top             =   5640
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   6
         Left            =   2160
         TabIndex        =   26
         Top             =   5640
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   7
         Left            =   3960
         TabIndex        =   50
         Top             =   5640
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   8
         Left            =   5760
         TabIndex        =   51
         Top             =   5640
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1650
         Index           =   9
         Left            =   7560
         TabIndex        =   53
         Top             =   3840
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   2910
      End
   End
   Begin VB.PictureBox PICFD 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   120
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   64
      Top             =   1440
      Visible         =   0   'False
      Width           =   10335
      Begin VB.PictureBox PDSK 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AD7900&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   3480
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   449
         TabIndex        =   94
         Top             =   3960
         Width           =   6735
         Begin ICEE.ICEE_KEY IDSK 
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   95
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
         End
      End
      Begin VB.PictureBox picGraph 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   1680
         ScaleHeight     =   225
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   329
         TabIndex        =   67
         Top             =   120
         Width           =   4935
         Begin VB.ListBox LSTDrives 
            Height          =   1860
            Left            =   840
            TabIndex        =   96
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
      End
      Begin VB.TextBox txtFree 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   66
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtUsed 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.TreeView trvDriveView 
         Height          =   2655
         Left            =   8880
         TabIndex        =   93
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   4683
         _Version        =   393217
         Indentation     =   295
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "磁盘基本信息:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   840
         TabIndex        =   92
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "每簇扇区数:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   13
         Left            =   840
         TabIndex        =   91
         Top             =   5640
         Width           =   990
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "每扇区字节数:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   14
         Left            =   840
         TabIndex        =   90
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "簇剩余:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   840
         TabIndex        =   89
         Top             =   6120
         Width           =   630
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "簇总计:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   840
         TabIndex        =   88
         Top             =   6360
         Width           =   630
      End
      Begin VB.Label lblSectorPerClusters 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   87
         Top             =   5640
         Width           =   270
      End
      Begin VB.Label lblBytesPerClusters 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   86
         Top             =   5880
         Width           =   270
      End
      Begin VB.Label lblFreeCluster 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   85
         Top             =   6120
         Width           =   270
      End
      Begin VB.Label lblTotalClusters 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   84
         Top             =   6360
         Width           =   270
      End
      Begin VB.Label lblUsed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   83
         Top             =   4440
         Width           =   270
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已用空间:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   840
         TabIndex        =   82
         Top             =   4440
         Width           =   810
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "剩余空间:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   840
         TabIndex        =   81
         Top             =   4200
         Width           =   810
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "磁盘总大小:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   20
         Left            =   840
         TabIndex        =   80
         Top             =   3960
         Width           =   990
      End
      Begin VB.Label lblVolumeName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   79
         Top             =   4680
         Width           =   270
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "卷标:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   21
         Left            =   840
         TabIndex        =   78
         Top             =   4680
         Width           =   450
      End
      Begin VB.Label lblFileSystem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   77
         Top             =   5160
         Width           =   270
      End
      Begin VB.Label lblSerialNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   76
         Top             =   4920
         Width           =   270
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件系统:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   22
         Left            =   840
         TabIndex        =   75
         Top             =   5160
         Width           =   810
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "序列号:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   23
         Left            =   840
         TabIndex        =   74
         Top             =   4920
         Width           =   630
      End
      Begin VB.Label lblFree 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   73
         Top             =   4200
         Width           =   270
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   72
         Top             =   3960
         Width           =   270
      End
      Begin VB.Label LA 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "字符长度:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   24
         Left            =   840
         TabIndex        =   71
         Top             =   5400
         Width           =   810
      End
      Begin VB.Label lblLenghtString 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2055
         TabIndex        =   70
         Top             =   5400
         Width           =   270
      End
      Begin VB.Label lblPercentUsed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   7200
         TabIndex        =   69
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblPercentFree 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   7800
         TabIndex        =   68
         Top             =   1200
         Width           =   450
      End
   End
   Begin VB.PictureBox PBK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7560
      Picture         =   "FRMEND.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   6840
      Picture         =   "FRMEND.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   8280
      Picture         =   "FRMEND.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSComctlLib.ImageList IM1 
      Left            =   9480
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEND.frx":0636
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEND.frx":09D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEND.frx":0D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEND.frx":1104
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEND.frx":149E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox IU 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9765
      Picture         =   "FRMEND.frx":1838
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   15
      Width           =   750
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   9480
      Top             =   480
   End
   Begin VB.PictureBox PSEE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox TXTSEE 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00404040&
         Height          =   6855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   120
         Width           =   10095
      End
   End
   Begin VB.PictureBox PTCP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   10335
      Begin MSComctlLib.ListView LvwTcpTable 
         Height          =   6615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   11668
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "进程"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "端口"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "本地IP"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "远程IP"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "端口"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "类型"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "状态"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.PictureBox PSOFT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   10335
      Begin ICEE.ICEE_KEY ICU 
         Height          =   495
         Left            =   8640
         TabIndex        =   13
         Top             =   6480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
      End
      Begin MSComctlLib.ListView lstview 
         Height          =   5775
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   10186
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   4210752
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "已安装的软件"
            Object.Width           =   17198
         EndProperty
      End
      Begin VB.Label lblprogname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Null"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   6000
         Width           =   465
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发行商:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   11
         Top             =   6495
         Width           =   630
      End
      Begin VB.Label lblpub 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   870
         TabIndex        =   10
         Top             =   6510
         Width           =   270
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   9
         Top             =   6780
         Width           =   450
      End
      Begin VB.Label lblprogver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   870
         TabIndex        =   8
         Top             =   6780
         Width           =   270
      End
   End
   Begin VB.PictureBox PCPU 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   10335
      Begin MSComctlLib.ListView lstTasks 
         Height          =   6855
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   12091
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "IM1"
         SmallIcons      =   "IM1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox PICFORM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   120
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   54
      Top             =   1440
      Visible         =   0   'False
      Width           =   10335
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
      Begin MSComctlLib.ListView lstWinList 
         Height          =   6735
         Left            =   2640
         TabIndex        =   55
         Top             =   360
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   11880
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "窗口句柄"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "窗口名称"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "是否可见"
            Object.Width           =   38100
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "是否激活"
            Object.Width           =   38100
         EndProperty
      End
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   58
         Top             =   1320
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   59
         Top             =   1800
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   60
         Top             =   2280
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   61
         Top             =   2760
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   62
         Top             =   3240
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
      Begin ICEE.ICEE_KEY ICF 
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   63
         Top             =   3720
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   873
      End
   End
   Begin VB.Label LC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件夹信息"
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
      Left            =   1560
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "FRMEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
 Const PROCESS_PRIORITY_IDLE = 4
 Const PROCESS_PRIORITY_NORMAL = 8
 Const PROCESS_PRIORITY_HIGH = 13
 Const PROCESS_PRIORITY_REALTIME = 24
Private Const HIGH_PRIORITY_CLASS = &H80                    ' Hogs CPU over idle and normal classes
Private Const IDLE_PRIORITY_CLASS = &H40                    ' Only runs when the CPU is idle
Private Const NORMAL_PRIORITY_CLASS = &H20                  ' Duh!
Private Const REALTIME_PRIORITY_CLASS = &H100               ' Highest priority. Even pre-empts operating system
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_TERMINATE = &H1&                       ' Used to kill a process
Private Const PROCESS_CREATE_THREAD = &H2&
Private Const PROCESS_VM_OPERATION = &H8&
Private Const PROCESS_VM_READ = &H10&
Private Const PROCESS_VM_WRITE = &H206
Private Const PROCESS_DUP_HANDLE = &H40&
Private Const PROCESS_CREATE_PROCESS = &H80&
Private Const PROCESS_SET_QUOTA = &H100&
Private Const PROCESS_SET_INFORMATION = &H200&               ' Used to set information on a process (like priority)
Private Const PROCESS_QUERY_INFORMATION = &H400&
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwIdProc As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
Private Declare Function Process32Next Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hnd As Long) As Boolean
Private Type Clsid
    id(16) As Byte
End Type
Private Type ProcessEntry
    dwSize As Long
    peUsage As Long
    peProcessID As Long
    peDefaultHeapID As Long
    peModuleID As Long
    peThreads As Long
    peParentProcessID As Long
    pePriority As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Dim hnd                             As Long         ' Handle to a process
Dim lRet                            As Long         ' Return value for API calls
Dim lExitCode                       As Long         ' Exit code
Dim lPriority                       As Long         ' Priority
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Const KEY_QUERY_VALUE = &H1
Private Const HKEY_DYN_DATA = &H80000006
Const RK_Processor = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
Private Const RK_Performance = "PerfStats\StatData"
Const RK_WIN32_OS = "SOFTWARE\Microsoft\Windows\CurrentVersion"
Const RK_WIN32_OS_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type
Private tmpVersionInfo As OSVERSIONINFO
Dim tmpRegKey As String, IS_MV As Boolean
Dim tmpBuffer As String * 255
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_DELETE As Long = &H3
Private Const FO_RENAME As Long = &H4
Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_CONFIRMMOUSE As Long = &H2
Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_NOCONFIRMMKDIR As Long = &H200
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'主要用于总在最上
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'关于
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'格式化字节大小
Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String

Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As Clsid, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Type SHELLEXECUTEINFO
    CBSIZE As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Const STILL_ACTIVE = &H103
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPheaplist = &H1
Private Const TH32CS_SNAPthread = &H4
Private Const TH32CS_SNAPmodule = &H8
Private Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Dim ProCount As Long    '当前进程数
Dim RamUse As Long  '当前已用内存
Dim theloop As Long
Dim IntString As String
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" _
                        (ByRef pTcpTable As Any, _
                        ByVal bOrder As Boolean, _
                        ByVal heap As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwFamily As Long) _
                        As Long

Private Declare Function AllocateAndGetUdpExTableFromStack Lib "iphlpapi.dll" _
                        (ByRef pUdpTable As Any, _
                        ByVal bOrder As Boolean, _
                        ByVal heap As Long, _
                        ByVal dwFlags As Long, _
                        ByVal dwFamily As Long) _
                        As Long

'SetTcpEntry函数可以帮助我们删除可疑连接
Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpRow As MIB_TCPROW) As Long

Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer '返回一个以主机字节顺序表达的数. 将主机的无符号短整形数转换成网络字节顺序.
Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Integer) As Long '将主机的无符号短整形数转换成网络字节顺序.
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inadr As Long) As Long '一个表示Internet主机地址的结构
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long '若无错误发生,inet_addr()返回一个无符号长整型数,其中以适当字节顺序存放Internet地址
Private Const MIB_TCP_STATE_DELETE_TCB = 12

Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Private Type MIB_TCPROW_EX
    dwState As Long         '连接状态
    dwLocalAddr As Long     '本地IP地址
    dwLocalPort As Long     '本地端口号
    dwRemoteAddr As Long    '远程IP地址
    dwRemotePort As Long    '远程端口号
    dwProcessId As Long     '进程ID
End Type

Private Type MIB_TCPTABLE_EX
    dwNumEntries As Long        '指出本机安装的网卡数
    table() As MIB_TCPROW_EX    'table指向一系列 MIB_IFROW 结构,每个结构指定了当前网卡的状态.这个结构包括了一些很实用的信息,包括网卡的名字(注意,WCHAR类型),网卡描述字串,最大速率,索引,接收到的字
End Type

Private Type MIB_UDPROW_EX
    dwLocalAddr As Long
    dwLocalPort As Long
    dwProcessId As Long
End Type

Private Type MIB_UDPTABLE_EX
    dwNumEntries As Long
    table() As MIB_UDPROW_EX
End Type
Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal Enable As Long, ByVal CurrentThread As Long, Enabled As Long) As Long
Private Const SE_DEBUG_PRIVILEGE = &H14
Private Const AF_INET = 2

Private Graph As New clsGraph
Public SHOW_NAME_DSK As String


'翻译地址
Public Function IpAddr(ByVal hAddr As Long) As String

    Dim sBuf As String
    Dim Ret As Long

    Ret = inet_ntoa(hAddr)
    sBuf = Space$(lstrlen(Ret))

    If lstrcpy(sBuf, Ret) Then IpAddr = sBuf

End Function


Sub GetSysInfo()
    GetComputerName tmpBuffer, 255
    lblComputerName.Caption = Trim$(tmpBuffer)
'-----------------------------------------------------------------------------------------------------------'
    GetUserName tmpBuffer, 255
    lblUserName.Caption = tmpBuffer
'-----------------------------------------------------------------------------------------------------------'
    tmpVersionInfo.dwOSVersionInfoSize = 148
    GetVersionEx tmpVersionInfo
'-----------------------------------------------------------------------------------------------------------'
    If tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        If tmpVersionInfo.dwMinorVersion = 0 Then
            lblOSPlatform.Caption = "Microsoft Windows '95"
        Else
            lblOSPlatform.Caption = "Microsoft Windows '98"
        End If
    ElseIf tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        If tmpVersionInfo.dwMajorVersion = 4 Then
            lblOSPlatform.Caption = "Microsoft Windows NT"
        Else
            lblOSPlatform.Caption = "Microsoft Windows 2000"
        End If
    End If
'-----------------------------------------------------------------------------------------------------------'
    lblOSVersion.Caption = tmpVersionInfo.dwMajorVersion & "." & _
        Format(tmpVersionInfo.dwMinorVersion, "00") & "." & _
        tmpVersionInfo.dwBuildNumber
    lblOSUpdate.Caption = Left(tmpVersionInfo.szCSDVersion, InStr(1, tmpVersionInfo.szCSDVersion, Chr(0)))
'-----------------------------------------------------------------------------------------------------------'
' Retrieve registration information, this is platform specific                                              '
'-----------------------------------------------------------------------------------------------------------'
    If tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        tmpRegKey = RK_WIN32_OS
    ElseIf tmpVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        tmpRegKey = RK_WIN32_OS_NT
    End If
    lblRegisteredOrganization.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "RegisteredOrganization")
    lblRegisteredUser.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "RegisteredOwner")
    lblProductID.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, tmpRegKey, "ProductID")
'-----------------------------------------------------------------------------------------------------------'
' Retrieve CPU information from the registry                                                                '
'-----------------------------------------------------------------------------------------------------------'
    lblProcessorMake.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "VendorIdentifier")
    lblProcessorModel.Caption = GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "Identifier")
    tmpBuffer = GetKeyValue(HKEY_LOCAL_MACHINE, RK_Processor, "~MHZ")
    If Len(Trim(tmpBuffer)) > 0 Then
        lblProcessorSpeed.Caption = Trim(tmpBuffer) & " MHz"
    End If
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'==========================================================================================================='
' Returns a specified key value from the registry                                                           '
'==========================================================================================================='
Dim lKey As Long
Dim tmpVal As String
Dim tmpKeySize As Long
Dim tmpKeyType As Long
Dim Counter As Integer
'-----------------------------------------------------------------------------------------------------------'
' Set up needed variables                                                                                   '
'-----------------------------------------------------------------------------------------------------------'
    tmpVal = String(1024, 0)
    tmpKeySize = 1024
'-----------------------------------------------------------------------------------------------------------'
' Open the registry key. Any value other than zero means something went wrong                               '
'-----------------------------------------------------------------------------------------------------------'
    If RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_QUERY_VALUE, lKey) <> 0 Then
        GetKeyValue = ""
        RegCloseKey lKey
       ' Exit Function
    End If
'-----------------------------------------------------------------------------------------------------------'
' Retrieve the registry value, any value other than zero means something went wrong                         '
'-----------------------------------------------------------------------------------------------------------'
    If RegQueryValueEx(lKey, SubKeyRef, 0, tmpKeyType, tmpVal, tmpKeySize) Then
        GetKeyValue = ""
        RegCloseKey lKey
        'Exit Function
    End If
'-----------------------------------------------------------------------------------------------------------'
' Extract the useful string from the garble                                                                 '
'-----------------------------------------------------------------------------------------------------------'
    If (Asc(Mid(tmpVal, tmpKeySize, 1)) = 0) Then
        tmpVal = Left(tmpVal, tmpKeySize - 1)
    Else
        tmpVal = Left(tmpVal, tmpKeySize)
    End If
'-----------------------------------------------------------------------------------------------------------'
' If the returned value is a dword we need to format the value to something meaningful                      '
'-----------------------------------------------------------------------------------------------------------'
    If tmpKeyType = 4 Then
        For Counter = Len(tmpVal) To 1 Step -1
            GetKeyValue = GetKeyValue + Hex(Asc(Mid(tmpVal, Counter, 1)))
        Next
        GetKeyValue = Format("&h" + GetKeyValue)
    Else
        GetKeyValue = tmpVal
    End If
'-----------------------------------------------------------------------------------------------------------'
' Clean up                                                                                                  '
'-----------------------------------------------------------------------------------------------------------'
    RegCloseKey lKey
    
End Function

'获得系统system32目录
Public Function GetSysDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetSystemDirectory(temp, Len(temp))
    GetSysDir = Left$(temp, x)
End Function

'获得Win目录
Public Function GetWinDir() As String
    Dim temp As String * 256
    Dim x As Integer
    x = GetWindowsDirectory(temp, Len(temp))
    GetWinDir = Left$(temp, x)
End Function



'获得程序所在目录
Public Function GetApp() As String
    If Right$(App.Path, 1) = "\" Then
        GetApp = App.Path
    Else
        GetApp = App.Path & "\"
    End If
End Function
Public Function GetAppF(str As String)
    If str = "" Then Exit Function
    For I = Len(str) To 1 Step -1
        If Mid$(str, I, 1) = "\" Then
            GetAppF = Left$(str, I - 1)
            Exit For
        End If
    Next
End Function

Public Function FormatLng(ByVal lng As Long) As String
    Dim Buffer As String
    Buffer = Space$(100)
    FormatLng = CheckStr(StrFormatByteSize(lng, Buffer, Len(Buffer)))
End Function

'去掉字符串的结束符
Public Function CheckStr(str As String) As String
    If Right$(str, 1) = Chr(0) Then
        CheckStr = Left$(str, Len(str) - 1)
    Else
        CheckStr = str
    End If
End Function

Public Sub ListProcess()
On Error Resume Next
    Dim I As Long, j As Long, n As Long
    Dim Jssl As Integer
    Dim proc As PROCESSENTRY32
    Dim snap As Long
    Dim exename As String
    Dim Item As ListItem
    Dim lngHwndProcess As Long
    Dim lngModules(1 To 200) As Long
    Dim lngCBSize2 As Long
    Dim lngReturn As Long
    Dim strModuleName As String
    Dim pmc As PROCESS_MEMORY_COUNTERS
    Dim WKSize As Long
    Dim strProcessName As String
    Dim strComment As String   '装载进程注释的字符串
    Dim ProClass As String
    Dim SMJC_S As Boolean  '扫描进程
    '开始进程循环
    snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
    proc.dwSize = Len(proc)
    theloop = ProcessFirst(snap, proc)
    I = 0
    n = 0
    While theloop <> 0
        I = I + 1
        'exename = proc.szExeFile
        lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, proc.th32ProcessID)
        If lngHwndProcess <> 0 Then
            lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCBSize2)
            If lngReturn <> 0 Then
                strModuleName = Space(MAX_PATH)
                lngReturn = GetModuleFileNameExA(lngHwndProcess, lngModules(1), strModuleName, 500)
                strProcessName = Left(strModuleName, lngReturn)
                strProcessName = CheckPath(Trim$(strProcessName))
                If strProcessName <> "" Then
                    j = HaveItem(proc.th32ProcessID)
                    If j = 0 Then  '如果没有该进程
                        '获取短文件名
                        exename = Dir(strProcessName, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
                        
                                Dim hand As Long, id As Long
                            
                            exename = Dir(strProcessName, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
                            
                            '添加进程item
                            Set Item = lstTasks.ListItems.Add(, "ID:" & CStr(proc.th32ProcessID), exename)
                            '进程ID
                            Item.SubItems(1) = proc.th32ProcessID
                            '内存使用
                            pmc.cb = LenB(pmc)
                            lRet = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
                            n = n + pmc.WorkingSetSize
                            WKSize = pmc.WorkingSetSize / 1024
                            Item.SubItems(2) = WKSize
                            '优先级
                            Item.SubItems(5) = GetProClass(proc.th32ProcessID)
                            '进程路径
                            Item.SubItems(6) = strProcessName
                            '进程图标
                            IM1.ListImages.Add , strProcessName, GetIcon(strProcessName)
                            Item.SmallIcon = IM1.ListImages.Item(strProcessName).Key
                            '这里判断是否为系统进程
                            strComment = ""
                        If strComment <> "" Then
                                Item.SubItems(3) = "系统"
                                Item.SubItems(4) = Left$(strComment, 2)
                                Item.SubItems(7) = Mid$(strComment, 4)
                        End If
                     
                    Else    '如果已经有该进程
                        pmc.cb = LenB(pmc)
                        lRet = GetProcessMemoryInfo(lngHwndProcess, pmc, pmc.cb)
                        n = n + pmc.WorkingSetSize
                        WKSize = pmc.WorkingSetSize / 1024
                        If CLng(lstTasks.ListItems.Item(j).SubItems(2)) <> WKSize Then lstTasks.ListItems.Item(j).SubItems(2) = WKSize
                        ProClass = GetProClass(proc.th32ProcessID)
                        If ProClass <> lstTasks.ListItems.Item(j).SubItems(5) Then lstTasks.ListItems.Item(j).SubItems(5) = ProClass
                    End If
                    End If
                    End If
                    End If
        theloop = ProcessNext(snap, proc)
    Wend
    CloseHandle snap
    If I <> ProCount Then
     IW(1).SETTXT lstTasks.ListItems.Count
      ProCount = I
    End If
    If n <> RamUse Then
         IW(6).SETTXT FormatLng(n)
        RamUse = n
    End If
End Sub
'设置进程优先级
Public Function SetProClass(ByVal PID As Long, ByVal ClassID As Long)
On Error Resume Next
    Dim hwd As Long
    hwd = OpenProcess(PROCESS_SET_INFORMATION, 0, PID)
    SetProClass = SetPriorityClass(hwd, ClassID)
End Function

'获取进程优先级
Public Function GetProClass(ByVal PID As Long) As String
On Error Resume Next
    Dim hwd As Long
    Dim rtn As Long
    hwd = OpenProcess(PROCESS_QUERY_INFORMATION, 0, PID)
    rtn = GetPriorityClass(hwd)
    Select Case rtn
    Case IDLE_PRIORITY_CLASS
        GetProClass = "低"
    Case NORMAL_PRIORITY_CLASS
        GetProClass = "标准"
    Case HIGH_PRIORITY_CLASS
        GetProClass = "高"
    Case REALTIME_PRIORITY_CLASS
        GetProClass = "实时"
    Case 16384
        GetProClass = "较低"
    Case 32768
        GetProClass = "较高"
    End Select
End Function


'检查进程是否存在多余的已经结束的进程
Public Sub CheckProcess()
On Error Resume Next
    Dim lExit As Long
    Dim lngHwndProcess As Long
    Dim I As Long, j As Long
    If lstTasks.ListItems.Count > 0 Then
        For I = lstTasks.ListItems.Count To 1 Step -1
            j = CLng(lstTasks.ListItems.Item(I).SubItems(1))
            lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, j)
            If lngHwndProcess <> 0 Then
                GetExitCodeProcess lngHwndProcess, lExit
                If lExit <> STILL_ACTIVE Then lstTasks.ListItems.REMOVE I
            Else
                lstTasks.ListItems.REMOVE I
            End If
        Next
    End If
End Sub

'判断item是否存在
Public Function HaveItem(ByVal itemID As Long) As Long
On Error GoTo aaaa
    HaveItem = lstTasks.ListItems("ID:" & CStr(itemID)).Index
Exit Function
aaaa:
    HaveItem = 0
End Function

Public Function CheckPath(ByVal PathStr As String) As String
On Error Resume Next
    PathStr = Replace(PathStr, "\..\", "")
    If UCase(Left$(PathStr, 12)) = "\SYSTEMROOT\" Then PathStr = GetWinDir & Mid$(PathStr, 12)
    CheckPath = PathStr
End Function


Private Sub Form_Activate()
If Me.BackColor = COLOR_NOR Then Exit Sub
Me.BackColor = COLOR_NOR
PDSK.BackColor = COLOR_NOR
Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is PictureBox Then
PBOX.BackColor = Me.BackColor
End If
Next
For I = 0 To IW.Count - 1
IW(I).HASLINE = False
IW(I).SET_STYLE 2
IW(I).SETCOLOR COLOR_HIGH, COLOR_HIGH
IW(I).SETTXTCOLOR vbWhite, vbWhite
Next
For I = 0 To IDSK.Count - 1
IDSK(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next

For I = 0 To ICF.Count - 1
ICF(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICF(I).L_M_R = 0
Next
ICU.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
IW(0).SETTIP "Windows信息"
IW(0).HASTIP = False
IW(0).SETPNG App.Path & "\SKIN\WIN.PNG", 25, (IW(0).Height - 64) / 2
IW(1).SETPNG App.Path & "\SKIN\PRO.PNG", (IW(1).Width - 64) / 2, (IW(1).Height - 64) / 2
IW(2).SETPNG App.Path & "\SKIN\TCP.PNG", (IW(2).Width - 64) / 2, (IW(2).Height - 64) / 2
IW(3).SETPNG App.Path & "\SKIN\SOFT.PNG", (IW(3).Width - 64) / 2, (IW(3).Height - 64) / 2
IW(4).SETPNG App.Path & "\SKIN\SAV.PNG", (IW(4).Width - 64) / 2, (IW(4).Height - 64) / 2
IW(5).SETPNG App.Path & "\SKIN\CPU.PNG", (IW(5).Width - 64) / 2, (IW(5).Height - 64) / 2
IW(6).SETPNG App.Path & "\SKIN\MEM.PNG", (IW(6).Width - 64) / 2, (IW(6).Height - 64) / 2
IW(7).SETPNG App.Path & "\SKIN\HDSK.PNG", (IW(7).Width - 64) / 2, (IW(7).Height - 64) / 2
IW(8).SETPNG App.Path & "\SKIN\FD.PNG", (IW(8).Width - 64) / 2, (IW(8).Height - 64) / 2
IW(9).SETPNG App.Path & "\SKIN\FORM.PNG", (IW(9).Width - 64) / 2, (IW(9).Height - 64) / 2
Call SHOWINFO
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
Picture1.BackColor = COLOR_HIGH
End Sub

Private Sub Form_Load()
IS_CPU_M = True
IS_MV = True
For I = 0 To IW.Count - 1
IW(I).HASLINE = False
IW(I).SET_STYLE 2
IW(I).SETCOLOR vbBlack, &H554E4
IW(I).SETTXTCOLOR vbWhite, vbWhite
Next
For I = 0 To ICF.Count - 1
ICF(I).M_STYLE = 2
ICF(I).L_M_R = 1
Next
ICF(0).SETTXT "全部显示"
ICF(1).SETTXT "只显示可见的窗口"
ICF(2).SETTXT "只显示激活的窗口"
ICF(3).SETTXT "显示可见的窗口文本"
ICF(4).SETTXT "只显示激活的可见窗口"
ICF(5).SETTXT "只显示激活和不可见的窗口"
ICF(6).SETTXT "只显示未激活的可见窗口"
ICF(7).SETTXT "只显示未激活和不可见的窗口"
ICF(0).IS_SELECT = True

IW(6).SETTIP "内存使用"
IW(5).SETTIP "CPU使用"
IW(3).SETTIP "软件管理"
IW(1).SETTIP "进程管理"
IW(2).SETTIP "TCP管理"
IW(4).SETTIP "文件写入情况"
IW(5).SETTXT "0%"
IW(7).SETTIP "磁盘信息"
IW(8).SETTIP "文件夹信息"
IW(9).SETTIP "窗体信息"
ICU.SETTXT "卸载"
ICU.SETCOLOR vbWhite, &H66FBFF, vbBlack
    lstTasks.ColumnHeaders.Add , , "进程名称", 120
    lstTasks.ColumnHeaders.Add , , "PID", 45
    lstTasks.ColumnHeaders.Add , , "内存(K)", 55
    lstTasks.ColumnHeaders.Add , , "种类", 0
    lstTasks.ColumnHeaders.Add , , "级别", 0
    lstTasks.ColumnHeaders.Add , , "优先", 36
    lstTasks.ColumnHeaders.Add , , "进程路径", 400
    lstWinList.ColumnHeaders(1).Width = lstWinList.Width * 0.2
    lstWinList.ColumnHeaders(2).Width = lstWinList.Width * 0.4
    lstWinList.ColumnHeaders(3).Width = lstWinList.Width * 0.2
    lstWinList.ColumnHeaders(4).Width = lstWinList.Width * 0.2
    EnumCondition = No_Filter
Call REFRESHTCP
Call ListProcess
Call GetSysInfo
Set sKeys = New Collection
Call RtlAdjustPrivilege(SE_DEBUG_PRIVILEGE, 1, 0, 0)
FSubClass (Me.hwnd)
Call SHNotify_Register(Me.hwnd)
Call LoadList
Call INLOAD
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MOVENOW
End Sub

Private Sub Form_Resize()
Dim x       As Long
Dim y       As Long
Dim lIdx    As Long
Dim lCols   As Long
            lCols = Int((PDSK.ScaleWidth) / IDSK(0).Width)
            For lIdx = 0 To IDSK.Count - 1
                x = (lIdx Mod lCols) * IDSK(0).Width
                y = Int(lIdx / lCols) * IDSK(0).Height
                IDSK(lIdx).Move x, y
            Next lIdx
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SHNotify_Unregister
  Call UnFSubClass(hwnd)
IS_CPU_M = False
End Sub

Private Sub ICF_Click(Index As Integer)
    Select Case Index
                
        '不筛选 全部显示
        Case 0:
            EnumCondition = No_Filter
            GetWinInfo
            
        '只显示可见的窗口
        Case 1:
            EnumCondition = Only_Visible
            GetWinInfo
        
        '只显示激活窗口
        Case 2:
            EnumCondition = Only_Enabled
            GetWinInfo
        
        '显示可见的窗口文本
        Case 3:
            EnumCondition = Only_Visible_WinTextNotEmpty
            GetWinInfo
        
        '只显示激活和可见的窗口
        Case 4:
            EnumCondition = Only_Enabled_Visible
            GetWinInfo
        
        '只显示激活和不可见窗口
        Case 5:
            EnumCondition = Only_Enabled_NonVisible
            GetWinInfo
        
        '只显示未激活和可见窗口
        Case 6:
            EnumCondition = Only_Disabled_Visible
            GetWinInfo
        
        '只显示未激活和不可见窗口
        Case 7:
            EnumCondition = Only_Disabled_NonVisible
            GetWinInfo
    End Select

Dim I As Integer
For I = 0 To ICF.Count - 1
ICF(I).IS_SELECT = False
Next
ICF(Index).IS_SELECT = True
End Sub

Private Sub ICU_Click()
On Error GoTo ERR
Dim Progname As String, ProgPub As String, ProgVer As String, strRemove As String
            strRemove = GetString(HKEY_LOCAL_MACHINE, IntString & lstview.SelectedItem.Key, "UninstallString")
            WinExec strRemove, 1
            Exit Sub
ERR:
Call SHOWWRONG("    卸载失败,可能未获得管理员权限或您未选中要卸载的软件", 2)
End Sub

Private Sub IDSK_Click(Index As Integer)
Dim I As Integer
For I = 0 To IDSK.Count - 1
IDSK(I).IS_SELECT = False
Next
IDSK(Index).IS_SELECT = True
SHOW_NAME_DSK = IDSK(Index).MY_TIT
Call SHOWINFO
End Sub

Private Sub IW_Click(Index As Integer)
If Index = 0 Or Index = 5 Or Index = 6 Then Exit Sub
PCPU.Visible = False
PTCP.Visible = False
PSOFT.Visible = False
PSEE.Visible = False
PMAIN.Visible = False
PICFD.Visible = False
PICFORM.Visible = False
LC.Visible = True
Select Case Index
Case 0

Case 1
PCPU.Visible = True
LC.Caption = "进程管理"
Case 2
PTCP.Visible = True
LC.Caption = "TCP连接情况"
Case 3
PSOFT.Visible = True
LC.Caption = "已安装软件"
Case 4
PSEE.Visible = True
LC.Caption = "磁盘读写情况"
Case 7
PICFD.Visible = True
LC.Caption = "硬盘信息"
Case 8
If frmma.Left > FRMEX.Width Then
FRMEX.Move frmma.Left - FRMEX.Width, frmma.Top
Else
FRMEX.Move frmma.Left + frmma.Width, frmma.Top
End If
PMAIN.Visible = True
FRMEX.Show
LC.Visible = False
PBK.Visible = False
Case 9
PICFORM.Visible = True
LC.Caption = "窗体信息"
Call GetWinInfo
End Select

PBK.Visible = True
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblBytesPerClusters_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblFileSystem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblFree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblFreeCluster_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblLenghtString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblPercentFree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblPercentUsed_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblProductID_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblprogname_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblprogver_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblpub_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblRegisteredOrganization_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblRegisteredUser_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblSectorPerClusters_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblSerialNumber_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblTotal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblTotalClusters_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblUsed_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lblVolumeName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub LC_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub lstTasks_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstTasks
   If (ColumnHeader.Index - 1) = .SortKey Then
  .SortOrder = (.SortOrder + 1) Mod 2
  .Sorted = True
   Else
  .Sorted = False
  .SortOrder = 0
  .SortKey = ColumnHeader.Index - 1
  .Sorted = True
   End If
End With

End Sub

Private Sub lstview_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstview
   If (ColumnHeader.Index - 1) = .SortKey Then
  .SortOrder = (.SortOrder + 1) Mod 2
  .Sorted = True
   Else
  .Sorted = False
  .SortOrder = 0
  .SortKey = ColumnHeader.Index - 1
  .Sorted = True
   End If
End With
End Sub
Private Sub IU_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
If IU.PICTURE = Me.X2.PICTURE Then IU.PICTURE = Me.X3.PICTURE
End Sub
Private Sub IU_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If IU.PICTURE = Me.X1.PICTURE Then IU.PICTURE = Me.X2.PICTURE
End Sub
Private Sub IU_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
If IU.PICTURE = Me.X3.PICTURE Then IU.PICTURE = Me.X1.PICTURE
Unload Me
End Sub

Private Sub lstTasks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu Frmm.任务管理
End Sub

Private Sub lstview_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim Progname As String, ProgPub As String, ProgVer As String
            Progname = lstview.SelectedItem.Text
            ProgPub = Trim(GetString(HKEY_LOCAL_MACHINE, IntString & lstview.SelectedItem.Key, "Publisher"))
            ProgVer = Trim(GetString(HKEY_LOCAL_MACHINE, IntString & lstview.SelectedItem.Key, "DisplayVersion"))
    
            If Len(ProgVer) = 0 Or Len(ProgPub) = 0 Then
               lblprogname = Progname
               lblpub = "N/A"
               lblprogver = "N/A"
            Else
               lblprogname = Progname
               lblpub = ProgPub
               lblprogver = ProgVer
            End If
End Sub

Private Sub LvwTcpTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With LvwTcpTable
   If (ColumnHeader.Index - 1) = .SortKey Then
  .SortOrder = (.SortOrder + 1) Mod 2
  .Sorted = True
   Else
  .Sorted = False
  .SortOrder = 0
  .SortKey = ColumnHeader.Index - 1
  .Sorted = True
   End If
End With

End Sub

Private Sub LvwTcpTable_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu Frmm.TCP
End Sub

Private Sub PICFD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub PMAIN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub
Private Sub Timer2_Timer()
        ListProcess
End Sub
Sub DELPRO()
On Error Resume Next
    Dim I As Long, hand As Long, id As Long
    id = CLng(lstTasks.SelectedItem.SubItems(1))
    If id <> 0 Then
        EndPro id
        DoEvents
        DoEvents
        FileDel CStr(lstTasks.SelectedItem.SubItems(3))
    End If
    ListProcess
End Sub
'结束一个进程
Public Sub EndPro(ByVal PID As Long)
On Error Resume Next
    Dim lngHwndProcess As Long
    Dim hand As Long
    Dim exitCode As Long
    hand = OpenProcess(PROCESS_TERMINATE, True, PID)
    TerminateProcess hand, exitCode
    CloseHandle hand
End Sub
Sub ENDIT()
On Error Resume Next
    'Dim I As Long, hand As Long, id As Long
    Shell "taskkill.exe /im" & lstTasks.SelectedItem.SubItems(1) & "/f", vbHide
    'id = CLng(lstTasks.SelectedItem.SubItems(1))
    'If id <> 0 Then
    '    EndPro id
    'End If
    ListProcess
End Sub
Sub PAPERPRO()
ShowProperties lstTasks.SelectedItem.SubItems(6), Me.hwnd
End Sub
Sub FOLDERPRO()
ShellExecute hwnd, "open", GetAppF(lstTasks.SelectedItem.SubItems(6)), "", "", 1
End Sub
Public Function IconToPicture(hIcon As Long) As IPictureDisp    'ICON 转 Picture

Dim cls_id As Clsid
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown

    With new_icon
        .CBSIZE = Len(new_icon)
        .PicType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    Dim CA As ColorConstants
    hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
    If hRes = 0 Then Set IconToPicture = lpUnk
    
End Function


Public Function GetIcon(filename, Optional ByVal SmallIcon As Boolean = True) As IPictureDisp   '获得文件ICON

Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO
    If SmallIcon = True Then
        SHGetFileInfo filename, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_SMALLICON
    Else
        SHGetFileInfo filename, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_LARGEICON
    End If
    hIcon = sh_info.hIcon
    Set icon_pic = IconToPicture(hIcon)
    Set GetIcon = icon_pic
End Function

Public Sub NotificationReceipt(wParam As Long, lParam As Long)
Dim Data As String
  Dim sOut As String
  Dim shns As SHNOTIFYSTRUCT
  sOut = SHNotify_GetEventStr(lParam) & vbCrLf
  MoveMemory shns, ByVal wParam, Len(shns)
  Select Case lParam
    Case SHCNE_FREESPACE
      Dim dwDriveBits As Long
      Dim wHighBit As Integer
      Dim wBit As Integer
      MoveMemory dwDriveBits, ByVal shns.dwItem1 + 2, 4
      wHighBit = Int(Log(dwDriveBits) / Log(2))
      For wBit = 0 To wHighBit
        If (2 ^ wBit) And dwDriveBits Then
          sOut = sOut & Chr$(vbKeyA + wBit) & ":\" & vbCrLf
        End If
      Next
    Case SHCNE_UPDATEIMAGE
      Dim iImage As Long
      
      MoveMemory iImage, ByVal shns.dwItem1 + 2, 4
      sOut = sOut & "Index of image in system imagelist: " & iImage & vbCrLf
    Case Else
      Dim sDisplayname As String
      
      If shns.dwItem1 Then
        sDisplayname = GetDisplayNameFromPIDL(shns.dwItem1)
        If Len(sDisplayname) Then
          sOut = Data & "--" & Now & vbCrLf & sOut & "原始文件名称: " & sDisplayname & vbCrLf
          sOut = sOut & "原始文件路径: " & GetPathFromPIDL(shns.dwItem1) & vbCrLf
        Else
          sOut = sOut & "对原始文件的操作是无效的" & vbCrLf
        End If
      End If
    
      If shns.dwItem2 Then
        sDisplayname = GetDisplayNameFromPIDL(shns.dwItem2)
        If Len(sDisplayname) Then
          sOut = Data & "--" & Now & vbCrLf & sOut & "目标文件名称: " & sDisplayname & vbCrLf
          sOut = sOut & "目标文件路径: " & GetPathFromPIDL(shns.dwItem2) & vbCrLf
        Else
          sOut = sOut & "对目标文件的操作是无效的" & vbCrLf
        End If
      End If
  End Select
  TXTSEE = TXTSEE & sOut & vbCrLf
  TXTSEE.SelStart = Len(TXTSEE)
End Sub
Private Sub LoadList()
On Error Resume Next
Dim StrDisName As String
Dim Icnt As Integer
    IntString = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
    GetKeyNames HKEY_LOCAL_MACHINE, IntString
    For Icnt = 1 To sKeys.Count - 1
        StrDisName = GetString(HKEY_LOCAL_MACHINE, IntString & sKeys(Icnt), "DisplayName")
        If Len(StrDisName) > 0 Then
            lstview.ListItems.Add , sKeys(Icnt), StrDisName, 1, 1
        End If
    Next
    lstview.ColumnHeaders(1).Width = lstview.Width - 50
    Set sKeys = Nothing
    StrDisName = ""
    lstview.ListItems(1).Selected = True
    IW(3).SETTXT lstview.ListItems.Count
End Sub
Sub REFRESHTCP()

    Dim I As Long
    Dim lpMTTE As Long               '指向 MIB_TCPTABLE_EX 的指针
    Dim lpMUTE As Long               '指向MIB_UDPTABLE_EX结构的指针
    Dim lngNum As Long

    Dim MTRE As MIB_TCPROW_EX
    Dim MURE As MIB_UDPROW_EX
    Dim MTTE As MIB_TCPTABLE_EX
    Dim MUTE As MIB_UDPTABLE_EX

    Dim lstvItem As ListItem

    LvwTcpTable.ListItems.Clear

    If AllocateAndGetTcpExTableFromStack(lpMTTE, 1, GetProcessHeap, 2, AF_INET) = 0 Then

        CopyMemory lngNum, ByVal lpMTTE, 4          '获取Tcp连接的数量
        ReDim MTTE.table(lngNum - 1)                '重新分配table
        CopyMemory MTTE.table(0), ByVal (lpMTTE + 4), lngNum * LenB(MTRE) 'lpMTTE + 4是因为lpMTTE指针占用了4个字节，后面才是tcp数据

        For I = 0 To lngNum - 1

            If (I Mod 20) = 0 Then DoEvents

            Set lstvItem = LvwTcpTable.ListItems.Add
            lstvItem.Text = PidToName(MTTE.table(I).dwProcessId)
            lstvItem.SubItems(1) = ntohs(MTTE.table(I).dwLocalPort)
            lstvItem.SubItems(2) = IpAddr(MTTE.table(I).dwLocalAddr)
            lstvItem.SubItems(3) = IpAddr(MTTE.table(I).dwRemoteAddr)
            lstvItem.SubItems(4) = ntohs(MTTE.table(I).dwRemotePort)
            lstvItem.SubItems(5) = "TCP"
            lstvItem.SubItems(6) = ConnState(MTTE.table(I).dwState)
            lstvItem.SmallIcon = 2
        Next

    End If

    If AllocateAndGetUdpExTableFromStack(lpMUTE, 1, GetProcessHeap, 2, AF_INET) = 0 Then

        '以下步骤基本同上
        CopyMemory lngNum, ByVal lpMUTE, 4

        ReDim MUTE.table(lngNum - 1)
        CopyMemory MUTE.table(0), ByVal (lpMUTE + 4), lngNum * LenB(MURE)

        For I = 0 To lngNum - 1

            If (I Mod 20) = 0 Then DoEvents

            Set lstvItem = LvwTcpTable.ListItems.Add
            lstvItem.Text = PidToName(MUTE.table(I).dwProcessId)
            lstvItem.SubItems(1) = ntohs(MUTE.table(I).dwLocalPort)
            lstvItem.SubItems(2) = IpAddr(MUTE.table(I).dwLocalAddr)
            lstvItem.SubItems(3) = "0.0.0.0"
            lstvItem.SubItems(4) = "0"
            lstvItem.SubItems(5) = "UDP"
            lstvItem.SmallIcon = 2
        Next

    End If

    Set lstvItem = Nothing

End Sub

'根据PID找进程名
Private Function PidToName(ByVal PID As Long) As String

    Dim lppe As PROCESSENTRY32
    Dim hSnapShot As Long
    Dim bLoop As Long

    hSnapShot = CreateToolhelpSnapshot(&H2, 0)

    lppe.dwSize = Len(lppe)
    bLoop = ProcessFirst(hSnapShot, lppe)

    Do While bLoop > 0

        If PID = lppe.th32ProcessID Then

            PidToName = lppe.szExeFile
            Exit Do

        End If

        bLoop = ProcessNext(hSnapShot, lppe)

    Loop

    CloseHandle (hSnapShot)

End Function

'连接状态
Private Function ConnState(ByVal lngConn As Long) As String

    Select Case lngConn

        Case 0

        ConnState = "未知"

        Case 1

        ConnState = "已关闭"

        Case 2

        ConnState = "监听中"

        Case 3

        ConnState = "SYN_SENT"

        Case 4

        ConnState = "SYN_RCVD"

        Case 5

        ConnState = "已连接"

        Case 6

        ConnState = "FIN_WAIT1"

        Case 7

        ConnState = "FIN_WAIT2"

        Case 8

        ConnState = "CLOSE_WAIT"

        Case 9

        ConnState = "关闭中"

        Case 10

        ConnState = "LAST_ACK"

        Case 11

        ConnState = "等待"

        Case 12

        ConnState = "DELETE_TCB"

    End Select

End Function

Sub KIILTCP()

    Dim lpTcp As MIB_TCPROW

    If LvwTcpTable.SelectedItem.SubItems(5) = "UDP" Then

        Call SHOWWRONG(" 无法断开UDP,请选择TCP", 2)
        Exit Sub

    End If

    '首先填充MIB_TCPROW
    lpTcp.dwLocalAddr = inet_addr(LvwTcpTable.SelectedItem.SubItems(2))
    lpTcp.dwLocalPort = htons(CLng(LvwTcpTable.SelectedItem.SubItems(1)))
    lpTcp.dwRemoteAddr = inet_addr(LvwTcpTable.SelectedItem.SubItems(3))
    lpTcp.dwRemotePort = htons(CLng(LvwTcpTable.SelectedItem.SubItems(4)))
    lpTcp.dwState = MIB_TCP_STATE_DELETE_TCB

    If SetTcpEntry(lpTcp) = 0 Then

        LvwTcpTable.ListItems.REMOVE LvwTcpTable.SelectedItem.Index
        REFRESHTCP

    Else

        Call SHOWWRONG("无法断开 " & LvwTcpTable.SelectedItem.Text & " 进程的连接", 2)

    End If

End Sub
Private Sub PBK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If IS_MV = False Then
IS_MV = True
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub PBK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
PMAIN.Visible = True
PBK.Visible = False
PCPU.Visible = False
PTCP.Visible = False
PSOFT.Visible = False
PSEE.Visible = False
PICFD.Visible = False
LC.Visible = False
PICFORM.Visible = False
End Sub

Sub MOVENOW()
If IU.PICTURE <> X1.PICTURE Then IU.PICTURE = X1.PICTURE
If IS_MV = True Then
IS_MV = False
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
End If
End Sub
Sub SHOWINFO()
Dim Bytes_Avail As LARGE_INTEGER
Dim Bytes_Total As LARGE_INTEGER
Dim Bytes_Free As LARGE_INTEGER

lReturn = GetDiskFreeSpace(sDrive, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)

lblSectorPerClusters.Caption = lSectorsPerCluster
lblBytesPerClusters.Caption = lBytesPerSector
lblFreeCluster.Caption = lFreeClusters
lblTotalClusters.Caption = lTotalClusters

On Error Resume Next

    GetDiskFreeSpaceEx SHOW_NAME_DSK, Bytes_Avail, Bytes_Total, Bytes_Free

    Dbl_Total = LargeIntegerToDouble(Bytes_Total.LowPart, Bytes_Total.HighPart)
    Dbl_Free = LargeIntegerToDouble(Bytes_Free.LowPart, Bytes_Free.HighPart)

    lblTotal.Caption = SizeString(Dbl_Total)
    lblFree.Caption = SizeString(Dbl_Free)
    lblUsed.Caption = SizeString(Dbl_Total - Dbl_Free)
    
    lblPercentFree.Caption = Format$(Dbl_Free / Dbl_Total, "percent")
    lblPercentUsed.Caption = "已使用" & Format$((Dbl_Total - Dbl_Free) / Dbl_Total, "percent")
    IW(7).SETTXT "可用" & Format$(Dbl_Free / Dbl_Total, "percent")
    txtFree.Text = Format$(Dbl_Free / Dbl_Total) * 100
    txtUsed.Text = Format$((Dbl_Total - Dbl_Free) / Dbl_Total) * 100
    
    Root = SHOW_NAME_DSK
    Volume_Name = Space$(1024)
    File_System_Name = Space$(1024)
DoEvents
If GetVolumeInformation(Root, Volume_Name, Len(Volume_Name), Serial_Number, Max_Component_Length, File_System_Flags, File_System_Name, Len(File_System_Name)) = 0 Then
   picGraph.Cls
    lblPercentFree.Caption = ""
    lblPercentUsed.Caption = ""
    lblVolumeName.Caption = ""
    lblSerialNumber.Caption = ""
    lblFileSystem.Caption = ""
    lblLenghtString.Caption = ""
    lblSectorPerClusters.Caption = ""
    lblBytesPerClusters.Caption = ""
    lblFreeCluster.Caption = ""
    lblTotalClusters.Caption = ""
    Call SHOWWRONG("没有磁盘!", 0)
Exit Sub
    End If
    Dim VolumeNameBuffer As String * 11
    Dim VolumeSerialNumber As Long
    Dim MaximumComponentLength As Long
    Dim FileSystemFlags As Long
    Dim FileSystemNameBuffer As String
   If GetVolumeInformation(Left$(LSTDrives, 2) & "\", VolumeNameBuffer, Len(VolumeNameBuffer), VolumeSerialNumber, MaximumComponentLength, FileSystemFlags, FileSystemNameBuffer, Len(FileSystemNameBuffer)) = 0 Then Exit Sub
    pos = InStr(Volume_Name, Chr$(0))
    Volume_Name = Left$(Volume_Name, pos - 1)
    lblVolumeName.Caption = Volume_Name
    lblSerialNumber.Caption = Format$(Serial_Number)
    pos = InStr(File_System_Name, Chr$(0))
    File_System_Name = Left$(File_System_Name, pos - 1)
    lblFileSystem.Caption = File_System_Name
lblLenghtString.Caption = Format$(Max_Component_Length)
    Graph.AddSegment txtFree.Text, "剩余空间", &HDBA349               'Magenta'
    Graph.AddSegment txtUsed.Text, "使用空间", &H7424F9               'Blue'
    Graph.DrawPie picGraph.hdc, picGraph.hwnd, False, ""
    Graph.Clear
picGraph.Refresh

End Sub

Sub INLOAD()
lBuffer = 26 * 4 + 1
sDriveNames = Space$(lBuffer)
lReturn = GetLogicalDriveStrings(lBuffer, sDriveNames)
nOffset = 1
Do
sTempStr = Mid$(sDriveNames, nOffset, 3)
If Left$(sTempStr, 1) = vbNullChar Then Exit Do
LSTDrives.AddItem UCase(sTempStr)
nOffset = nOffset + 4
Loop
LSTDrives.ListIndex = 0
IDSK(0).IS_SELECT = True
Dim lIdx    As Long, lPicCnt As Integer
    While IDSK.Count > 1
        Unload IDSK(IDSK.Count - 1)
    Wend
    DoEvents
        For lIdx = 0 To LSTDrives.ListCount - 1
            ERR.Clear
                If lPicCnt > 0 Then
                    Load IDSK(lPicCnt)
                    Set IDSK(lPicCnt).Container = PDSK
                End If
                DoEvents
                IDSK(lPicCnt).L_M_R = 0
                IDSK(lPicCnt).SETTXT LSTDrives.List(lIdx)
                IDSK(lPicCnt).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
                IDSK(lPicCnt).Visible = True
                 IDSK(lPicCnt).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
                 lPicCnt = lPicCnt + 1
        Next lIdx
        Call Form_Resize
End Sub
Private Sub LSTDrives_Click()
On Error Resume Next
Call SHOWINFO
trvDriveView.Nodes.Clear
    trvDriveView.Nodes.Add , , "main", "该驱动器包含的目录" & SHOW_NAME_DSK
    '获取根目录文件夹...
    trvDriveView.Nodes(1).Expanded = True
    trvDriveView.Refresh
    DoEvents
    'Call ShowFolderList(SHOW_NAME_DSK, "")
End Sub

Sub ShowFolderList(ByVal mvDrive As String, ByVal mvPath As String)
On Error Resume Next
    If Right(mvPath, 10) <> "|NONE|HERE" Then
        Dim fs, f, F1, fc, s
        Dim mvFound As Boolean
        mvFound = False
        Set fs = CreateObject("Scripting.FileSystemObject")
        If mvPath = "main" Then mvPath = ""
        Set f = fs.GetFolder(mvDrive & mvPath)
        Set fc = f.SubFolders
        For Each F1 In fc
            mvFound = True
            If mvPath = "" Then
            DoEvents
                trvDriveView.Nodes.Add "main", tvwChild, F1.name, F1.name & " (" & Format((F1.Size / 1024) / 1024, "#0.00") & " MB)"
            Else
            DoEvents
                trvDriveView.Nodes.Add mvPath, tvwChild, mvPath & "\" & F1.name, F1.name & " (" & Format((F1.Size / 1024) / 1024, "#0.00") & " MB)"
            End If
            trvDriveView.Refresh
         DoEvents
        Next
         DoEvents
        If mvFound = False Then
        
            If mvPath = "" Then
            DoEvents
                trvDriveView.Nodes.Add "main", tvwChild, "|NONE|HERE", "当前路径没有文件夹!"
            Else
            DoEvents
                trvDriveView.Nodes.Add mvPath, tvwChild, mvPath & "\|NONE|HERE", "当前路径没有文件夹!"
            End If
        End If
    End If
End Sub

Private Sub trvDriveView_NodeClick(ByVal Node As MSComctlLib.Node)
    ShowFolderList SHOW_NAME_DSK, Node.Key
End Sub


