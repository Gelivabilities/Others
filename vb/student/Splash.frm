VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3420
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraEdge 
      Height          =   3330
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "开发环境：Visual Basic 6.0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2640
         TabIndex        =   6
         Top             =   1560
         Width           =   3450
      End
      Begin VB.Image imgLogo 
         Height          =   825
         Left            =   360
         Picture         =   "Splash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   735
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "版权所有，违者必究！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   480
         TabIndex        =   2
         Top             =   2640
         Width           =   2100
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "1.0.0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4680
         TabIndex        =   3
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "数据环境：Access"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Top             =   2040
         Width           =   2100
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "授权给： 任何给本系统提出宝贵意见的人"
         Height          =   180
         Index           =   5
         Left            =   480
         TabIndex        =   1
         Top             =   3000
         Width           =   3330
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "学生信息管理系统"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   4320
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'该窗体有两个作用，一为系统启动时的窗体，二为系统运行时的“关于...”窗体，而mbAbout即为标识
'若mbAbout为true, 则表示为系统启动时的窗体
'若mbAbout为false，则表示为系统运行时的“关于...”窗体
Public mbAbout As Boolean

Sub UnloadForm()
    Unload Me
    ''如果当前为系统启动时所显示窗体，则在退出本窗体之后，需要加载登录窗体
    If Not mbAbout Then frmLogin.Show
End Sub
'以下各代码，表示：如果点击窗体上的任何部分，或者按下任一个键，都会调用UnloadForm子程序
Private Sub Form_Click()
    UnloadForm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    UnloadForm
End Sub

Private Sub fraEdge_Click()
   UnloadForm
End Sub

Private Sub imgLogo_Click()
    UnloadForm
End Sub

Private Sub lblInfo_Click(Index As Integer)
    UnloadForm
End Sub
