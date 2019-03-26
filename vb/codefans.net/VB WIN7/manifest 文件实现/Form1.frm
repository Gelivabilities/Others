VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB 实现WIN7风格 （manifest 文件实现）"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8655
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1065
      Left            =   7470
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Form1.frx":0000
      Top             =   1275
      Width           =   840
   End
   Begin VB.FileListBox File1 
      Height          =   1530
      Left            =   5610
      TabIndex        =   18
      Top             =   615
      Width           =   1605
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   3180
      TabIndex        =   17
      Top             =   3195
      Width           =   2340
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3045
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   2430
      Width           =   2865
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   3210
      TabIndex        =   15
      Top             =   195
      Width           =   2130
   End
   Begin VB.DirListBox Dir1 
      Height          =   1560
      Left            =   3210
      TabIndex        =   14
      Top             =   555
      Width           =   2205
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   480
      Left            =   3060
      TabIndex        =   13
      Top             =   4710
      Width           =   3360
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Left            =   7860
      TabIndex        =   12
      Top             =   3165
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   1260
      Left            =   6150
      ScaleHeight     =   1200
      ScaleWidth      =   1110
      TabIndex        =   11
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1650
      Left            =   300
      TabIndex        =   7
      Top             =   420
      Width           =   2325
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   270
         Left            =   225
         TabIndex        =   9
         Top             =   780
         Width           =   1410
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   360
         Left            =   225
         TabIndex        =   8
         Top             =   330
         Width           =   1890
      End
      Begin VB.Label Label2 
         Caption         =   "不用加Picture"
         Height          =   315
         Left            =   210
         TabIndex        =   10
         Top             =   1125
         Width           =   1890
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Top             =   3210
      Width           =   1185
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   420
      Left            =   60
      TabIndex        =   5
      Top             =   2655
      Width           =   1320
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   2310
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   630
      Left            =   1725
      TabIndex        =   3
      Top             =   2760
      Width           =   1140
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   675
      Left            =   45
      TabIndex        =   1
      Top             =   3855
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1191
      _Version        =   327682
   End
   Begin ComctlLib.ProgressBar p 
      Height          =   450
      Left            =   165
      TabIndex        =   0
      Top             =   4710
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   794
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "XP风格的manifest文件和win7风格的manifest文件不同，请不要混用！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   75
      TabIndex        =   20
      Top             =   6270
      Width           =   8160
   End
   Begin VB.Label Label1 
      Caption         =   "请使用5.0版的Microsoft.Windows.Common-Controls,在XP里可能无法实现"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   0
      TabIndex        =   2
      Top             =   5250
      Width           =   7905
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Private Sub Command1_Click()
MsgBox "看下面！", vbYesNo + vbCritical
End Sub

Private Sub Form_Load()
p.Value = 50
For i = 1 To 20
List1.AddItem Str(i) & "aa"
Next
End Sub

