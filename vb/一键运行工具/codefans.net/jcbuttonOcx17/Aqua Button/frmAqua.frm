VERSION 5.00
Begin VB.Form frmAqua 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Aqua Button"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin Aqua_Button.jcbutton cmdClose 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ButtonStyle     =   9
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14800597
      Caption         =   "Close"
      CaptionEffects  =   0
      ColorScheme     =   2
   End
   Begin Aqua_Button.AquaButton AquaButton2 
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Thanks to Fred.CPP"
      CaptionEffects  =   0
      RightToLeft     =   -1  'True
   End
   Begin Aqua_Button.AquaButton AquaButton1 
      Height          =   1995
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3519
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      PictureNormal   =   "frmAqua.frx":0000
      CaptionEffects  =   0
      MaskColor       =   16777215
      ToolTip         =   "Tooltip?"
      TooltipTitle    =   "Tooltip Title"
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "and also can be used along with jcButton :-)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "This Aqua Button has all the features of the jcButton"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   3720
      Width           =   2535
   End
End
Attribute VB_Name = "frmAqua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
'Download by http://www.codefans.net
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub

