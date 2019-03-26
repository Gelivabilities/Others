VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   900
   ClientLeft      =   2505
   ClientTop       =   2145
   ClientWidth     =   11700
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Image Image1 
      Height          =   900
      Index           =   0
      Left            =   11520
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   16
      Left            =   0
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   15
      Left            =   720
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   14
      Left            =   1440
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   13
      Left            =   2160
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   12
      Left            =   2880
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   11
      Left            =   3600
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   10
      Left            =   4320
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   9
      Left            =   5040
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   8
      Left            =   5760
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   7
      Left            =   6480
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   6
      Left            =   7200
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   5
      Left            =   7920
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   4
      Left            =   8640
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   3
      Left            =   9360
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   2
      Left            =   10080
      Top             =   0
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   1
      Left            =   10800
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1

Private Sub Form_Load()
Me.BackColor = &H0
SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes Me.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY
End Sub
