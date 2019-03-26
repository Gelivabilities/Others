VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4770
   ClientLeft      =   2835
   ClientTop       =   4230
   ClientWidth     =   12870
   LinkTopic       =   "Form2"
   ScaleHeight     =   4770
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   900
      Index           =   16
      Left            =   1020
      Top             =   3100
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000006&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   1020
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   655
      Top             =   2645
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   2775
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "Form2"
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
SetLayeredWindowAttributes Me.hwnd, &H0, 200, LWA_ALPHA Or LWA_COLORKEY '这里的255是透明度，0-255之间


Form1.Top = Form2.Top + 3100

Form3.Left = Form2.Left + 200
Form3.Top = Form2.Top + 100

Image2.Picture = LoadPicture(App.Path & "\image\pm.bmp")
Image1(16).Picture = LoadPicture(App.Path & "\image\dong.gif")
Image1(0).Picture = LoadPicture(App.Path & "\image\form2bg.bmp")
End Sub

Private Sub Form_Unload(cancel As Integer)
End
End Sub







