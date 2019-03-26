VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2745
   ClientLeft      =   2970
   ClientTop       =   4920
   ClientWidth     =   15435
   LinkTopic       =   "Form4"
   ScaleHeight     =   2745
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image2 
      Height          =   735
      Left            =   13560
      Top             =   105
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1980
      Left            =   0
      Top             =   0
      Width           =   2025
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)


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
SetLayeredWindowAttributes Me.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY '这里的255是透明度，0-255之间
Image1.Picture = LoadPicture(App.Path & "\image\jsq.gif")
Image2.Picture = LoadPicture(App.Path & "\image\jsq0.bmp")
Form4.Left = Form2.Left - 1785
Form4.Top = 6795
End Sub
