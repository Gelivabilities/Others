VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2565
   ClientLeft      =   16395
   ClientTop       =   405
   ClientWidth     =   2895
   LinkTopic       =   "Form3"
   ScaleHeight     =   2565
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   240
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   840
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   1560
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
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

Public timeLast As Integer




Private Sub Form_load() '倒数读秒窗体
    Me.BackColor = &H0
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY '这里的255是透明度，0-255之间

    
    Image1.Picture = LoadPicture(App.Path & "\image\round.gif")
    
End Sub



Private Sub Timer1_Timer()
If timeLast >= 0 Then
    If timeLast / 10 >= 1 Then
        Image2.Picture = LoadPicture(App.Path & "\image\t" & Int(timeLast / 10) & ".bmp")
        Image3.Left = 1560
    Else
        Image2.Visible = False
        Image3.Left = 1200
    End If
    Image3.Picture = LoadPicture(App.Path & "\image\t" & timeLast Mod 10 & ".bmp")
    
    timeLast = timeLast - 1
Else
    Form1.Flag = True
    Form1.Timer2.Enabled = True
    Form3.Top = -99999 '
    Form3.Timer1.Enabled = False
End If
End Sub
