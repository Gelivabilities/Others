VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20490
   FillColor       =   &H80000002&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleMode       =   0  'User
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000A0A0A&
      Height          =   255
      Left            =   20040
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   11280
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   1200
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   18240
      TabIndex        =   0
      Top             =   11040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   11535
      Index           =   3
      Left            =   19800
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   11535
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":02DF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   11280
      Top             =   10320
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   10320
      Top             =   10320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   8400
      Top             =   10320
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":05BE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20490
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":089D
      Stretch         =   -1  'True
      Top             =   11025
      Width           =   20490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal Scan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Const KEYEVENTF_KEYUP = &H2 '释放按键常数

Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1

Private Const SWP_NOSIZE = &H1
Public handle As Long
Public Flag As Boolean



Private Sub Command1_Click()
Form2.addCoins
End Sub

Private Sub Form_load() '黑边+币数窗体

    Flag = True


    
    Form5.Show vbModeless, Form4
    Form1.Show vbModeless, Form5
    Form3.Show vbModeless, Form1
    Form2.Show vbModeless, Form3
    Me.BackColor = &H0
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY '这里的255是透明度，0-255之间
    Image2.Picture = LoadPicture(App.Path & "\Image\coin(s).gif")
    Image3.Picture = LoadPicture(App.Path & "\Image\0.gif")
    Form4.Image1.Picture = LoadPicture(App.Path & "\Image\main.bmp")
    Form4.Image2.Picture = LoadPicture(App.Path & "\Image\insertcoins.bmp")
    Form4.Image3.Picture = LoadPicture(App.Path & "\Image\3.bmp")
    Image4.Left = Int(Form1.Width / 2 - 1300) + 2000
    Form5.Image1.Picture = LoadPicture(App.Path & "\Image\gameover.bmp")

    On Error Resume Next
    Dim myval As Long
    
    Shell "taskkill.exe /f /im explorer.exe"
    
    myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
   
    myval = SetWindowPos(Form2.hwnd, -1, 0, 0, 0, 0, 3)
    myval = SetWindowPos(Form3.hwnd, -1, 0, 0, 0, 0, 3)
    myval = SetWindowPos(Form4.hwnd, -1, 0, 0, 0, 0, 3)
    myval = SetWindowPos(Form5.hwnd, -1, 0, 0, 0, 0, 3)
End Sub

Private Sub Image1_Click(index As Integer)
Form2.addCoins
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

      handle = FindWindowEx(0, 0, vbNullString, Form2.jiroText)
      
      
'If GetForegroundWindow <> handle Then

If Form2.Left > 20000 Then
    SetWindowPos handle, HWND_TOPMOST, 41, 1, 1288, 745, 0
    SetActiveWindow (handle)
End If
End Sub

Private Sub Timer2_Timer()
If Form2.selectingSongs = True Then
    If Flag = True Then
        Call keybd_event(70, 0, 0, 0)
        Sleep (30)
        Call keybd_event(70, 0, KEYEVENTF_KEYUP, 0)
    Else

        Call keybd_event(73, 0, 0, 0)
        Sleep (30)
        Call keybd_event(73, 0, KEYEVENTF_KEYUP, 0)
    End If
    Flag = Not Flag
End If
End Sub

Public Function esc()
        Call keybd_event(27, 0, 0, 0)
        Sleep (30)
        Call keybd_event(27, 0, KEYEVENTF_KEYUP, 0)
End Function

Public Function space()
        Call keybd_event(32, 0, 0, 0)
        Sleep (30)
        Call keybd_event(32, 0, KEYEVENTF_KEYUP, 0)
End Function

