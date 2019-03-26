VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form5"
   ScaleHeight     =   3180
   ScaleWidth      =   9630
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   960
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'间接窗体
Dim i As Integer

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1

Private Sub Form_Load()
Form1.Show
Form2.Show
Form3.Show
Form4.Show

Form2.Width = 6435
Form2.Height = 2385
Form1.Left = Form2.Left + 1020 + 1000 * 15

Form2.Image1(16).Visible = False

i = 0


SetLayeredWindowAttributes Form4.hwnd, &H0, 10, LWA_ALPHA Or LWA_COLORKEY

SetLayeredWindowAttributes Form3.hwnd, &H0, 10, LWA_ALPHA Or LWA_COLORKEY

SetLayeredWindowAttributes Form1.hwnd, &H0, 10, LWA_ALPHA Or LWA_COLORKEY

SetLayeredWindowAttributes Form2.hwnd, &H0, 5, LWA_ALPHA Or LWA_COLORKEY
End Sub

Private Sub Timer1_Timer()
Form2.Width = Form2.Width + 429
Form2.Height = Form2.Height + 159



i = i + 15


Form4.Top = 7020 - i

Form1.Left = Form2.Left + 1020 + 15000 - 1000 * (i / 15)
Form1.Width = 11700 * (i ^ 3 / 225 ^ 3)

SetLayeredWindowAttributes Form4.hwnd, &H0, 10 + i, LWA_ALPHA Or LWA_COLORKEY
SetLayeredWindowAttributes Form3.hwnd, &H0, 10 + i, LWA_ALPHA Or LWA_COLORKEY
SetLayeredWindowAttributes Form1.hwnd, &H0, 10 + i, LWA_ALPHA Or LWA_COLORKEY

If i >= 60 Then
SetLayeredWindowAttributes Form2.hwnd, &H0, i - 50, LWA_ALPHA Or LWA_COLORKEY
End If

If Form2.Width >= 12870 Then
SetLayeredWindowAttributes Form4.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY
SetLayeredWindowAttributes Form2.hwnd, &H0, 200, LWA_ALPHA Or LWA_COLORKEY
Form2.Image1(16).Visible = True
Timer1.Enabled = False
End If
End Sub

