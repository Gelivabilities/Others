VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   930
   ClientLeft      =   3000
   ClientTop       =   4410
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7680
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   0
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
   Begin VB.Image Image1 
      Height          =   900
      Index           =   0
      Left            =   11520
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
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




'Private Sub Command1_Click()
'whichmode (modeNum)
'started = True
'End Sub
Sub key_down(ByVal code As Long)
   If (code = 70 Or code = 74) Or code = 192 Then
whichmode (modeNum)
started = True
End If
If (code = 68 Or code = 75) Or code = 109 Then
whichmode (modeNum)
started = True
End If
End Sub

Sub key_up(ByVal code As Long)
    Me.Command1.Caption = code
End Sub



Private Sub Form_Load()
Me.BackColor = &H0
SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes Me.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY '�����255��͸���ȣ�0-255֮��


' RegHook


Form2.Show vbModeless, Form5



Form3.Show vbModeless, Form2

Form1.Show vbModeless, Form2

Form6.Show vbModeless, Form2

Form4.Show vbModeless, Form1
For i = 0 To 16
randomimage (i)
Next

For i = 0 To 16
Form6.Image1(i).Picture = Image1(i).Picture
Next

movet = 0

On Error Resume Next
    Dim myval, myval1, myval2 As Long
    
   myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
   
   myval = SetWindowPos(Form2.hwnd, -1, 0, 0, 0, 0, 3)
   myval = SetWindowPos(Form3.hwnd, -1, 0, 0, 0, 0, 3)
   myval = SetWindowPos(Form4.hwnd, -1, 0, 0, 0, 0, 3)


    
modeNum = 1
biansu = False
step = 50
started = False
stopped = False


Form6.Top = Form1.Top
Form6.Left = Form1.Left
visibleForm = 1
End Sub

Private Function randomimage(i As Integer)
Randomize
t = Int(Rnd * 2) + 1
If t Mod 2 = 0 Then
Image1(i).Picture = LoadPicture(App.Path & "\image\dong.gif")

Else
Image1(i).Picture = LoadPicture(App.Path & "\image\ka.gif")
End If

End Function
Private Sub movecolor()
For i = 16 To 1
Image1(i).Picture = Image1(i - 1).Picture
Next
Do While i > 0
Image1(i).Picture = Image1(i - 1).Picture
i = i - 1
Loop
randomimage (0)
End Sub

Private Sub Timer1_Timer() '����ģʽ����ģʽ

Form1.Width = 11520 + Form2.Left - Form1.Left + 1200

Form1.Left = Form1.Left - 120
movet = movet + 1
If movet = 1 Then Image1(16).Visible = False
If movet = 4 Then Form2.Image1(16).Visible = False
If movet = 6 Then
movet = 0
Image1(16).Visible = True
Form1.Left = Form2.Left + 1000

movecolor
Command1.SetFocus
Form1.Width = 11520 + Form2.Left - Form1.Left + 1200
Form2.Image1(16).Visible = True
Timer1.Enabled = False
End If
End Sub

'Private Sub command1_keydown(keycode As Integer, shift As Integer)


'If (keycode = 70 Or keycode = 74) Or keycode = 192 Then
'Call Command1_Click
'End If
'If (keycode = 68 Or keycode = 75) Or keycode = 109 Then
'Call Command1_Click
'End If
'End Sub

Private Function whichmode(i As Integer)
If stopped = False Then
Select Case i
Case 1

Form2.Image1(16).Picture = Image1(16).Picture
If Timer1.Enabled = False Then
Timer1.Enabled = True
Else
Form1.Left = Form2.Left + 1000
movecolor
Form1.Width = 11520 + Form2.Left - Form1.Left + 600
End If
times = times + 1
If times = step Then
stopped = True

Form1.Width = 11520 + Form2.Left - Form1.Left + 600
End If

Case 3
movet = 0

'һ����ʾ����һ�����벻��ʾ
If visibleForm = 1 Then
visibleShow (6)

Else
If visibleForm = 6 Then
visibleShow (1)
End If
End If


'����˲����ʾ������ɫ
If visibleForm = 1 Then
Form2.Image1(16).Picture = Image1(16).Picture
End If

If visibleForm = 6 Then
Form2.Image1(16).Picture = Form6.Image1(16).Picture
End If
'************************************************************

'��Ϊ�������£�˲����ʾ����
Form2.Image1(16).Visible = True



'If visibleform=1 Then '���°�����˭��ʾ˭��Ҫ����720
'Form1.Left = Form1.Left + 720
'Else
'Form6.Left = Form6.Left + 720
'End If

'��form6��form1������ȫ��ͬ,˭����ʾ��˭
If visibleForm = 1 Then
For i = 1 To 16
Form6.Image1(i).Picture = Image1(i).Picture
Next
Else
If visibleForm = 6 Then
For i = 1 To 16
Image1(i).Picture = Form6.Image1(i).Picture
Next
End If
End If

If visibleForm = 1 Then '�жϣ�˭��ʾ˭����ǰ��
'form6��form1�ұ�
Form6.Left = Form1.Left + 720
Else
'form1��form6�ұ�
Form1.Left = Form6.Left + 720
End If

'�˶��ؼ�������״̬
Timer2.Enabled = True






If visibleForm = 1 Then '�жϣ�˭����ʾ˭��ɫҪ��һλ
movecolor2 'form6��ɫ��һλ
Else
If visibleForm = 6 Then
movecolor 'form1��ɫ��һλ
End If
End If

End Select
End If

End Function

Private Sub Timer2_Timer()



If visibleForm = 1 Then 'form6���ɼ����ƶ�form1
Form1.Width = 11520 + Form2.Left - Form1.Left + 1200
Form1.Left = Form1.Left - 90



End If

If visibleForm = 6 Then 'form1���ɼ����ƶ�form6
Form6.Width = 11520 + Form2.Left - Form6.Left + 1200
Form6.Left = Form6.Left - 90



End If

If visibleForm = 1 Then '�жϣ�˭��ʾ˭�ͱ�������һ��ǰ��720
'form6��form1�ұ�
Form6.Left = Form1.Left + 720
Else
'form1��form6�ұ�
Form1.Left = Form6.Left + 720
End If


'��һ��
movet = movet + 1
If movet > 2 Then
Form2.Image1(16).Visible = False
End If

If visibleForm = 1 Then
'���ˣ�ֹͣ
If Form1.Left < Form2.Left + 280 Then
Timer2.Enabled = False
stopped = True
Form1.Left = Form2.Left + 300 + 720
visibleShow (1)
End If
'���ˣ�ֹͣ
If Form1.Left > Form2.Left + 2440 Then
Timer2.Enabled = False
stopped = True
Form1.Left = Form2.Left + 300 + 720
visibleShow (1)
End If
End If

If visibleForm = 6 Then
'���ˣ�ֹͣ
If Form6.Left < Form2.Left + 280 Then
Timer2.Enabled = False
stopped = True
Form1.Left = Form2.Left + 300 + 720
visibleShow (1)
End If
'���ˣ�ֹͣ
If Form6.Left > Form2.Left + 2440 Then
Timer2.Enabled = False
stopped = True
Form1.Left = Form2.Left + 300 + 720
visibleShow (1)
End If
End If
End Sub

Private Sub movecolor2()
For i = 16 To 1
Form6.Image1(i).Picture = Form6.Image1(i - 1).Picture
Next
Do While i > 0
Form6.Image1(i).Picture = Form6.Image1(i - 1).Picture
i = i - 1
Loop
randomimage2 (0)
End Sub

Private Function randomimage2(i As Integer)
Randomize
t = Int(Rnd * 2) + 1
If t Mod 2 = 0 Then
Form6.Image1(i).Picture = LoadPicture(App.Path & "\image\dong.gif")

Else
Form6.Image1(i).Picture = LoadPicture(App.Path & "\image\ka.gif")
End If

End Function

Private Function visibleShow(i As Integer)
Select Case i
Case 1
visibleForm = 1
SetLayeredWindowAttributes Form1.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY
SetLayeredWindowAttributes Form6.hwnd, &H0, 0, LWA_ALPHA Or LWA_COLORKEY
Case 6
visibleForm = 6
SetLayeredWindowAttributes Form6.hwnd, &H0, 255, LWA_ALPHA Or LWA_COLORKEY
SetLayeredWindowAttributes Form1.hwnd, &H0, 0, LWA_ALPHA Or LWA_COLORKEY

End Select
End Function