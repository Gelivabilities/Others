VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   14265
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   600
      Top             =   2760
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   450
      Left            =   1920
      TabIndex        =   23
      Top             =   2880
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   10440
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   615
      Left            =   3840
      TabIndex        =   22
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   495
      Left            =   6480
      TabIndex        =   21
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      Height          =   735
      Left            =   11280
      TabIndex        =   20
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   39.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   19
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   17
      Left            =   2040
      TabIndex        =   18
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   16
      Left            =   2760
      TabIndex        =   17
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   15
      Left            =   3480
      TabIndex        =   16
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   14
      Left            =   4200
      TabIndex        =   15
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   13
      Left            =   4920
      TabIndex        =   14
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   12
      Left            =   5520
      TabIndex        =   13
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   11
      Left            =   6360
      TabIndex        =   12
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   10
      Left            =   6960
      TabIndex        =   11
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   9
      Left            =   7680
      TabIndex        =   10
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   8
      Left            =   8520
      TabIndex        =   9
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   7
      Left            =   9240
      TabIndex        =   8
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   6
      Left            =   9960
      TabIndex        =   7
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   5
      Left            =   10680
      TabIndex        =   6
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   4
      Left            =   11400
      TabIndex        =   5
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   3
      Left            =   12120
      TabIndex        =   4
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   2
      Left            =   12840
      TabIndex        =   3
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   375
      Index           =   1
      Left            =   13680
      TabIndex        =   2
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   16
      Left            =   1800
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   15
      Left            =   2520
      Picture         =   "Form1.frx":0D70
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   14
      Left            =   3240
      Picture         =   "Form1.frx":1AE0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   13
      Left            =   3960
      Picture         =   "Form1.frx":2850
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   12
      Left            =   4680
      Picture         =   "Form1.frx":35C0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   11
      Left            =   5400
      Picture         =   "Form1.frx":4330
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   10
      Left            =   6120
      Picture         =   "Form1.frx":50A0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   9
      Left            =   6840
      Picture         =   "Form1.frx":5E10
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   8
      Left            =   7560
      Picture         =   "Form1.frx":6B80
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   7
      Left            =   8280
      Picture         =   "Form1.frx":78F0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   6
      Left            =   9000
      Picture         =   "Form1.frx":8660
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   5
      Left            =   9720
      Picture         =   "Form1.frx":93D0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   4
      Left            =   10440
      Picture         =   "Form1.frx":A140
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   3
      Left            =   11160
      Picture         =   "Form1.frx":AEB0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   2
      Left            =   11880
      Picture         =   "Form1.frx":BC20
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   1
      Left            =   12600
      Picture         =   "Form1.frx":C990
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Index           =   0
      Left            =   13320
      Picture         =   "Form1.frx":D700
      Stretch         =   -1  'True
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)


Private Sub Command1_Click()
Label2.Caption = "按对第一个音符开始计时"
Label4.Caption = 1
For i = 0 To 16
ramdomimage (i)
Next
Label1(0).Caption = 0
End Sub

Private Sub Form_Load()
Call Command1_Click
Combo1.AddItem "定量", 0
Combo1.AddItem "定时", 1
Combo1.AddItem "定速", 2
Combo1.ListIndex = 0

End Sub

Private Sub command1_keydown(keycode As Integer, shift As Integer)
If Label4.Caption = 1 Then

If (keycode = 70 Or keycode = 74) And Label1(17).Caption = 0 Then
moveonce
press (1)
Else
If (keycode = 68 Or keycode = 75) And Label1(17).Caption = 1 Then
moveonce
press (2)
Else
If ((keycode = 68 Or keycode = 75) And Label1(17).Caption = 0) Or ((keycode = 70 Or keycode = 74) And Label1(17).Caption = 1) Then
Label2.Caption = "按错，游戏结束"
Label4.Caption = 0
End If
End If
End If
End If
End Sub

Private Function press(x As Integer)
Dim i As Integer
i = 16
If Label4.Caption = 1 Then

If x = 1 Then
Do While i > 0
Image1(i).Picture = Image1(i - 1).Picture
Label1(i + 1).Caption = Label1(i).Caption
i = i - 1
Loop
ramdomimage (0)
Label1(0).Caption = Label1(0).Caption + 1
Else
End If

If x = 2 Then
Do While i > 0
Image1(i).Picture = Image1(i - 1).Picture
Label1(i + 1).Caption = Label1(i).Caption
i = i - 1
Loop
ramdomimage (0)
Label1(0).Caption = Label1(0).Caption + 1
Else
End If

End If

End Function

Private Sub ramdomimage(i As Integer)
Randomize
t = Int(Rnd * 2) + 1
If t Mod 2 = 0 Then
Image1(i).Picture = LoadPicture(App.Path & "\dong.gif")

Else
Image1(i).Picture = LoadPicture(App.Path & "\ka.gif")
End If

Label1(i + 1).Caption = t Mod 2

End Sub


Private Function moveonce()
Timer2.Enabled = True
End Function

Private Sub Timer2_Timer()
For i = 0 To 15
Image1(i).Left = Image1(i).Left - 180
Next

Image1(16).Top = Image1(16).Top - 300
Image1(16).Left = Image1(16).Left + 400
Label6.Caption = Label6.Caption + 1
If Label6.Caption = 4 Then
Label6.Caption = 0

For i = 0 To 15
Image1(i).Left = Image1(i).Left + 720
Next
Image1(16).Left = 1800
Image1(16).Top = 960
Timer2.Enabled = False
End If
End Sub
