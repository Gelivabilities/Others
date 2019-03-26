VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "操作与设置"
   ClientHeight    =   4515
   ClientLeft      =   3885
   ClientTop       =   2505
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   30
      Text            =   "taikojiro.exe"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   29
      Text            =   "taikojiro.exe"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "浏览"
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   22
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "浏览"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   21
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   2
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "本体位置"
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   5775
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   28
         Text            =   "taikojiro.exe"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "保存全部"
         Height          =   1335
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "运行全部"
         Height          =   1335
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "浏览"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Index           =   0
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "敲20次"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "敲10次"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "默认"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "RK"
      Height          =   1935
      Left            =   5040
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "RD"
      Height          =   1935
      Left            =   4200
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LD"
      Height          =   1935
      Left            =   3360
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LK"
      Height          =   1935
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "60"
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "开始游戏"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   840
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "3"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "减币"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加币"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "模拟操作"
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton Command10 
         Caption         =   "选中歌曲"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1200
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "结束游戏"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1200
         TabIndex        =   26
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "选曲限时"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "每盘币数"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "左右交替敲打鼓边10次开启魔王"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label12 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4200
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszsoundname As String, ByVal uflags As Long) As Long

Dim coins, jiaoti, jiroType, jiroExeName(0 To 2) As Integer
Dim youxizhong As Boolean
Public jiroText As String

Sub key_down(ByVal code As Long)
    Select Case code
        Case 71
            Call Command1_Click
        Case 68
            Call Command4_Click
        Case 70
            Call Command5_Click
        Case 74
            Call Command6_Click
        Case 75
            Call Command7_Click
        Case 72
            If Form2.Left <> 25000 Then
                Form2.Left = 25000
            Else
                Form2.Left = 5000
            End If
    End Select
End Sub

Sub key_up(ByVal code As Long)
    
End Sub



Public Function readAddress()
On Error Resume Next

Open App.Path & "\address.txt" For Input As #1

For i = 0 To 2
Line Input #1, s
Text3(i).Text = s
Next

Close #1

End Function

Public Function readJiroExe()
On Error Resume Next

Open App.Path & "\jiroexe.txt" For Input As #1

For i = 0 To 2
Line Input #1, s
Text2(i).Text = s
Next

Close #1

End Function
 
Public Function jiroSelect(i As Integer)
On Error Resume Next
    Shell Text3(i).Text, vbNormalFocus
End Function






Private Sub Command1_Click()
    If coins < 99 Then
    coins = coins + 1
    End If
    
    Form1.Image3.Picture = LoadPicture(App.Path & "\Image\" & (coins Mod 10) & ".gif")
    Form1.Image4.Picture = LoadPicture(App.Path & "\image\" & Int(coins / 10) & ".gif")
    SoundFile = App.Path & "\sound\insertcoins.wav"
    Result = sndplaysound(SoundFile, 1)
    
    If Int(coins / 10) = 0 Then
        Form1.Image4.Visible = False
    Else
        Form1.Image4.Visible = True
    End If
    
    If coins < Int(Form2.Text7.Text) Then
        Form4.Image3.Visible = True
        Form4.Image3.Picture = LoadPicture(App.Path & "\Image\" & Text7.Text - coins & ".bmp")
    Else
            Form4.Image2.Picture = LoadPicture(App.Path & "\Image\hittostart.bmp")
            Form4.Image3.Visible = False
            If youxizhong = False Then
                Command3.Enabled = True
            End If
            
    End If
    
End Sub

Private Sub Command10_Click()
Form3.Visible = False
End Sub

Private Sub Command11_Click()
On Error Resume Next
For i = 0 To 2
jiroSelect (i)
Next

End Sub

Private Sub Command12_Click()
Open App.Path & "\address.txt" For Output As #1
For i = 0 To 2
Print #1, Text3(i).Text
Next
Close #1

Open App.Path & "\jiroexe.txt" For Output As #1
For i = 0 To 2
Print #1, Text2(i).Text
Next
Close #1
End Sub

Private Sub Command2_Click()
    If coins > 0 Then coins = coins - 1
    Form1.Image3.Picture = LoadPicture(App.Path & "\Image\" & (coins Mod 10) & ".gif")
    Form1.Image4.Picture = LoadPicture(App.Path & "\image\" & Int(coins / 10) & ".gif")
    If Int(coins / 10) = 0 Then
        Form1.Image4.Visible = False
    Else
        Form1.Image4.Visible = True
    End If
    
        If coins < Int(Form2.Text7.Text) Then
        Form4.Image3.Visible = True
        Command3.Enabled = False
        Form4.Image2.Picture = LoadPicture(App.Path & "\Image\insertcoins.bmp")
        Form4.Image3.Picture = LoadPicture(App.Path & "\Image\" & Text7.Text - coins & ".bmp")
    Else
            Form4.Image3.Visible = False
    End If

End Sub



Private Sub Command3_Click()

    Form1.Timer1.Enabled = True
    
    youxizhong = True
    Form4.Top = -99999
    coins = coins - Text7.Text
    Form1.Image3.Picture = LoadPicture(App.Path & "\Image\" & (coins Mod 10) & ".gif")
    If Int(coins / 10) > 0 Then
    Form1.Image4.Picture = LoadPicture(App.Path & "\image\" & Int(coins / 10) & ".gif")
    Else
    Form1.Image4.Visible = False

    End If
    Command3.Enabled = False
    Text1.Enabled = False
    Form3.Timer1.Enabled = True

    Form3.timeLast = Int(Text1.Text)
    Form3.Image2.Picture = LoadPicture(App.Path & "\image\t" & Int(Form3.timeLast / 10) & ".bmp")
    Form3.Image3.Picture = LoadPicture(App.Path & "\image\t" & Form3.timeLast Mod 10 & ".bmp")
    Form3.timeLast = Int(Text1.Text - 1)
    
    Form3.Show
    
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command9.Enabled = True
    Command10.Enabled = True
    jiroSelect (jiroType)
End Sub

Private Sub Command4_Click()
    If jiaoti = 0 Or jiaoti = 2 Or jiaoti = 4 Or jiaoti = 6 Or jiaoti = 8 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "左右交替敲打鼓边10次开启魔王：" & jiaoti
    End If
    If jiaoti = 10 Or jiaoti = 12 Or jiaoti = 14 Or jiaoti = 16 Or jiaoti = 18 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "开启魔王！再敲打10次开启里谱：" & jiaoti
    End If
End Sub

Private Sub Command5_Click()
If Form3.Visible = False And Command3.Enabled = True Then
    Call Command3_Click
End If
End Sub
Private Sub Command6_Click()
If Form3.Visible = False And Command3.Enabled = True Then
    Call Command3_Click
End If
End Sub

Private Sub Command7_Click()
    If jiaoti = 1 Or jiaoti = 3 Or jiaoti = 5 Or jiaoti = 7 Or jiaoti = 9 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "左右交替敲打鼓边10次开启魔王：" & jiaoti
    End If
    
    If jiaoti = 10 Then
        Label13.Caption = "开启魔王！再敲打10次开启里谱"
        jiroType = 1
        jiroText = Text2(jiroType).Text
    End If
    
        If jiaoti = 11 Or jiaoti = 13 Or jiaoti = 15 Or jiaoti = 17 Or jiaoti = 19 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "开启魔王！再敲打10次开启里谱：" & jiaoti
    End If
    
        If jiaoti = 20 Then
        Label13.Caption = "开启里谱！"
        jiroType = 2
        jiroText = Text2(jiroType).Text
    End If
End Sub



Private Sub Command8_Click(index As Integer)
setAddress (index)
End Sub

Public Function setAddress(i As Integer)
    CommonDialog1.DialogTitle = "浏览"
    CommonDialog1.InitDir = ""
    CommonDialog1.Filter = "*.exe|*.exe;"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then Text3(i).Text = CommonDialog1.FileName
End Function

Private Sub Command9_Click()
youxizhong = False

gameover = App.Path & "\sound\gameover.wav"
Result = sndplaysound(gameover, 1)

Form5.Timer1.Enabled = True
Form5.v = 300

If coins >= Int(Text7.Text) Then
    Command3.Enabled = True
End If
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command10.Enabled = False
Form3.Visible = False
Form3.Timer1.Enabled = False

Command9.Enabled = False
End Sub

Private Sub Form_Unload(cancel As Integer)
End
End Sub


Private Sub Form_load()
youxizhong = False
readAddress
readJiroExe
RegHook
End Sub







Private Sub text7_keypress(keyascii As Integer)
If keyascii < 48 Or keyascii > 57 Then keyascii = 0
End Sub


Private Sub text1_keypress(keyascii As Integer)
If keyascii < 48 Or keyascii > 57 Then keyascii = 0
End Sub



Private Sub Text7_Change()

        If coins < Int(Form2.Text7.Text) Then
        Form4.Image3.Visible = True
        Command3.Enabled = False
        Form4.Image2.Picture = LoadPicture(App.Path & "\Image\insertcoins.bmp")
        Form4.Image3.Picture = LoadPicture(App.Path & "\Image\" & Text7.Text - coins & ".bmp")
    Else
            Form4.Image3.Visible = False
            If youxizhong = False Then
                Command3.Enabled = True
            End If
            Form4.Image2.Picture = LoadPicture(App.Path & "\Image\hittostart.bmp")

    End If

End Sub
