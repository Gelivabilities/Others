VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������"
   ClientHeight    =   4515
   ClientLeft      =   9435
   ClientTop       =   2790
   ClientWidth     =   5985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4200
      Top             =   2160
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   32
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   31
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   29
      Text            =   "taikojiro.exe"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   28
      Text            =   "taikojiro.exe"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "���"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   22
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "���"
      Height          =   375
      Index           =   1
      Left            =   1440
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
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "����λ�á����⡢��������Ҫȷ��ȫ����ȷ��"
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   5775
      Begin VB.TextBox Text4 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   27
         Text            =   "taikojiro.exe"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "����ȫ��"
         Height          =   1335
         Left            =   3840
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "����ȫ��"
         Height          =   1335
         Left            =   4800
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "���"
         Height          =   375
         Index           =   0
         Left            =   1320
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
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "��20��"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "��10��"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ĭ��"
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
      Caption         =   "��ʼ��Ϸ"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   120
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "3"
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ӱ�"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "ģ�����"
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   600
         TabIndex        =   34
         Text            =   "3"
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "������Ϸ"
         Enabled         =   0   'False
         Height          =   855
         Left            =   1200
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "0"
         Height          =   615
         Left            =   1680
         TabIndex        =   33
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label15 
         Caption         =   "ѡ����ʱ"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "��    ��"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1485
         Width           =   855
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "���ҽ����ô�ı�10�ο���ħ��"
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

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Dim jiaoti, jiroExeName(0 To 2) As Integer
Dim youxizhong As Boolean
Public jiroText As String
Public coins, jiroType As Integer
Public selectingSongs As Boolean

Sub key_down(ByVal code As Long)
    Select Case code
        Case 71
            Call Command1_Click
        Case 68
            Call Command4_Click
        Case 70
            Call Command5_Click
        Case 69
            If Command9.Enabled = True Then
                Call Command9_Click
            End If
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

Public Function readJiroText()
On Error Resume Next

Open App.Path & "\jirotext.txt" For Input As #1

For i = 0 To 2
Line Input #1, s
Text2(i).Text = s
Next

Close #1

End Function

Public Function readJiroExe()
On Error Resume Next

Open App.Path & "\jiroexe.txt" For Input As #1

For i = 0 To 2
Line Input #1, s
Text4(i).Text = s
Next

Close #1

End Function
 
Public Function jiroSelect(i As Integer)
On Error Resume Next
    Shell Text3(i).Text, vbNormalFocus
End Function






Public Function addCoins()
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
    
End Function



Private Sub Command1_Click()
addCoins
End Sub

Private Sub Command11_Click()
On Error Resume Next
For i = 0 To 2
jiroSelect (i)
Next

End Sub

Public Function save()
Open App.Path & "\address.txt" For Output As #1
For i = 0 To 2
Print #1, Text3(i).Text
Next
Close #1

Open App.Path & "\jirotxt.txt" For Output As #1
For i = 0 To 2
Print #1, Text2(i).Text
Next
Close #1

Open App.Path & "\jiroexe.txt" For Output As #1
For i = 0 To 2
Print #1, Text4(i).Text
Next
Close #1
End Function

Public Function removeCoins()
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

End Function

Public Function stgEnabled()
    Form1.Timer1.Enabled = True
    Form2.Timer1.Enabled = True
    Command9.Enabled = True
    Form3.Timer1.Enabled = True
End Function

Public Function stgDisabled()
    Command3.Enabled = False
    Text1.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False

    For i = 0 To 2
        Text2(i).Enabled = False
        Text4(i).Enabled = False
    Next
End Function

Public Function startReadTime()
    Form3.Image3.Left = 1560
    Form3.timeLast = Int(Text1.Text)
    Form3.Image2.Picture = LoadPicture(App.Path & "\image\t" & Int(Form3.timeLast / 10) & ".bmp")
    Form3.Image3.Picture = LoadPicture(App.Path & "\image\t" & Form3.timeLast Mod 10 & ".bmp")
    Form3.timeLast = Int(Text1.Text - 1)
    Form3.Show
    Form3.Top = 405
    Form3.Image2.Visible = True
End Function

Public Function stgRemoveCoins()
    coins = coins - Text7.Text
    Form1.Image3.Picture = LoadPicture(App.Path & "\Image\" & (coins Mod 10) & ".gif")
    If Int(coins / 10) > 0 Then
    Form1.Image4.Picture = LoadPicture(App.Path & "\image\" & Int(coins / 10) & ".gif")
    Else
    Form1.Image4.Visible = False
    End If
End Function
Public Function stgVisible()
    Form4.Top = -99999
    Form5.Top = -99999
End Function

Public Function startGame()
    jiroSelect (jiroType)
    stgDisabled
    stgEnabled
    startReadTime
    youxizhong = True
    stgRemoveCoins
    stgVisible
End Function

Private Sub Command12_Click()
save
End Sub

Private Sub Command2_Click()
removeCoins
End Sub

Private Sub Command3_Click()
startGame
End Sub

Public Function lKaBeforeStart()
    If jiaoti = 0 Or jiaoti = 2 Or jiaoti = 4 Or jiaoti = 6 Or jiaoti = 8 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "���ҽ����ô�ı�10�ο���ħ����" & jiaoti
    End If
    If jiaoti = 10 Or jiaoti = 12 Or jiaoti = 14 Or jiaoti = 16 Or jiaoti = 18 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "����ħ�������ô�10�ο������ף�" & jiaoti
    End If
End Function

Private Sub Command4_Click()
lKaBeforeStart
End Sub

Public Function dongBeforeStart()
If Form3.Visible = False And Command3.Enabled = True Then
    Call Command3_Click
End If
End Function

Private Sub Command5_Click()
dongBeforeStart
End Sub

Private Sub Command6_Click()
dongBeforeStart
End Sub

Public Function rKaBeforeStart()
    If jiaoti = 1 Or jiaoti = 3 Or jiaoti = 5 Or jiaoti = 7 Or jiaoti = 9 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "���ҽ����ô�ı�10�ο���ħ����" & jiaoti
    End If
    
    If jiaoti = 10 Then
        Label13.Caption = "����ħ�������ô�10�ο�������"
        jiroType = 1
        jiroText = Text2(jiroType).Text
    End If
    
        If jiaoti = 11 Or jiaoti = 13 Or jiaoti = 15 Or jiaoti = 17 Or jiaoti = 19 Then
        jiaoti = jiaoti + 1
        Label13.Caption = "����ħ�������ô�10�ο������ף�" & jiaoti
    End If
    
        If jiaoti = 20 Then
        Label13.Caption = "�������ף�"
        jiroType = 2
        jiroText = Text2(jiroType).Text
    End If
End Function



Private Sub Command7_Click()
rKaBeforeStart
End Sub

Private Sub Command8_Click(index As Integer)
setAddress (index)
End Sub

Public Function setAddress(i As Integer)
    CommonDialog1.DialogTitle = "���"
    CommonDialog1.InitDir = ""
    CommonDialog1.Filter = "*.exe|*.exe;"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then Text3(i).Text = CommonDialog1.FileName
End Function

Public Function gameEnd()
youxizhong = False

Form5.Top = 420
gameOver = App.Path & "\sound\gameover.wav"
Result = sndplaysound(gameOver, 1)

Form5.Timer1.Enabled = True
Form5.v = 300



Form3.Visible = False
Form3.Timer1.Enabled = False

Command9.Enabled = False

Form2.Timer1.Enabled = False
Form1.Timer2.Enabled = False
End Function

Private Sub Command9_Click()
gameEnd
End Sub

Private Sub Form_Unload(cancel As Integer)
Shell "c:\windows\explorer.exe", vbMaximizedFocus
End
End Sub


Private Sub Form_load()
youxizhong = False
readAddress
readJiroText
readJiroExe
RegHook
jiroText = Text2(0).Text
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

Private Sub Timer1_Timer()
On Error Resume Next
If Form3.timeLast <= Int(Text1.Text) - 2 Then
    hwndSrc = Form1.handle
    hSrcDC = GetDC(hwndSrc) '��������߿��ȫ���� ���� GetWindowDC
    
    a = GetPixel(hSrcDC, 30, 500)

    b = GetPixel(hSrcDC, 30, 400)
    ReleaseDC hwndSrc, hSrcDC
    
    Label4.Caption = a & " " & b
    
    If a = 11887286 Then
        Form1.Timer2.Enabled = False
        Form3.Visible = False
    End If
    
    If b = 12709364 Then
        selectingSongs = True
    Else
        selectingSongs = False
    End If
End If
End Sub
