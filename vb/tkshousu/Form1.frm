VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "按键频率测试"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":103E
   ScaleHeight     =   6480
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   10440
      Top             =   5400
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   4080
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "相当于bpm为   的16分音符"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   17.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "最终成绩：     打/s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "成功：  次，共   次"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "当前速率：     打/s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   375
      Left            =   9960
      TabIndex        =   17
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "左手起"
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "74"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "70"
      Height          =   495
      Left            =   8280
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   8760
      TabIndex        =   11
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "F"
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   8640
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "想测一下你的手速吗？请从左边选择要测试的类型"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      DrawMode        =   9  'Not Mask Pen
      Height          =   2055
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "拇指模式"
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0C0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   615
      Index           =   1
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0C0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   615
      Index           =   0
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      DrawMode        =   9  'Not Mask Pen
      Height          =   2295
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long)

Private Sub Check1_Click()
Command1.SetFocus
Select Case Check1.Value
Case 0
Label2.Caption = Replace(Label2.Caption, "`", "F")
Label2.Caption = Replace(Label2.Caption, "-", "J")
Label9.Caption = 70
Label10.Caption = 74
Case 1
Label2.Caption = Replace(Label2.Caption, "F", "`")
Label2.Caption = Replace(Label2.Caption, "J", "-")
Label9.Caption = 192
Label10.Caption = 109
End Select
End Sub

Private Sub Check2_Click()
Command1.SetFocus
Select Case Check2.Value
Case 0
Label2.Caption = Replace(Label2.Caption, "左", "右")
Label7.Caption = "F"
Case 1
Label2.Caption = Replace(Label2.Caption, "右", "左")
Label7.Caption = "J"
End Select
End Sub

Private Sub Form_Load()
Label1(0).Caption = vbCrLf & "瞬间手速测试"
Label1(1).Caption = vbCrLf & "持续手速测试"
    k = GetTickCount()
 Label3.Caption = k

End Sub

Private Sub Label1_Click(index As Integer)
Label2.Alignment = 0
Label9.Caption = 70
Label10.Caption = 74
Label12.Caption = 0
Label16.Caption = 0
Label18.Visible = False
Label2.Font.Size = 24
Check1.Visible = True
Check1.Value = 0
Label4.Caption = 1
Check2.Visible = True
Check2.Value = 0
Label1(2).Visible = True
Label1(3).Visible = True
Label1(1).Top = 2760
Shape1(1).Top = 2760
Label1(0).Top = 2040
Shape1(0).Top = 2040
Command1.SetFocus
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label11.Visible = False
Label5.Visible = False
Select Case index
Case 0
Label5.Caption = 3
Label2.Caption = "请以最快速度交替按下F和J，右手开始，3次按键完成测试"


Case 1
Label2.Caption = "请以最快速度交替按下F和J，右手开始，100次按键完成测试"

Label5.Caption = 100




End Select
End Sub


Private Sub Timer1_Timer()
Label14.Caption = "当前速率：     打/s"
Label15.Caption = "成功：  次，共   次"
Label2.Font.Size = 45

Dim k As Long
    k = GetTickCount()

 Label2.Caption = Format((k - Label3.Caption - Label8.Caption) / 1000, "0.000")
 If Label2.Caption <> 0 Then
    Label11.Caption = Format(Label6.Caption / Label2.Caption, "0.00")
    
    Label18.Caption = "相当于BPM为" & Format(Label11.Caption, "0") * 15 & "的16分音符"
    Else
    End If

  If Label12.Caption = 5 Then
          Label2.FontSize = 24
        Label2.Caption = "5次违反按键规则，测试结束"
        End If
End Sub
Private Function sjss()
Timer1.Interval = 1
End Function

Private Sub command1_keydown(keycode As Integer, shift As Integer)



If Timer1.Interval = 0 Then Label6.Caption = 0

If keycode = Label9.Caption Then
    If Label4.Caption = "1" Then
     If Label7.Caption = "J" Then
        Label17.Top = 2400
        Label11.Top = 2400
        Label5.Visible = True
        Label11.Visible = True
        Label13.Visible = True
        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label6.Visible = True

        Check1.Visible = False
        Check2.Visible = False
        Label1(0).Visible = False
        Label1(1).Visible = False
        Label1(2).Visible = False
        Label1(3).Visible = False
        Shape1(0).Visible = False
        Shape1(1).Visible = False
        
        If Label6.Caption < Label5.Caption - 1 Then
            Label6.Caption = Label6.Caption + 1
            If Label6.Caption = 1 Then
            Label8.Caption = GetTickCount() - Label3.Caption
            Else
            End If
 
            Timer1.Interval = 1
            Label7.Caption = "F"
            
            Label2.Alignment = 1
            Else
            Timer1.Interval = 0
            Label9.Caption = 0
            Label10.Caption = 0
            

            Label1(0).Visible = True
            Label1(1).Visible = True
            Shape1(0).Visible = True
            Shape1(1).Visible = True
            
            Label1(0).Top = 1440
            Label1(1).Top = 2640
            Shape1(0).Top = 1440
            Shape1(1).Top = 2640

            Label14.Visible = False
            Label15.Visible = False
            Label16.Visible = False
            Label17.Visible = True
            Label11.Visible = True
            Label5.Visible = False

            Label18.Visible = True
            End If
        Else
        If Label6.Caption <> 0 Then Label12.Caption = Label12.Caption + 1
        
        If Label12.Caption = 5 Then
        Label2.FontSize = 24
        Label2.Alignment = 0
        Label2.Caption = "5次违反按键规则，测试结束"
        Timer1.Interval = 0

            Label9.Caption = 0
            Label10.Caption = 0
            

            Label1(0).Visible = True
            Label1(1).Visible = True
            Shape1(0).Visible = True
            Shape1(1).Visible = True
            
            Label1(0).Top = 1440
            Label1(1).Top = 2640
            Shape1(0).Top = 1440
            Shape1(1).Top = 2640
            
            Label13.Visible = False
            Label14.Visible = False
            Label15.Visible = False
            Label16.Visible = False
            Label17.Visible = False
            Label5.Visible = False
            Label11.Visible = False
        End If
        End If
    Else
    End If
Else
End If

If keycode = Label10.Caption Then

    If Label4.Caption = 1 Then
     If Label7.Caption = "F" Then
        Label17.Top = 2400
        Label11.Top = 2400
        Label5.Visible = True
        Label11.Visible = True
        Label13.Visible = True
        Label14.Visible = True
        Label15.Visible = True
        Label16.Visible = True
        Label6.Visible = True
        Check1.Visible = False
        Check2.Visible = False
        Label1(0).Visible = False
        Label1(1).Visible = False
        Label1(2).Visible = False
        Label1(3).Visible = False
        Shape1(0).Visible = False
        Shape1(1).Visible = False
        If Label6.Caption < Label5.Caption - 1 Then
            Label6.Caption = Label6.Caption + 1
             If Label6.Caption = 1 Then
            Label8.Caption = GetTickCount() - Label3.Caption
            Else
            End If
            Timer1.Interval = 1
            
            Label2.Alignment = 1
            
            Label7.Caption = "J"
            
            
            Else
            Timer1.Interval = 0
            Label9.Caption = 0
            Label10.Caption = 0
                        

            Label1(0).Visible = True
            Label1(1).Visible = True
            Shape1(0).Visible = True
            Shape1(1).Visible = True
            
            Label17.Visible = True
            Label14.Visible = False
            
            Label15.Visible = False
            Label16.Visible = False
            Label17.Visible = True
            Label11.Visible = True
            Label5.Visible = False
            
            Label1(0).Top = 1440
            Label1(1).Top = 2640
            Shape1(0).Top = 1440
            Shape1(1).Top = 2640
            
            Label18.Visible = True
            End If
        Else
        If Label6.Caption <> 0 Then Label12.Caption = Label12.Caption + 1
        
        If Label12.Caption = 5 Then
        Label2.FontSize = 24
        Label2.Alignment = 0
        Label2.Caption = "5次违反按键规则，测试结束"
        Timer1.Interval = 0

            Label9.Caption = 0
            Label10.Caption = 0
            

            Label1(0).Visible = True
            Label1(1).Visible = True
            Shape1(0).Visible = True
            Shape1(1).Visible = True
            
            Label1(0).Top = 1440
            Label1(1).Top = 2640
            Shape1(0).Top = 1440
            Shape1(1).Top = 2640
            
            Label13.Visible = False
            Label14.Visible = False
            Label15.Visible = False
            Label16.Visible = False
            Label17.Visible = False
            Label5.Visible = False
            Label11.Visible = False
        End If
        End If
    Else
    End If
Else
End If


End Sub

Private Sub Timer2_Timer()
If Timer1.Interval <> 0 Then
Label16.Caption = Label6.Caption
Else
If Label9.Caption = 0 Then
Label16.Caption = Label6.Caption + 1
End If
End If
End Sub
