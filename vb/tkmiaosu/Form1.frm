VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "̫�Ĥ��_��-����/���ټ���"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3015
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2775
      Begin VB.Label Label4 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "BPM"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Height = 3555
Dim s As Integer
For i = 1 To Len(Text1.Text)
If Mid(Text1.Text, i, Len(".")) = "." Then
s = s + 1
End If
Next
If s > 1 Then
Label4.Caption = "����Ϊ�Ҳ�֪��������" & s & "����ô"
Exit Sub
Else
If Text1.Text = "." Then
Label4.Caption = "����������ҳ���û��"
Else
If Text1.Text = "" Then
Label4.Caption = "��������д����"
Else
If Text2.Text = "" Then
Label4.Caption = "��������д����"
Else
If Text3.Text = "" Then
Label4.Caption = "��������д����"
Else
Dim bpm As Double
Dim b As Double
Dim n As Double
Dim f As Double
bpm = Text1.Text
x = Text2.Text
n = Text3.Text
f = bpm * x / 240
If bpm > 1000 Then
Label4.Caption = "��������ȷ����"
Else
If bpm < 15 Then
Label4.Caption = "��������ȷ����"
Else
If x > 512 Then
Label4.Caption = "��������ȷ����"
Else
If f > 60 Then
Label4.Caption = "��������ȷ����"
Else
If n < 2 Then
Label4.Caption = "��������ȷ����"
Else
If x = 0 Then
Label4.Caption = "��������ȷ����"
Else
Label4.Caption = "��������Ϊ" & Format(f, "0.00") & "��/s" & vbCrLf & "ȫ��������������=" & Format(vbCrLf & (n - 1) * f / (0.05 * f + n - 1), "0.00") & "��/s" & vbCrLf & "ȫ��������������=" & Format((n - 1) * f / (0.15 * f + n - 1), "0.00") & "��/s"
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Form_Load()
Command1.Caption = "���㡡"
End Sub

Private Sub text1_keypress(keyascii As Integer)
If keyascii = 46 And Not CBool(InStr(txbNumber, ".")) Then Exit Sub
If keyascii = 8 Then Exit Sub
If keyascii < 48 Or keyascii > 57 Then keyascii = 0
End Sub
Private Sub text2_keypress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If keyascii < 48 Or keyascii > 57 Then keyascii = 0
End Sub
Private Sub text3_keypress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If keyascii < 48 Or keyascii > 57 Then keyascii = 0
End Sub


