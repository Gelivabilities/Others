VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CF��ս��"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   1905
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command23 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   27
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command22 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command21 
      Caption         =   "�㽭����"
      Height          =   375
      Left            =   1920
      TabIndex        =   25
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command20 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command19 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command18 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      Caption         =   "���϶���"
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command16 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      Caption         =   "�㽭һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      Caption         =   "�Ϻ�����"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command13 
      Caption         =   "��������"
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      Caption         =   "�Ĵ�һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "����һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��΢һ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�Ϸ�����"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��ʼʹ��"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�Ϻ�һ��"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�㶫����"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�㶫����"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�㶫����"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�㶫һ��"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "ѡ�����ǰ���ȵ���ٵ�½"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   28
      Top             =   4560
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "�����뱾����������"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "���Ŵ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "QQ���룺"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=318"
End If
End Sub

Private Sub Command10_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=344"
End If
End Sub

Private Sub Command11_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=333"
End If
End Sub

Private Sub Command13_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=329"
End If
End Sub

Private Sub Command14_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=326"
End If
End Sub

Private Sub Command15_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=325"
End If
End Sub

Private Sub Command16_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=341"
End If
End Sub

Private Sub Command17_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=340"
End If
End Sub

Private Sub Command18_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=332"
End If
End Sub

Private Sub Command12_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=328"
End If
End Sub

Private Sub Command19_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=330"
End If
End Sub

Private Sub Command2_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=327"
End If
End Sub

Private Sub Command20_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=324"
End If
End Sub

Private Sub Command21_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=349"
End If
End Sub

Private Sub Command22_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=352"
End If
End Sub

Private Sub Command23_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=353"
End If
End Sub

Private Sub Command3_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=338"
End If
End Sub

Private Sub Command4_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=339"
End If
End Sub

Private Sub Command5_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=320"
End If
End Sub
Private Sub Command6_Click()
a = Text1.Text
If Text1 = "" Then
MsgBox "�����������QQ����"
Else
If Text2 = "" Then
MsgBox "����������������"
Else
If a * 8926 + 29257168 <> Text2.Text Then
MsgBox "������������������"
Else
Label2.Visible = True
Command1.Visible = True
Command2.Visible = True
Text2.Visible = False
Label4.Visible = False
Command6.Visible = False
Form1.Height = 5370
Form1.Width = 3960
If Dir("C:\CF��ս��\x.bat") <> "" Then
Shell "C:\CF��ս��\x.bat"
Else
MsgBox "���Ȱ�ѹ���ļ���ѹ��C��"
Unload Me
End If
End If
End If
End If
End Sub

Private Sub Command7_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=342"
End If
End Sub

Private Sub Command8_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=347"
End If
End Sub

Private Sub Command9_Click()
If Text1 = "" Then
MsgBox "������������鿴��QQ����"
Else
Dim a
a = Text1
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "http://clan.cf.qq.com/cgi-bin/cfclan/User_Standings.cgi?clanid=477501&uin=" & a & "&areaid=348"
End If
End Sub

Private Sub Form_Load()
MsgBox " �����ֻ��ѧϰ֮�࣬�����ַ�������˽������Ը���"
Label2.Visible = False
Command1.Visible = False
Command2.Visible = False
End Sub
