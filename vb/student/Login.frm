VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���¼"
   ClientHeight    =   2580
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1524.349
   ScaleMode       =   0  'User
   ScaleWidth      =   3985.824
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraUser 
      Caption         =   "ѡ�����"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton optUserType 
         Caption         =   "ѧ��"
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optUserType 
         Caption         =   "���������Ա"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmLogin 
      Caption         =   "��¼"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3975
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��"
         Height          =   390
         Left            =   2880
         TabIndex        =   7
         Top             =   840
         Width           =   900
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��"
         Default         =   -1  'True
         Height          =   390
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   900
      End
      Begin VB.TextBox txtPwd 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   885
         Width           =   1485
      End
      Begin VB.TextBox txtUser 
         Height          =   345
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "�û�����"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   3
         Top             =   915
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ʾ��ǰ�û���¼��ѡ�����ݣ����û�����, 0-��ʾ���������Ա��1-��ʾѧ��
Dim mnUserType As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

      '��ʾMDI����, �����û����ͺ��û�������MDI�����е�mnUserType, msUserName��
      Load MDIMain
      With MDIMain
        .mnUserType = mnUserType
        .msUserName = "436346"
        .Show
      End With
      Unload Me

End Sub

Private Sub Form_Load()
    optUserType(0).Value = True
End Sub

Private Sub optUserType_Click(Index As Integer)
    mnUserType = Index
End Sub
