VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请登录"
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
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraUser 
      Caption         =   "选择身份"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton optUserType 
         Caption         =   "学生"
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optUserType 
         Caption         =   "教务管理人员"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmLogin 
      Caption         =   "登录"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3975
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消"
         Height          =   390
         Left            =   2880
         TabIndex        =   7
         Top             =   840
         Width           =   900
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定"
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
         Caption         =   "用户名："
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "口令："
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

'表示当前用户登录所选择的身份，即用户类型, 0-表示教务管理人员；1-表示学生
Dim mnUserType As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

      '显示MDI窗体, 并将用户类型和用户名传到MDI窗体中的mnUserType, msUserName中
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
