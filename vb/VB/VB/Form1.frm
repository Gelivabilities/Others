VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "快捷键"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4455
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   3840
   End
   Begin VB.CommandButton Command5 
      Caption         =   "QQ2008文字转换ALT代码软件"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "打开本机ie首页"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "任务管理器"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CMD命令提示符"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算圆面积"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   4215
      Begin VB.Label Label1 
         Caption         =   "制作：LYC QQ：754571662       本程序用VB制作"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "Form1.frx":08CA
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "系统时间："
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "计算圆面积.exe", vbNormalFocus
End Sub

Private Sub Command2_Click()
Shell "cmd", vbNormalFocus
End Sub

Private Sub Command3_Click()
Shell "taskmgr", vbNormalFocus
End Sub

Private Sub Command4_Click()
Shell "C:\Program Files\Internet Explorer\iexplore.exe", vbNormalFocus
End Sub

Private Sub Command5_Click()
Shell "QQ文字转ALT代码—BY ▄︻︼-—一754571662.exe", vbNormalFocus
End Sub
Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub

