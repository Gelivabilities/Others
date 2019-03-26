VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   3090
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   1320
   End
   Begin VB.TextBox Text5 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "生成的代码"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "下载地址"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "关键词"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "曲名"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Text5.Text = "<name>" & Text1.Text & "</name><keywords>" & Text2.Text & "</keywords><url>" & Text3.Text & "</url>"
End Sub
