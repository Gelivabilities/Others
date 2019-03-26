VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ScaleHeight     =   390
   ScaleWidth      =   4830
   Begin VB.CommandButton 搜索 
      Caption         =   "百度一下"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub 搜索_Click()
Dim ie, wd
wd = Text1
Set ie = CreateObject("internetexplorer.application")
ie.Visible = True
ie.navigate "http://www.google.cn/search?hl=zh-CN&q=" & wd
End Sub
