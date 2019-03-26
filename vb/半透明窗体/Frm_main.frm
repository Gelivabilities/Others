VERSION 5.00
Begin VB.Form Frm_main 
   Caption         =   "半透明窗体"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   Picture         =   "Frm_main.frx":0000
   ScaleHeight     =   4395
   ScaleWidth      =   7155
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "退      出"
      Height          =   495
      Left            =   5310
      TabIndex        =   1
      Top             =   3765
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置窗体半透明"
      Height          =   495
      Left            =   3525
      TabIndex        =   0
      Top             =   3765
      Width           =   1695
   End
End
Attribute VB_Name = "Frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Load Frm_Transpare
    Frm_Transpare.Show
    Frm_Transpare.Top = Me.Top + Me.Height
    
End Sub
'Download by http://www.codefans.net
Private Sub Command2_Click()
    End
End Sub
