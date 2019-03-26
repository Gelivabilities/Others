VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Transpare 
   Caption         =   "设置透明度"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4950
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.Slider Slider1 
      Height          =   570
      Left            =   15
      TabIndex        =   0
      Top             =   465
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_Transpare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Download by http://www.codefans.net
Private Sub Form_Load()
    Slider1.Min = 0
    Slider1.Max = 100
    Slider1.Value = 70
    Slider1_Change
End Sub

Private Sub Slider1_Change()
    transparence Frm_main, Slider1.Value * 0.01
End Sub

