VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CPU占用率检测"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2745
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4560
      Top             =   840
   End
   Begin VB.Label lblCPU 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'★★★★★****************************★★★★★**********************★★★★★
'金诺VB园-收藏整理
'本站是专注于VB和VBNET编程的源码下载站
'发布日期：2008-4-2 16:47:08
'网    站：http://www.vbget.com/          (金诺VB园)
'网    站：http://www.vbget.com/daohan/   (VB编程网址导航)
'E-Mail  ：vbget@yahoo.cn
'QQ      ：158676144
'源码作者：如果您有VB商业源码需要获得收益，本站将有VIP收费下载频道可供你发布!
'         您有权定价;改价;删除;及即时查看下载量(即收益)，所有收益全部归您！
'         本站将在双方协商的一个金额周期内打款到作者帐户中，您只需负责打款费用！
'         本站只作为一个平台提供最新VB源码咨讯和源码下载！
'本注释由<站长工具之智能注释>软件自动添加！金诺VB园有此软件下载！
'★★★★★****************************★★★★★**********************★★★★★

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private CPU As clsCPUUsage

Private Sub Form_Load()

    Set CPU = New clsCPUUsage

End Sub



Private Sub Timer1_Timer()

    lblCPU.Caption = "CPU占用率: " & CPU.Usage & "%"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set CPU = Nothing

End Sub
