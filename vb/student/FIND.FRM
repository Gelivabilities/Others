VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2295
   ClientLeft      =   1995
   ClientTop       =   2595
   ClientWidth     =   4875
   HelpContextID   =   2016128
   Icon            =   "FIND.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1912.104
   ScaleMode       =   0  'User
   ScaleWidth      =   4797.773
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFields 
      Height          =   1680
      ItemData        =   "FIND.frx":000C
      Left            =   240
      List            =   "FIND.frx":000E
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox lstOperators 
      Height          =   1680
      ItemData        =   "FIND.frx":0010
      Left            =   2040
      List            =   "FIND.frx":0012
      TabIndex        =   1
      Top             =   360
      Width           =   960
   End
   Begin VB.TextBox txtExpression 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "运算符："
      Height          =   180
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "字段："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "值或表达式："
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public msFindField As String                '查找的字段
Public msFindOp As String                   '查找的运算符
Public msFindExpr As String                 '查找的表达式的值
'表示用户是否取消查询，如果取消查询则为就为True；否则为False
Public mbFindFailed As Boolean

Private Sub cmdCancel_Click()
    mbFindFailed = True
    Me.Hide
End Sub

Private Sub txtExpression_Change()
  cmdOK.Enabled = Len(lstFields.Text) > 0 And Len(lstOperators.Text) > 0 And Len(txtExpression.Text) > 0
End Sub

Private Sub lstFields_Click()
  cmdOK.Enabled = Len(lstFields.Text) > 0 And Len(lstOperators.Text) > 0 And Len(txtExpression.Text) > 0
End Sub

Private Sub lstOperators_Click()
  cmdOK.Enabled = Len(lstFields.Text) > 0 And Len(lstOperators.Text) > 0 And Len(txtExpression.Text) > 0
End Sub

Private Sub Form_Load()
   '加载查询所需要使用的运算符号
   lstOperators.AddItem "="
   lstOperators.AddItem "<>"
   lstOperators.AddItem ">="
   lstOperators.AddItem "<="
   lstOperators.AddItem ">"
   lstOperators.AddItem "<"
   lstOperators.AddItem "Like"
   lstOperators.ListIndex = 0
        
   mbFindFailed = True
End Sub

Private Sub cmdOK_Click()
   mbFindFailed = False
   
   '改变指针，告知读者当前处于忙的状态
   Screen.MousePointer = vbHourglass

    '取得查询所需要的字段、符号和值
   msFindField = lstFields.Text
   msFindExpr = txtExpression.Text
   msFindOp = lstOperators.Text

   Me.Hide
   
   '恢复指针，告知读者系统已经不忙了
   Screen.MousePointer = vbDefault
End Sub
