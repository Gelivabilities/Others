VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Post模拟提交工具"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   10305
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton 开始 
      Caption         =   "开始"
      Height          =   280
      Left            =   9240
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton 清空Cookie 
      Caption         =   "清空Cookie"
      Height          =   280
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton 时间戳 
      Caption         =   "时间戳"
      Height          =   280
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "数据包"
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   10095
      Begin VB.TextBox 数据包 
         Height          =   5895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "协议头"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   10095
      Begin VB.TextBox 协议头 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame5 
         Caption         =   "提交地址"
         Height          =   650
         Left            =   120
         TabIndex        =   8
         Top             =   200
         Width           =   9855
         Begin VB.ComboBox 提交方式 
            Height          =   300
            ItemData        =   "Form4.frx":030A
            Left            =   8760
            List            =   "Form4.frx":0314
            TabIndex        =   11
            Text            =   "Get"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox 编码方式 
            Height          =   300
            ItemData        =   "Form4.frx":0323
            Left            =   7680
            List            =   "Form4.frx":0330
            TabIndex        =   10
            Text            =   "UTF-8"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox POST_GET_URL 
            Height          =   300
            Left            =   120
            TabIndex        =   9
            Text            =   "http://www.meilishuo.com/"
            Top             =   240
            Width           =   7455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "POST内容"
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   9855
         Begin VB.TextBox POST_GET_DATE 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   9615
         End
      End
   End
   Begin VB.Menu sjc 
      Caption         =   "时间戳"
      Visible         =   0   'False
      Begin VB.Menu 生成13位时间戳 
         Caption         =   "生成13位时间戳"
      End
      Begin VB.Menu 生成10位时间戳 
         Caption         =   "生成10位时间戳"
      End
      Begin VB.Menu 取消 
         Caption         =   "取消"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 开始_Click()
    POST_GET_URL.Text = Replace(POST_GET_URL.Text, "【时间戳10位】", 时间戳B())
    POST_GET_URL.Text = Replace(POST_GET_URL.Text, "【时间戳13位】", 时间戳A())
    POST_GET_DATE.Text = Replace(POST_GET_DATE.Text, "【时间戳10位】", 时间戳B())
    POST_GET_DATE.Text = Replace(POST_GET_DATE.Text, "【时间戳13位】", 时间戳A())
    If 提交方式.Text = "GET" Then
        数据包.Text = GetData(POST_GET_URL.Text, 编码方式.Text)
    Else
        数据包.Text = PostData(POST_GET_URL.Text, POST_GET_DATE.Text, 编码方式.Text)
    End If
End Sub

Private Sub 清空Cookie_Click()
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351", vbMaximizedFocus
End Sub

Private Sub 生成10位时间戳_Click()
    POST_GET_DATE.SetFocus
    SendKeys "【时间戳10位】"
End Sub

Private Sub 生成13位时间戳_Click()
    POST_GET_DATE.SetFocus
    SendKeys "【时间戳13位】"
End Sub

Private Sub 时间戳_Click()
    Me.PopupMenu sjc, , 时间戳.Left, 时间戳.Top + 时间戳.Height '显示时间戳菜单
End Sub
