VERSION 5.00
Begin VB.Form FrmHelp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "帮助与支持-来自ICEE最权威的解析"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9240
      Picture         =   "FrmHelp.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   16
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9240
      Picture         =   "FrmHelp.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9240
      Picture         =   "FrmHelp.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   14
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   120
      ScaleHeight     =   561
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox PKB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00241D0A&
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   360
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   13
         Top             =   120
         Width           =   900
      End
      Begin ICEE.ICEE_TEXT HOT 
         Height          =   6495
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11456
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电影吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   7680
         TabIndex        =   26
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "美剧吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   6960
         TabIndex        =   25
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重口味吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   6000
         TabIndex        =   24
         Top             =   8040
         Width           =   720
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "内涵吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   5280
         TabIndex        =   23
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "肩上的脚丫吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   4080
         TabIndex        =   22
         Top             =   8040
         Width           =   1080
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "恐怖吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   3360
         TabIndex        =   21
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PS吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   2760
         TabIndex        =   20
         Top             =   8040
         Width           =   360
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   19
         Top             =   8040
         Width           =   360
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "小米手机贴吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   18
         Top             =   8040
         Width           =   1080
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "官方贴吧"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   8040
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   136
         X2              =   342
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Question and Answers"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   0
         Left            =   2040
         TabIndex        =   12
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "问题名称"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   120
         Width           =   840
      End
      Begin VB.Shape SB 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   0
         Top             =   7800
         Width           =   9735
      End
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   975
      Index           =   6
      Left            =   3480
      TabIndex        =   6
      Top             =   6600
      Width           =   2895
      _ExtentX        =   5953
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1095
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   7680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1931
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3413
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   3615
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6376
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1335
      Index           =   2
      Left            =   6480
      TabIndex        =   2
      Top             =   5160
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1935
      Index           =   3
      Left            =   6480
      TabIndex        =   3
      Top             =   3120
      Width           =   3375
      _ExtentX        =   4471
      _ExtentY        =   1296
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   2175
      Index           =   5
      Left            =   6480
      TabIndex        =   5
      Top             =   6600
      Width           =   3375
      _ExtentX        =   6800
      _ExtentY        =   1296
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1335
      Index           =   7
      Left            =   3480
      TabIndex        =   7
      Top             =   5160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2355
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1935
      Index           =   8
      Left            =   3480
      TabIndex        =   8
      Top             =   3120
      Width           =   2895
      _ExtentX        =   11245
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IS_MV As Boolean

Private Sub Form_Activate()
Me.BackColor = COLOR_NOR

Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is PictureBox Then
PBOX.Cls
PBOX.BackColor = Me.BackColor
End If
Next
Call PaintPng(App.Path & "\SKIN\H_T.PNG", Me.hdc, 8, 8)
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
HOT.SETBACKCOLOR Me.BackColor
HOT.SETFORECOLOR vbWhite
End Sub

Private Sub Form_Load()
For i = 0 To IHELP.Count - 1
IHELP(i).HASLINE = False
IHELP(i).HASTIP = False
Next
IS_MV = False
IHELP(0).SETTXT "播放器问题"
IHELP(1).SETTXT "涂鸦问题"
IHELP(2).SETTXT "文件下载问题"
IHELP(3).SETTXT "辅助类功能问题"
IHELP(4).SETTXT ""
IHELP(5).SETTXT "聊天类问题"
IHELP(6).SETTXT "文件传输类问题"
IHELP(7).SETTXT "UI界面问题"
IHELP(8).SETTXT "术语解释"
IHELP(1).SETCOLOR RGB(100, 28, 40), RGB(146, 19, 41)
IHELP(2).SETCOLOR RGB(170, 48, 63), RGB(203, 75, 75)
IHELP(3).SETCOLOR RGB(9, 43, 84), RGB(14, 83, 146)
IHELP(4).SETCOLOR &H25614B, &H2EBC7C
IHELP(5).SETCOLOR &H5B2989, &H563AB6
IHELP(6).SETCOLOR RGB(8, 70, 112), RGB(26, 109, 161)
IHELP(7).SETCOLOR RGB(67, 135, 148), RGB(77, 172, 190)
IHELP(0).SETCOLOR RGB(50, 28, 40), RGB(96, 19, 41)
IHELP(8).SETCOLOR vbBlack, COLOR_HIGH
If frmma.Left > Me.Width Then
Me.Move frmma.Left - Me.Width, frmma.Top
Else
Me.Move frmma.Left + frmma.Width, frmma.Top
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = True Then
IS_MV = False
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
End If

X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub
Private Sub IHELP_CLICK(Index As Integer)
Select Case Index
Case 0
HOT.SETTXT "Q:如何保存播放列表." & vbCrLf & _
"A:程序退出后会自动保存列表,您也可以在菜单选择导出列表" & vbCrLf & _
"Q:如何查找歌曲封面." & vbCrLf & _
"A:本程序提供了在线查找封面功能,由酷狗提供,您可以在封面单击右上角的搜索按钮,播放音乐窗的音乐暂时关闭搜索功能,您也可以将封面拖拽到封面区域手动设置" & vbCrLf & _
"Q:如何播放音乐." & vbCrLf & _
"A:您可以单击打开按钮或者在列表菜单中选择添加文件或文件夹,或者从音乐窗内打开喜欢的歌曲,拖拽文件至程序可见的部分也可以播放" & vbCrLf & _
"Q:音乐窗音乐加载有时失效." & vbCrLf & _
"A:您好,当前ICEE只是一个媒介,音乐库来自百度音乐链接,部分歌曲可能失效导致无法播放,有时端口被占用程序就无法加载列表,建议重新启动程序" & vbCrLf & _
"Q:播放顺序经常出错. " & vbCrLf & _
"A:当前版本播放顺序可能会存在错误,作者会在日后优化代码" & vbCrLf & _
"Q:如何分享音乐." & vbCrLf & _
"A:听到好的音乐时,总是迫不及待得想与好朋友分享,您只要到播放列表菜单选择 分享音乐 即可" & vbCrLf & _
"Q:播放英文歌曲时歌词秀经常出错." & vbCrLf & _
"A:您好,ICEE的歌词库来自中文歌词站,英文歌词肯定不是很全的,当然,也有一些歌曲是重名的,您可以选择 删除歌词 后在手动进行搜索" & vbCrLf & _
"Q:播放列表内的文件夹列表有什么用." & vbCrLf & _
"A:根据个人爱好,文件夹列表可以方便用户播放文件夹内所有歌曲 " & vbCrLf & _
"Q:如何更改歌手." & vbCrLf & _
"A:发现音乐文件歌手错误时,您可以单击音乐名称,出现一个封面进行修改" & vbCrLf & "Q:主窗体最小化任务栏会出现迷你播放器" & vbCrLf & _
"A:为了方便用户,如果音乐文件正在播放程序在最小化后会在任务栏出现迷你播放器,方便用户快速切歌" & vbCrLf & _
"Q:音乐文件的查找" & vbCrLf & _
"A:您好,目前版本音乐列表内查找歌曲还是没能实现的" & vbCrLf & _
"Q:物理删除文件." & vbCrLf & _
"A:所谓物理删除就是从磁盘上删除,删除后文件会被移除列表并移入回收站,您可以在播放列表选择 物理删除,或者播放器右上角的垃圾桶进行删除" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Case 1
HOT.SETTXT "Q:涂鸦是什么." & vbCrLf & "A:很多不了解涂鸦的人们会认为涂鸦就是乱涂乱画，其实不然. GRAFFITI 是一种视觉字体设计艺术，涂鸦内容包括很多:主要以变形英文字体为主，其次有3D写实.人物写实.各种场景写实 .卡通人物等等.配上艳丽的颜色让人产生强烈的视觉效果的和宣传效果" & vbCrLf & _
"Q:怎么打开涂鸦." & vbCrLf & "A:打开涂鸦的方式可以在主菜单打开 涂鸦画板 /个性相册中使用 涂鸦此图片/图像处理中文件预览选择菜单 涂鸦此图片,均可进入涂鸦界面" & vbCrLf & "Q:涂鸦画板的功能" & vbCrLf & _
"A:涂鸦画板分为画板区与工具栏区域,画板区域自然是绘画的,ICEE的画笔是羽化的,不是生硬的,这也更贴合涂鸦的本质" & vbCrLf & "Q:保存涂鸦" & vbCrLf & "A:工具栏提供保存涂鸦,打开图片,打印图片,分享图片四种功能,您也可以通过菜单保存涂鸦为BMP或者JPG格式,JPG失真率比较高,推荐使用BMP" & vbCrLf & _
"Q:涂鸦的画笔调整" & vbCrLf & "A:涂鸦的画笔可以改变粗细,在1至20之间选择,硬度可以在10至100中选择,适当调整硬度与粗细会使作品更细腻" & vbCrLf & _
"Q:对于鼠绘涂鸦有什么技巧." & vbCrLf & "A:好的作品需要丰富的颜色,您可以预先设置6个色域在下方的工具栏以方便调用,也可以通过色板进行颜色的选择,丰富的色彩有了,自然需要涂抹,不要嫌麻烦,一层一层颜色的累积会使画面更好" & vbCrLf & _
"Q:如何分享涂鸦." & vbCrLf & "画出一幅好画自然要显摆显摆,您需要登录ICEE方可分享,登陆成功后,单击分享即可选择要分享的好友" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "<本段结束>"
Case 2
HOT.SETTXT "Q:如何打开文件下载." & vbCrLf & "A:文件下载是指将文件从网络服务器拷贝至本地硬盘,需要连接互联网,您可以通过主界面右下角的下载与音乐播放器上方的下载进入" & vbCrLf & _
"Q:如何添加下载任务." & vbCrLf & "A:进入下载器后您会发现界面非常简单,您只需要单击添加任务即可弹出添加任务界面,输入地址,确定即可,目前支持大部分的网络文件(迅雷,快车,旋风的地址均可)" & vbCrLf & _
"Q:下载速度过慢怎么办." & vbCrLf & "A:您好,下载速度与网络速度有关,您可能有占用网络资源的进程在运行,打开任务管理器结束即可,也有可能是文件服务器过载或维护." & vbCrLf & _
"Q:下载文件地址的操作" & vbCrLf & "右键单击任务从菜单可以对任务进行复制,删除,停止,打开保存位置等操作" & vbCrLf & _
"Q:下载失败." & vbCrLf & "您好,这可能是由于网络文件失效导致,也可能是您未连接到互联网"
Case 3
HOT.SETTXT "Q:ICEE辅助功能都有哪些." & vbCrLf & "A:ICEE集成了多种辅助功能,包括屏幕软键盘/屏幕放大镜/截屏/桌面便签/系统资源监视/计算器/剪切板监视/文件搜索等功能" & vbCrLf & vbCrLf & _
 vbCrLf & "Q:如何打开放大镜." & vbCrLf & "A:您好,屏幕放大镜可以通过 主菜单-屏幕放大镜 进入,进入后主窗体会全部显示为屏幕内容,您可以对放大的比例进行调整" & vbCrLf & "Q:关于截屏的一些问题." & vbCrLf & _
"A:1.24版本在主窗体中加入了2个快捷按钮，其中有截屏按钮及便签的按钮,这样可以方便用户找到,您可以点击红色的按钮进入,如果您开启了快捷键,F8也是可以截屏,在个性相册底部的工具栏也有截屏的按钮." & vbCrLf & "Q:如何添加桌面便签." & vbCrLf & "A:①主界面-桌面便签 ②主菜单=新建便签 ③主界面黄色按钮" & vbCrLf & _
"Q:桌面便签有数量限制." & vbCrLf & "A:为了不让内存占用无限制,ICEE对桌面便签进行限制,您只可以建立10个便签,不过,对于用户正常的需求应该是合理的数字" & vbCrLf & "Q:关闭便签后内容会被保存吗." & vbCrLf & "A:您好,便签如果不是正常关闭的话内容是不会被保存的,便签的界面有两种关闭方式,[清除]是无痕迹关闭,[X]则是保留内容并关闭" & vbCrLf & _
"Q:ICEE为什么要对系统资源进行监视." & vbCrLf & "A:您好,ICEE的资源监视可以方便用户观察系统的变化,1.24的界面更直观,您可以查看CPU/内存/USB/虚拟内存的变化" & vbCrLf & "Q:计算器功能." & vbCrLf & "A:主界面-计算器进入,直接输入方程式进行运算即可" & vbCrLf & "Q:剪切板的监视." & vbCrLf & _
"A:剪切板可以获得图像文件与文本文件,图像文件可以保存,涂鸦,分享及打印,文本将会自动保存为文本文件,用户可以通过文本框下方的按钮查看有记录以来的所有文本." & vbCrLf & "Q:剪切板监视失效了怎么办." & vbCrLf & "A:通过 失效了点这里 即可恢复ICEE对剪切板的监视" & vbCrLf & _
"Q:如何搜索文件." & vbCrLf & "A:您好,本产品暂不支持文件名搜索,目前只搜索后缀文件.用户可以通过音乐播放器-播放列表-本地搜索 进入搜索模式,输入后缀即可搜索,搜索结果将会保存至Media文件夹下,下次运行后会自动加载搜索结果." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Case 4
HOT.SETTXT "还没想好写什么"
Case 5
HOT.SETTXT "Q:如何注册ICEE的账号." & vbCrLf & "A:ICEE的服务器目前不是很稳定,服务器的地址需要用户手动输入,也可以通过 设置 搜索主机,如果您是新用户,您可以输入ID及密码后将登陆界面的[以新用户登录]勾选登录即可,如果您不是新用户而将其勾选了,将会失败" & vbCrLf & "Q:如何添加好友." & vbCrLf & "A:登陆后在搜索好友框输入ID按回车即可" & vbCrLf & "Q:删除好友及屏蔽好友的常见问题" & vbCrLf & _
"A:右键单击好友列表,弹出菜单选择相关选项即可." & vbCrLf & "Q:与好友快速的聊天" & vbCrLf & "A:双击好友ID弹出泡泡文本框,输入内容即可发送,对方将会收到泡泡文字" & vbCrLf & "即时聊天与快速聊天的区别是什么." & vbCrLf & "快速聊天,顾名思义,是以纯文本的形式聊天,即时聊天时您可以有更多的功能选择,比如发送文件,发送表情,举报,远程协助等功能" & vbCrLf & _
"Q:发送文件的一些问题" & "A:您好,这段解答比较长,作者将会在[文件传输]类进行详细解答" & vbCrLf & "Q:如何更换密码." & vbCrLf & "在好友列表菜单选择修改密码即可进入子界面,密码长度不超过16位,修改成功后会进行消息框通知." & vbCrLf & _
""
Case 6
HOT.SETTXT "作者撰稿中"
Case 7
HOT.SETTXT "作者撰稿中"
Case 8
HOT.SETTXT "作者撰稿中"
End Select
PO.Visible = True
LA(2).Caption = IHELP(Index).MYTIT
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub LF_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Dim WILL_URL As String
Select Case Index
Case 0
WILL_URL = "http://tieba.baidu.com/f.kw=icee&fr=index"
Case 1
WILL_URL = "http://tieba.baidu.com/f.kw=%E5%B0%8F%E7%B1%B3&fr=index&fp=0&ie=utf-8"
Case 2
WILL_URL = "http://tieba.baidu.com/f.kw=vb&fr=index"
Case 3
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=PS"
Case 4
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E6%81%90%E6%80%96"
Case 5
WILL_URL = "http://tieba.baidu.com/f.kw=%BC%E7%C9%CF%B5%C4%BD%C5%D1%BE&fr=index"
Case 6
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E5%86%85%E6%B6%B5"
Case 7
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E9%87%8D%E5%8F%A3%E5%91%B3"
Case 8
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E7%BE%8E%E5%89%A7"
Case 9
WILL_URL = "http://tieba.baidu.com/f.kw=%B5%E7%D3%B0&fr=ala0"
End Select
ShellExecute 0&, vbNullString, WILL_URL, vbNullString, vbNullString, 0 '调用ie
End Sub

Private Sub LF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LF(Index).FOREColor <> &H30F1F1 Then LF(Index).FOREColor = &H30F1F1
End Sub
Private Sub PKB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = False Then
IS_MV = True
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PKB.hdc, 0, 0)
End If
End Sub

Private Sub PKB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PO.Visible = False
End Sub

Private Sub PO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = True Then
IS_MV = False
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
End If
X1.Visible = True
X2.Visible = False
X3.Visible = False
Dim i As Integer
For i = 0 To LF.Count - 1
If LF(i).FOREColor <> vbWhite Then LF(i).FOREColor = vbWhite
Next
End Sub

Private Sub x1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then Unload Me
End Sub
