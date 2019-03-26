VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[太鼓の達人]iOS歌曲信息查询器"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List2 
      Height          =   5100
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只显示里谱"
      Height          =   300
      Left            =   2880
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "魔王"
      Height          =   1335
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   2895
      Begin VB.Label Label5 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "困难（松）"
      Height          =   1335
      Left            =   4200
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
      Begin VB.Label Label4 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "普通（竹）"
      Height          =   1335
      Left            =   4200
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
      Begin VB.Label Label3 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "简单（梅花）"
      Height          =   1335
      Left            =   4200
      TabIndex        =   2
      Top             =   4440
      Width           =   2895
      Begin VB.Label Label2 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Sub Check1_Click()

If Check1.Value = 1 Then '被选中，只显示里谱
List2.Visible = True '显示里谱的list框可见
Check1.Left = 1680  '调整位置，防止与谱面选择combobox有重叠
Combo2.Visible = True '谱面选择combobox
List2.ListIndex = 1 '同一列表框同一项被连续选中第2次以上时，是不作任何操作的，于是切换列表框时，选中的是一个list的歌曲，属性还是上一个list
List2.ListIndex = 0 '选中第一项，直接显示歌曲属性
List1.Visible = False '隐藏显示所有歌曲的list框，防止通过操作tab键接触到里面的东西
Else
If Check1.Value = 0 Then '取消选中，显示所有谱面
List1.Visible = True '显示所有歌曲的list框可见
List1.ListIndex = 1 '同上
List1.ListIndex = 0 '同上
List2.Visible = False '隐藏显示里谱的list框，防止通过操作tab键接触到里面的东西
Check1.Left = 2880 '移回原位，看起来没那么别扭
Combo2.ListIndex = 0 '选回表谱面，没有这条的话，如果刚才是里谱面，选其他没里谱的歌，魔王那里还是里谱，就不正确了
Combo2.Visible = False '既然所有歌曲列出框中每页第一个都是非里谱的歌，干脆就隐藏了
Else
End If
End If

End Sub


Private Sub Combo1_Click()

  Combo2.ListIndex = 0 '变表谱，原因跟check1相同
 Combo2.Visible = False '原因跟check1相同
 x = Combo1.ListIndex 'x=系列对应编号
 listthesongs (x) '根据系列对应编号，调用自动列歌曲函数
 List1.ListIndex = 0 '选中第一首歌，不然切换后原来那首歌属性还在会很别扭


End Sub

Private Sub Combo2_Click()
If List2.Visible = False Then '其实就是等价于没有勾选只显示里谱，list2不存在，下面代码改变list1的值
    If Combo2.ListIndex = 1 Then '选了里谱面
        Frame4.Caption = "魔王（里）" '魔王增加“里”，下一句读取里谱面
        Label5.Caption = "难度：★×" & ini.mfncGetFromIni(List1.Text, "lind", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List1.Text, "lilj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List1.Text, "litj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List1.Text, "licx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List1.Text, "ligc", App.Path & "\songdata.ini")
    Else '选了表谱面
         Frame4.Caption = "魔王" '魔王去掉“里”，下一句读取魔王（默认为表）
        Label5.Caption = "难度：★×" & ini.mfncGetFromIni(List1.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List1.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List1.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List1.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List1.Text, "mwgc", App.Path & "\songdata.ini")
    End If
Else
        If Combo2.ListIndex = 1 Then '等价于勾选了只显示里谱，下面代码改变list2的值，道理同上
        Frame4.Caption = "魔王（里）"
        Label5.Caption = "难度：★×" & ini.mfncGetFromIni(List2.Text, "lind", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List2.Text, "lilj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List2.Text, "litj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List2.Text, "licx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List2.Text, "ligc", App.Path & "\songdata.ini")
    Else
         Frame4.Caption = "魔王"
        Label5.Caption = "难度：★×" & ini.mfncGetFromIni(List2.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List2.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List2.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List2.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List2.Text, "mwgc", App.Path & "\songdata.ini")
    End If
End If
End Sub

Private Sub Form_Load()
'下面9条代码增加下拉框歌曲分类选项
Combo1.AddItem "全部", 0
Combo1.AddItem "Namco原创", 1
Combo1.AddItem "JPOP", 2
Combo1.AddItem "古典", 3
Combo1.AddItem "游戏", 4
Combo1.AddItem "动漫", 5
Combo1.AddItem "儿童", 6
Combo1.AddItem "V家", 7
Combo1.AddItem "综艺", 8

'下面两条代码使combo2有表里谱面选项
Combo2.AddItem "表谱面", 0
Combo2.AddItem "里谱面", 1

'两combobox均初始化为选择第一个
Combo1.ListIndex = 0
Combo2.ListIndex = 0

'列出全部歌曲，正好与上面Combo1.ListIndex = 0相对应
listthesongs (0)

'选第一首歌，原因见check1
List1.ListIndex = 0

'将此时已列出全部歌曲的list1中含里谱的歌曲导入到list2
        i = 1
        Do While i <= 259 '现在一共有这么多曲
        li = ini.mfncGetFromIni(List1.List(i), "li", App.Path & "\songdata.ini") '读取ini文件中的里谱信息，li=1表示该歌含里谱
        If li = "1" Then '判断出歌曲含里谱
        List2.AddItem List1.List(i) '添加里谱歌曲到list2
        Else
        End If
        i = i + 1
        Loop
        
If List2.ListCount <= 1 Or List1.ListCount <= 1 Then '由于check1要选到第二首歌（原因见check1），因此少于两个选项会导致出错，为避免出错，采用这种处理方式
MsgBox "读取出现错误。可能原因如下：" & vbCrLf & "1、歌曲列表或数据文件不存在" & vbCrLf & "2、上述信息文件存在异常" '提示出错
Unload Me '结束程序
End If

End Sub
Public Function listthesongs(x) As Integer '列出歌曲
List1.Clear '清空，不然只会继续重复添加歌曲
Select Case x 'x的值与combo1被选中的分类代号对应

    Case 0 '一个个分类列出所有歌曲
                i = 1
        Do While i <= 68
            List1.AddItem ini.mfncGetFromIni("Namco原创", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
         i = 1
        Do While i <= 82
            List1.AddItem ini.mfncGetFromIni("JPOP", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
        i = 1
        Do While i <= 11
            List1.AddItem ini.mfncGetFromIni("古典", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
      
        i = 1
        Do While i <= 24
            List1.AddItem ini.mfncGetFromIni("游戏", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 59
            List1.AddItem ini.mfncGetFromIni("动漫", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 2
            List1.AddItem ini.mfncGetFromIni("儿童", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 4
            List1.AddItem ini.mfncGetFromIni("V家", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 9
            List1.AddItem ini.mfncGetFromIni("综艺", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 1 '只列出该分类的歌曲
        i = 1
        Do While i <= 68
            List1.AddItem ini.mfncGetFromIni("Namco原创", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 2 '只列出该分类的歌曲
        i = 1
        Do While i <= 82
            List1.AddItem ini.mfncGetFromIni("JPOP", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 3 '只列出该分类的歌曲
        i = 1
        Do While i <= 11
            List1.AddItem ini.mfncGetFromIni("古典", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 4 '只列出该分类的歌曲
        i = 1
        Do While i <= 24
            List1.AddItem ini.mfncGetFromIni("游戏", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 5 '只列出该分类的歌曲
        i = 1
        Do While i <= 59
            List1.AddItem ini.mfncGetFromIni("动漫", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 6 '只列出该分类的歌曲
        i = 1
        Do While i <= 2
            List1.AddItem ini.mfncGetFromIni("儿童", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 7 '只列出该分类的歌曲
        i = 1
        Do While i <= 4
            List1.AddItem ini.mfncGetFromIni("V家", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 8 '只列出该分类的歌曲
        i = 1
        Do While i <= 9
            List1.AddItem ini.mfncGetFromIni("综艺", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
 End Select
End Function

Private Sub List1_Click()
    Combo2.ListIndex = 0 '变表谱，原因同check1
    
    If List1.Text = "" Then '缺少歌曲文件时读取出来的歌曲名会出现空白，如果空白则歌曲信息皆为空
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Else '读取歌曲名成功，下面代码从另一文件读取歌曲属性
    Label2.Caption = "难度：★×" & ini.mfncGetFromIni(List1.Text, "jdnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List1.Text, "jdlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List1.Text, "jdtj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List1.Text, "jdcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List1.Text, "jdgc", App.Path & "\songdata.ini")
    Label3.Caption = "难度：★×" & ini.mfncGetFromIni(List1.Text, "ptnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List1.Text, "ptlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List1.Text, "pttj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List1.Text, "ptcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List1.Text, "ptgc", App.Path & "\songdata.ini")
    Label4.Caption = "难度：★×" & ini.mfncGetFromIni(List1.Text, "knnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List1.Text, "knlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List1.Text, "kntj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List1.Text, "kncx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List1.Text, "kngc", App.Path & "\songdata.ini")
    Label5.Caption = "难度：★×" & ini.mfncGetFromIni(List1.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List1.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List1.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List1.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List1.Text, "mwgc", App.Path & "\songdata.ini")

    sc = ini.mfncGetFromIni(List1.Text, "sc", App.Path & "\songdata.ini") '读取俗称
    End If
    If sc = "" Then '没俗称，下面代码显示属性，但不显示俗称
        Label1.Caption = "歌曲名称：" & List1.Text & vbCrLf & "BPM：" & ini.mfncGetFromIni(List1.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf
    Else '有俗称，下面代码显示俗称和其他属性
        Label1.Caption = "歌曲名称：" & List1.Text & vbCrLf & "BPM：" & ini.mfncGetFromIni(List1.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf & "俗称：" & ini.mfncGetFromIni(List1.Text, "sc", App.Path & "\songdata.ini")
    End If
    
'原理在form1和check1中有
    li = ini.mfncGetFromIni(List1.Text, "li", App.Path & "\songdata.ini")
    If li = "1" Then
        Combo2.Visible = True
        Check1.Left = 1680
    Else
    Combo2.Visible = False
    Check1.Left = 2880
    End If
End Sub

Private Sub List2_Click() '原理同list1
    Combo2.ListIndex = 0
    Label2.Caption = "难度：★×" & ini.mfncGetFromIni(List2.Text, "jdnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List2.Text, "jdlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List2.Text, "jdtj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List2.Text, "jdcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List2.Text, "jdgc", App.Path & "\songdata.ini")
    Label3.Caption = "难度：★×" & ini.mfncGetFromIni(List2.Text, "ptnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List2.Text, "ptlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List2.Text, "pttj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List2.Text, "ptcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List2.Text, "ptgc", App.Path & "\songdata.ini")
    Label4.Caption = "难度：★×" & ini.mfncGetFromIni(List2.Text, "knnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List2.Text, "knlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List2.Text, "kntj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List2.Text, "kncx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List2.Text, "kngc", App.Path & "\songdata.ini")
    Label5.Caption = "难度：★×" & ini.mfncGetFromIni(List2.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "最大连击数：" & ini.mfncGetFromIni(List2.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "天井：" & ini.mfncGetFromIni(List2.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "初项：" & ini.mfncGetFromIni(List2.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "公差：" & ini.mfncGetFromIni(List2.Text, "mwgc", App.Path & "\songdata.ini")

    sc = ini.mfncGetFromIni(List2.Text, "sc", App.Path & "\songdata.ini")
    If sc = "" Then
        Label1.Caption = "歌曲名称：" & List2.Text & vbCrLf & "BPM：" & ini.mfncGetFromIni(List2.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf
    Else
        Label1.Caption = "歌曲名称：" & List2.Text & vbCrLf & "BPM：" & ini.mfncGetFromIni(List2.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf & "俗称：" & ini.mfncGetFromIni(List2.Text, "sc", App.Path & "\songdata.ini")
    End If
End Sub


