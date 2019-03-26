VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB代码管家_第三版 By_5mao"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   10350
   ScaleWidth      =   16350
   StartUpPosition =   2  '屏幕中心
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2730
      ItemData        =   "Form1.frx":030A
      Left            =   3960
      List            =   "Form1.frx":030C
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton 皮肤 
      Caption         =   "皮肤"
      Height          =   270
      Left            =   14400
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton 附件 
      Caption         =   " 附件"
      Height          =   270
      Left            =   13440
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox 过滤内容 
      Height          =   270
      Left            =   960
      TabIndex        =   8
      ToolTipText     =   "请输入汉语拼音首字母进行查找操作"
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   9420
      ItemData        =   "Form1.frx":030E
      Left            =   120
      List            =   "Form1.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   3495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9975
      Width           =   16350
      _ExtentX        =   28840
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   15954
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton 工具箱 
      Caption         =   "工具箱"
      Height          =   270
      Left            =   15360
      TabIndex        =   5
      Top             =   115
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10680
      Top             =   0
   End
   Begin VB.CommandButton 删除 
      Caption         =   "删除"
      Height          =   270
      Left            =   5640
      TabIndex        =   4
      Top             =   115
      Width           =   855
   End
   Begin VB.CommandButton 更新 
      Caption         =   "更新"
      Height          =   270
      Left            =   3720
      TabIndex        =   3
      Top             =   115
      Width           =   855
   End
   Begin VB.CommandButton 添加 
      Caption         =   "添加"
      Height          =   270
      Left            =   4680
      TabIndex        =   2
      Top             =   115
      Width           =   855
   End
   Begin VB.TextBox 代码内容 
      Height          =   9060
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      ToolTipText     =   "此处为代码的内容"
      Top             =   840
      Width           =   12495
   End
   Begin VB.TextBox 代码标题 
      Height          =   270
      Left            =   3720
      TabIndex        =   0
      ToolTipText     =   "此处为代码的标题"
      Top             =   480
      Width           =   12495
   End
   Begin VB.Label Label1 
      Caption         =   "自动查找"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   170
      Width           =   735
   End
   Begin VB.Menu pf 
      Caption         =   "皮肤设置"
      Visible         =   0   'False
      Begin VB.Menu pf_xia 
         Caption         =   "下一组皮肤"
      End
      Begin VB.Menu pf_shang 
         Caption         =   "上一组皮肤"
      End
      Begin VB.Menu pf_qx 
         Caption         =   "取消"
      End
   End
   Begin VB.Menu fj 
      Caption         =   "附件"
      Visible         =   0   'False
      Begin VB.Menu drfj 
         Caption         =   "导入附件"
      End
      Begin VB.Menu dcfj 
         Caption         =   "导出附件"
      End
      Begin VB.Menu sqfj 
         Caption         =   "删除附件"
      End
      Begin VB.Menu fj_qx 
         Caption         =   "取消"
      End
   End
   Begin VB.Menu gjx 
      Caption         =   "工具箱"
      Visible         =   0   'False
      Begin VB.Menu Spy 
         Caption         =   "Spy++"
      End
      Begin VB.Menu 编码转换 
         Caption         =   "编码转换"
      End
      Begin VB.Menu Post模拟 
         Caption         =   "Post模拟"
      End
      Begin VB.Menu gjx_qx 
         Caption         =   "取消"
      End
   End
   Begin VB.Menu Tray 
      Caption         =   "托盘"
      Visible         =   0   'False
      Begin VB.Menu Display 
         Caption         =   "显示"
      End
      Begin VB.Menu exit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection '声明数据库对像
Dim rs As New ADODB.Recordset  '声明表对像
Dim 缓存标题 As String
Dim 标题数组() As String

Private Declare Function SkinH_SetAero Lib "SkinH.dll" (ByVal hwnd As Long) As Long
Private Declare Function SkinH_Attach Lib "SkinH.dll" () As Long
Private Declare Function SkinH_AttachEx Lib "SkinH.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
'-----------------------------she格式皮肤-----------------------------
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'------------------------------------声名API
Private nfIconData As NOTIFYICONDATA
Const MAX_TOOLTIP As Integer = 50    '提示字符串中预显示的个数
Const NIF_ICON = &H2                 '预添加的图标
Const NIF_MESSAGE = &H1              '事件消息,比如鼠标抬起或按下
Const NIF_TIP = &H4                  '预显示的文字
Const NIM_ADD = &H0                  '添加托盘图标
Const NIM_DELETE = &H2               '删除托盘图标
Const WM_MOUSEMOVE = &H200           '鼠标移动
Const WM_LBUTTONDOWN = &H201         '按下右键
Const WM_LBUTTONUP = &H202           '左键抬起
Const WM_LBUTTONDBLCLK = &H203       '左键双击
Const WM_RBUTTONDOWN = &H204         '按下右键
Const WM_RBUTTONUP = &H205           '右键抬起
Const WM_RBUTTONDBLCLK = &H206       '右键双击
Const SW_RESTORE = 9                 '状态恢复
Const SW_HIDE = 0                    '状态隐藏
'------------------------------------声名常量
Private Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type

Private Sub Display_Click() '显示窗体
    Me.WindowState = 0          '还原窗体
    Form1.Visible = True
    Form1.Show
End Sub

Private Sub dcfj_Click() '导出附件
    If StatusBar1.Panels(5).Text = "附件信息：无" Then '判断是否有附件存在
        MsgBox "该条代码附件不存在！", vbInformation, "VB代码管家"
    Else
        On Error GoTo ErrHandle          '用户取消时触发错误
        CommonDialog1.CancelError = True
        CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "请选择zip附件导出路径"
        CommonDialog1.Flags = &H80000
        CommonDialog1.Filter = "zip文件(*.zip) |*.zip"
        CommonDialog1.ShowSave
        If CommonDialog1.FileName <> "" Then
            Set mstream = New ADODB.Stream
            mstream.Type = adTypeBinary
            mstream.Open
            mstream.Position = 0
            mstream.Write rs.Fields("附件").Value
            mstream.SaveToFile CommonDialog1.FileName, adSaveCreateOverWrite
            MsgBox "附件导出成功！", vbInformation, "VB代码管家"
        End If
        Exit Sub
ErrHandle:                       '错误处理
        Select Case Err.Number
        Case 32755
            MsgBox "您未选择任何文件！", vbInformation, "VB代码管家"
        End Select
    End If
End Sub

Private Sub drfj_Click() '导入附件
    On Error GoTo ErrHandle          '用户取消时触发错误
    CommonDialog1.CancelError = True
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "请选择要导入的zip压缩文件"
    CommonDialog1.Flags = &H80000
    CommonDialog1.Filter = "zip文件(*.zip) |*.zip"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Set mstream = New ADODB.Stream
        mstream.Type = adTypeBinary
        mstream.Open
        mstream.LoadFromFile CommonDialog1.FileName
        rs.Fields("附件").Value = mstream.Read
        rs.Update
        Call 检查是否有附件存在
        MsgBox "附件导入成功！", vbInformation, "VB代码管家"
    End If
    Exit Sub
ErrHandle:                       '错误处理
    Select Case Err.Number
    Case 32755
        MsgBox "您未选择任何文件！", vbInformation, "VB代码管家"
    End Select
End Sub

Private Sub sqfj_Click() '删除附件
    If StatusBar1.Panels(5).Text = "附件信息：无" Then '判断是否有附件存在
        MsgBox "该条代码附件不存在！", vbInformation, "VB代码管家"
    Else
        '————————————————————
        Dim v As String
        v = MsgBox("您确定要删除标题为:“" & List1.List(List1.ListIndex) & "”的附件信息吗？", vbOKCancel, "温馨提示")
        If v = vbOK Then
            rs("附件") = ""
            rs.Update
            Call 检查是否有附件存在
            MsgBox "附件删除成功！", vbInformation, "VB代码管家"
        End If
        '————————————————————提示是否真的要删除附件
    End If
End Sub

Private Sub exit_Click() '退出
    Call 退出
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lMsg As Single
    lMsg = x / Screen.TwipsPerPixelX
    If lMsg = WM_RBUTTONUP Then '如果单击右键
        Me.PopupMenu Tray            '菜单显示在光标处
    End If
    
    If lMsg = WM_LBUTTONDBLCLK Then '如果左键双击
        Call Display_Click   '显示
    End If
End Sub '此事件中的代码只针对托盘上的图标
'---------------------------------------------------显示菜单

Private Sub 添加托盘图标()
    nfIconData.hwnd = Me.hwnd
    nfIconData.uID = Me.Icon
    nfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    nfIconData.uCallbackMessage = WM_MOUSEMOVE
    nfIconData.hIcon = Me.Icon.Handle
    nfIconData.szTip = "VB代码管家" & vbNullChar  'vbNullChar表示删除右边多于的空格
    nfIconData.cbSize = Len(nfIconData)
    
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)    '添加图标到托盘
End Sub

Private Sub Form_Load() '打开数据库
    '——————————————————————————————————
    If App.PrevInstance = True Then
        End
    End If
    '——————————————————————————————————禁止重复运行
    '——————————————————————————————————
    If Dir(App.Path & "\SkinH.dll") = "" Then
        MsgBox "皮肤Dll文件不存在，程序将自动退出！", vbInformation, "VB代码管家"
        End
    End If
    '——————————————————————————————————判断皮肤Dll
    '——————————————————————————————————
    If Dir(App.Path & "\皮肤", vbDirectory) = "" Then
        MsgBox "皮肤文件夹不存在，程序将自动退出！", vbInformation, "VB代码管家"
        End
    End If
    '——————————————————————————————————判断皮肤文件夹
    '——————————————————————————————————
    If Dir(App.Path & "\VB代码数据库.mdb") = "" Then
        MsgBox "数据库文件不存在，程序将自动退出！", vbInformation, "VB代码管家"
        End
    End If
    '——————————————————————————————————判断数据库
    代码内容.BackColor = RGB(238, 238, 238)
    代码标题.BackColor = RGB(238, 238, 238)
    过滤内容.BackColor = RGB(238, 238, 238)
    Call 添加托盘图标
    
    On Error GoTo ErrHandle                    '错误处理
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\VB代码数据库.mdb;Jet OLEDB:database password=admin" 'admin为数据库密码
    rs.Open "Select * From 代码库 ", db, 1, 3
ErrHandle:                                 '错误处理
    
    Select Case Err.Number
    Case -2147217843
        MsgBox "数据库密码错误，请将数据库密码修改为admin后再试！", vbExclamation, "错误提示"
    End Select
    
    If rs.State = adStateOpen Then '如果连接成功
        Call 导入代码列表事件
    End If
    
    '——————————————————————————————————————————
    List2.Clear
    Dim abc As String
    Dim genmulu As String
    genmulu = App.Path & "\皮肤\"                      '路径,记得路径后面一样要加"\"
    abc = Dir(genmulu, vbNormal)
    Do While abc <> ""
        If abc <> "." And abc <> ".." And Right(abc, 4) = ".she" Then
            List2.AddItem genmulu & abc
        End If
        abc = Dir                                      '再次调用dir函数,此时可以不带参数
    Loop
    
    If List2.ListCount > 0 Then '如果皮肤不为空
        List2.ListIndex = 0 '选中第一个皮肤
        '————————————————————————————
        If GetSetting(App.Title, "Settings", "pifu", "") <> "" Then '————————————————如果上一次设置过皮肤
            If Dir(GetSetting(App.Title, "Settings", "pifu", "")) <> "" Then '如果上一次设置的皮肤依然存在
                Dim n As Integer
                For n = 0 To List2.ListCount - 1
                    '—————————————————
                    If List2.List(n) = GetSetting(App.Title, "Settings", "pifu", "") Then
                        SkinH_AttachEx GetSetting(App.Title, "Settings", "pifu", ""), "" '加载上一次设置的皮肤
                        List2.ListIndex = n '选中第n个皮肤
                    End If
                    '—————————————————查找上次皮肤的路径
                Next n
            End If
        Else
            List2.ListIndex = 0 '选中第一个皮肤
        End If
        '————————————————————————————皮肤加载方式
    Else
        MsgBox "皮肤文件缺失，程序将自动退出！", vbInformation, "VB代码管家"
        End
    End If
    '——————————————————————————————————————————加载皮肤列表
End Sub

Private Sub pf_shang_Click() '切换到上一组皮肤
    If List2.ListIndex = 0 Then
        MsgBox "上一组皮肤为空，请切换到下一组！", vbInformation, "VB代码管家"
    Else
        List2.ListIndex = List2.ListIndex - 1
        SaveSetting App.Title, "Settings", "pifu", List2.List(List2.ListIndex) '保存选中皮肤
    End If
End Sub

Private Sub pf_xia_Click()   '切换到下一组皮肤
    If List2.ListIndex = List2.ListCount - 1 Then
        MsgBox "下一组皮肤为空，请切换到上一组！", vbInformation, "VB代码管家"
    Else
        List2.ListIndex = List2.ListIndex + 1
        SaveSetting App.Title, "Settings", "pifu", List2.List(List2.ListIndex) '保存选中皮肤
    End If
End Sub

Private Sub List2_Click()    '设置皮肤
    SkinH_AttachEx List2.List(List2.ListIndex), ""
End Sub

Private Sub 导入代码列表事件()
    Dim 当前行 As Integer
    
    代码标题 = "在此输入代码标题"
    代码内容 = "在此输入代码内容"
    List1.Clear '清空列表
    
    If rs.RecordCount <> 0 Then '************************************1
        ReDim 标题数组(1 To rs.RecordCount) '重新定义数组范围
        rs.MoveFirst '----------------------指向第一条
        
        For 当前行 = 1 To rs.RecordCount
            标题数组(当前行) = rs.Fields("标题")
            rs.MoveNext '-----------------------指向下一条
        Next 当前行
        
        If 过滤内容.Text = "" Then '**********************2
            For 当前行 = 1 To rs.RecordCount
                List1.AddItem 标题数组(当前行)
            Next 当前行
        Else '----------------------**********************2
            For 当前行 = 1 To rs.RecordCount
                If InStr(UCase(test(标题数组(当前行))), UCase(test(过滤内容.Text))) > 0 Then
                    List1.AddItem 标题数组(当前行)
                End If
            Next 当前行
        End If '--------------------**********************2
        
    Else                        '************************************1
        MsgBox "该数据库没有任何记录！", vbInformation, "温馨提示"
    End If                      '************************************1
    
    '------------------------------------------------
    If List1.ListCount <> 0 Then
        List1.ListIndex = 0
        For 当前行 = 0 To List1.ListCount - 1
            If 缓存标题 = List1.List(当前行) Then
                List1.ListIndex = 当前行
                If List1.ListCount - 1 - 当前行 > 30 Then
                    List1.ListIndex = 当前行 + 30: List1.ListIndex = 当前行
                Else
                    List1.ListIndex = List1.ListCount - 1: List1.ListIndex = 当前行
                End If
                Exit For
            End If
        Next 当前行
    End If
    '------------------------------------------------选中某行
End Sub

Private Sub List1_Click()  '显示当前内容
    If List1.ListCount > 0 Then '**************
        
        rs.MoveFirst                   '指向第一条
        rs.Find "标题='" & List1.List(List1.ListIndex) & "'"
        If rs.EOF = False Then
            代码标题 = rs.Fields("标题")
            代码内容 = rs.Fields("内容")
            StatusBar1.Panels(4).Text = "修改日期：" & rs.Fields("修改日期")
            Call 检查是否有附件存在
        Else
            MsgBox "对不起，啥也没找到！", vbInformation, "您好"
            Call 导入代码列表事件
        End If
        
    End If '--------------------------------***************
    
    '--------------------------------------------
    'If Clipboard.GetFormat(1) = True Then Clipboard.SetText List1.List(List1.ListIndex)
    '--------------------------------------------复制代码标题
End Sub

Private Sub 检查是否有附件存在()
    '——————————————————————————————————
    On Error GoTo ErrHandle          '用户取消时触发错误
    If CStr(rs.Fields("附件")) <> "" Then
        StatusBar1.Panels(5).Text = "附件信息：有"
    End If
    Exit Sub
ErrHandle:                                   '错误处理
    StatusBar1.Panels(5).Text = "附件信息：无"
    '——————————————————————————————————检查是否有附件存在
End Sub

Private Sub Spy_Click()
    Form2.Show
End Sub

Private Sub 编码转换_Click()
    Form3.Show
End Sub

Private Sub Post模拟_Click()
    Form4.Show
End Sub

Private Sub Timer1_Timer()
    If List1.ListCount = 0 Then '****************1
        
        删除.Enabled = False
        更新.Enabled = False
        附件.Enabled = False
        
        If 代码标题 = "" Then '***3
            添加.Enabled = False
        Else                  '***3
            添加.Enabled = True
        End If                '***3
        
        StatusBar1.Panels(1).Text = "当前总行数：0" '给第一个面板设置文字
        StatusBar1.Panels(2).Text = "当前选中行：0" '给第二个面板设置文字
    Else '-----------------------****************1
        删除.Enabled = True
        
        If 代码标题 = "" Or 代码内容 = "" Then '*********2
            更新.Enabled = False
            添加.Enabled = False
            附件.Enabled = False
        Else                                   '*********2
            更新.Enabled = True
            添加.Enabled = True
            附件.Enabled = True
        End If                                 '*********2
        
        StatusBar1.Panels(1).Text = "当前总行数：" & List1.ListCount                '给第一个面板设置文字
        StatusBar1.Panels(2).Text = "当前选中行：" & List1.ListIndex + 1            '给第二个面板设置文字
    End If '---------------------****************1
    
    If Me.WindowState = 1 Then '如果最小化窗体
        Call 添加托盘图标          '添加图标到托盘
        Form1.Visible = False
    End If
    
    StatusBar1.Panels(3).Text = "系统时间：" & Now
    StatusBar1.Panels(6).Text = "联系作者：QQ:1668066802"
    
End Sub

Private Sub 附件_Click()
    Me.PopupMenu fj, , 附件.Left, 附件.Top + 附件.Height '显示符件菜单
End Sub

Private Sub 工具箱_Click()
    Me.PopupMenu gjx, , 工具箱.Left, 工具箱.Top + 工具箱.Height '显示工具箱菜单
End Sub

Private Sub 皮肤_Click()
    Me.PopupMenu pf, , 皮肤.Left, 皮肤.Top + 皮肤.Height '显示皮肤列表菜单
End Sub

Private Sub 代码标题_KeyPress(KeyAscii As Integer) '过滤字符
    Dim a As String
    a = "`~!@#$%^&*_+-=[];'\./{}:|<>?" '用来存放不接受的字符
    If InStr(1, a, Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub 过滤内容_KeyPress(KeyAscii As Integer) '过滤字符
    Dim a As String
    a = "`~!@#$%^&*_+-=[];'\./{}:|<>?" '用来存放不接受的字符
    If InStr(1, a, Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub 过滤内容_Change() '过滤符合条件标题
    Call 导入代码列表事件
End Sub

Private Sub 更新_Click() '更新记录(更新的是当前行)
    '_________________________________________________________________________________________________________1
    Dim v As String
    v = MsgBox("更新后原有的数据将会被覆盖，是否继续？", vbOKCancel, "VB代码管家")
    If v = vbOK Then '________________________________________________________________________________________1
        Dim 动态数值 As Integer
        Dim 当前行 As Integer
        动态数值 = 1
        For 当前行 = 1 To rs.RecordCount
            If 标题数组(当前行) = 代码标题.Text And 标题数组(当前行) <> List1.List(List1.ListIndex) Then
                动态数值 = 动态数值 + 1
            End If
        Next 当前行
        
        If 动态数值 = 1 Then '_____________2
            rs("标题") = 代码标题           '对应标题列
            rs("内容") = 代码内容           '对应内容列
            rs("修改日期") = Now            '对应修改时间列
            rs.Update
            缓存标题 = 代码标题
            过滤内容 = ""
            Call 导入代码列表事件
            MsgBox "代码更新成功！", vbOKOnly, "提示"
        Else '_____________________________2
            MsgBox "该标题已存在，请重新修改标题后再添加！", vbExclamation, "警告"
            缓存标题 = 代码标题
            过滤内容 = ""
            Call 导入代码列表事件
        End If '___________________________2
    End If '__________________________________________________________________________________________________1
End Sub

Private Sub 删除_Click() '删除当前行
    Dim v As String
    v = MsgBox("您确定要删除标题为:“" & List1.List(List1.ListIndex) & "”的数据吗？", vbOKCancel, "温馨提示")
    If v = vbOK Then          'vbCancel也可换成vbOK则表示确定键
        If List1.ListIndex > 1 Then
            缓存标题 = List1.List(List1.ListIndex - 1)
        End If
        
        rs.Delete
        
        过滤内容 = ""
        Call 导入代码列表事件
    End If
End Sub

Private Sub 添加_Click() '添加记录
    Dim 判断 As Boolean
    Dim 当前行 As Integer
    判断 = True
    
    For 当前行 = 1 To rs.RecordCount
        If 标题数组(当前行) = 代码标题.Text Then
            判断 = False
            Exit For
        End If
    Next 当前行
    
    If 判断 = True Then '***********1
        rs.AddNew
        rs("标题") = 代码标题           '对应标题列
        rs("内容") = 代码内容           '对应内容列
        rs("修改日期") = Now            '对应修改时间列
        rs.Update
        缓存标题 = 代码标题
        过滤内容 = ""
        Call 导入代码列表事件
        MsgBox "代码添加成功！", vbOKOnly, "提示"
    Else '--------------***********1
        MsgBox "该标题已存在，请重新修改标题后再添加！", vbExclamation, "警告"
        缓存标题 = 代码标题
        过滤内容 = ""
        Call 导入代码列表事件
    End If '------------***********1
End Sub

Private Sub Form_Unload(Cancel As Integer) '关闭数据库
    Cancel = True  '取消关闭
    Call 退出
End Sub

Private Sub 退出()
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData) '从托盘删除
    If rs.State = adStateOpen Then
        rs.Close '关闭表
        db.Close '关闭数据库
        
        Name App.Path & "\VB代码数据库.mdb" As App.Path & "\VB代码数据库2.mdb"
        Dim miJRO As JRO.JetEngine
        Set miJRO = New JRO.JetEngine
        miJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0; " & "Data Source=" & App.Path & "\VB代码数据库2.mdb;Jet OLEDB:Database Password=admin", _
        "Provider=Microsoft.Jet.OLEDB.4.0; " & "Data Source=" & App.Path & "\VB代码数据库.mdb;Jet OLEDB:Database Password=admin"
        Kill App.Path & "\VB代码数据库2.mdb"
    End If
End            '关闭
End Sub
