VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPY++"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8070
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox PID值 
      Height          =   270
      Left            =   3840
      TabIndex        =   23
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox 进程名 
      Height          =   270
      Left            =   960
      TabIndex        =   22
      Top             =   855
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5760
      Top             =   1800
   End
   Begin VB.TextBox 颜色值 
      Height          =   270
      Left            =   6720
      TabIndex        =   10
      Top             =   2655
      Width           =   1215
   End
   Begin VB.TextBox 当前标题 
      Height          =   270
      Left            =   960
      TabIndex        =   9
      Top             =   2655
      Width           =   4815
   End
   Begin VB.TextBox 顶层标题 
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Top             =   2295
      Width           =   4815
   End
   Begin VB.TextBox 进程路径 
      Height          =   270
      Left            =   960
      TabIndex        =   7
      Top             =   1935
      Width           =   4815
   End
   Begin VB.TextBox 类名过程 
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   1575
      Width           =   4815
   End
   Begin VB.TextBox 子句柄 
      Height          =   270
      Left            =   960
      TabIndex        =   5
      Top             =   495
      Width           =   1935
   End
   Begin VB.TextBox 句柄过程 
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Top             =   1215
      Width           =   4815
   End
   Begin VB.TextBox 父句柄 
      Height          =   270
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox 截取坐标 
      Height          =   270
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox 当前坐标 
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   5880
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   840
         Picture         =   "Form2.frx":030A
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Label Label13 
      Caption         =   "P I D 值:"
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "进 程 名:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "R G B 值:"
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   2685
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "拖动图标到目标位置上面:"
      Height          =   240
      Left            =   5880
      TabIndex        =   20
      Top             =   135
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "当前标题:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2685
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "顶层标题:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2325
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "进程路径:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1965
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "类名过程:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1605
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "句柄过程:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1245
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "子 句 柄:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "父 句 柄:"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "截取坐标:"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "当前坐标:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   165
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'---------------------鼠标形状
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'---------------------取坐标下的句柄
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'---------------------根据子句柄取得主窗体句柄
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpclassname As String, ByVal nMaxCount As Long) As Long
'---------------------取窗体类名
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'---------------------取某顶层标题
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'---------------------根据句柄取进程路径
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'---------------------运行文件
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
'---------------------取鼠标坐标

'-------------------------------------------------------------------------------------------------------------------
Dim a     As POINTAPI
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long '获取指定窗口的设备场景
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long '释放设备场景
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long '取RGB值
Private Sub GetRGB(ByVal Col As Long, ByRef r As Long, ByRef g As Long, ByRef B As Long) '转化为RGB
    r = Col Mod 256
    g = ((Col And &HFF00&) \ 256&) Mod 256&
    B = (Col And &HFF0000) \ 65536
End Sub
'-------------------------------------------------------------------------------------------------------------------取鼠标下的颜色值


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Visible = False
    SetCursor Image1.Picture
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    截取坐标.Text = ""
    子句柄.Text = ""
    父句柄.Text = ""
    句柄过程.Text = ""
    类名过程.Text = ""
    进程路径.Text = ""
    顶层标题.Text = ""
    当前标题.Text = ""
    颜色值.Text = ""
    进程名.Text = ""
    PID值.Text = ""
    
    Dim 句柄储存 As String
    
    Dim a As POINTAPI
    GetCursorPos a
    
    截取坐标 = a.x & "," & a.y
    
    子句柄 = WindowFromPoint(a.x, a.y)
    父句柄 = WindowFromPoint(a.x, a.y)
    
    '***************************************取当前标题
    当前标题 = ""
    Dim dq As String
    dq = Space(255) '此句相当于空隔
    GetWindowText 子句柄, dq, 255 'hwnd为窗体句柄
    当前标题 = dq
    '***************************************
    
    句柄过程 = ""
    句柄过程 = 父句柄
    Dim leiming As String
    leiming = Space(255)
    GetClassName 父句柄, leiming, 255
    类名过程 = ""
    类名过程 = leiming
    
1:
    句柄储存 = GetParent(父句柄) '父句柄为子句柄,句柄储存为上一层控件的句柄
    If 句柄储存 <> 0 Then '如果不是主窗体句柄
        句柄过程 = 句柄过程 & "\" & 句柄储存
        GetClassName 句柄储存, leiming, 255
        类名过程 = 类名过程 & "\" & leiming
        父句柄 = 句柄储存
        句柄储存 = ""
        GoTo 1: '则跳转到标签1：处
    End If
    
    '***************************************取顶层标题
    顶层标题 = ""
    Dim bt As String
    bt = Space(255) '此句相当于空隔
    GetWindowText 父句柄, bt, 255 'hwnd为窗体句柄
    顶层标题 = bt
    '***************************************
    
    '***************************************根据句柄取进程路径
    Dim PID As Long
    GetWindowThreadProcessId 父句柄, PID '根据句柄取得进程ID,text1为句柄
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process")
    For Each objprocess In colProcesses
        
        If objprocess.processid = PID Then  '当某进程ID与句柄ID相同时
            进程路径 = objprocess.ExecutablePath  '取得进程路径
            进程名 = objprocess.Name
            PID值 = PID
        End If
        
    Next
    '***************************************
    '***************************************取鼠标下的颜色值
    Dim hdc     As Long, Col       As Long
    Dim r     As Long, g       As Long, B       As Long
    hdc = GetDC(0)
    GetCursorPos a
    cor = GetPixel(hdc, a.x, a.y) 'a.x为横坐标，a.y为竖坐标
    GetRGB cor, r, g, B
    ReleaseDC Me.hwnd, hdc
    颜色值 = r & "," & g & "," & B
    Frame1.BackColor = RGB(r, g, B) '改变Picture1的颜色值死活都改不过来，只好用Frame1
    Image1.Visible = True
End Sub

Private Sub Timer1_Timer()
    Dim a As POINTAPI
    GetCursorPos a
    当前坐标 = a.x & "," & a.y
End Sub
