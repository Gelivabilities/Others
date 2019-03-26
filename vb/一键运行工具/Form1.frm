VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "一键运行BY：kill_BCL"
   ClientHeight    =   3810
   ClientLeft      =   675
   ClientTop       =   585
   ClientWidth     =   5640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5640
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "一键运行BY：kill_BCL （LYC） 制作者QQ：754571662"
      Top             =   2640
      Width           =   5175
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "一键运行BY：Gelivability 754571662"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "一键运行BY：kill_BCL（LYC）754571662"
         Top             =   120
         Width           =   4935
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5295
      Left            =   0
      TabIndex        =   8
      Top             =   4320
      Width           =   14175
      ExtentX         =   25003
      ExtentY         =   9340
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin 一键运行工具.CandyButton CandyButton31 
      Height          =   375
      Left            =   13200
      TabIndex        =   7
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "转到该页"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   13215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5175
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "系统时间"
         Top             =   120
         Width           =   4815
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "系统信息"
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5415
      Begin VB.Frame Frame11 
         Caption         =   "文件处理"
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         Begin VB.Timer Timer2 
            Interval        =   1
            Left            =   3120
            Top             =   600
         End
      End
      Begin VB.Label Label1 
         Caption         =   "正在检测CPU"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   33.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   5175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private CPU As clsCPUUsage
  Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
  Private Const CCHDEVICENAME = 32
  Private Const CCHFORMNAME = 32
  Private Const ENUM_CURRENT_SETTINGS = 1
  Private Type DEVMODE
                  dmDeviceName   As String * CCHDEVICENAME
                  dmSpecVersion   As Integer
                  dmDriverVersion   As Integer
                  dmSize   As Integer
                  dmDriverExtra   As Integer
                  dmFields   As Long
                  dmOrientation   As Integer
                  dmPaperSize   As Integer
                  dmPaperLength   As Integer
                  dmPaperWidth   As Integer
                  dmScale   As Integer
                  dmCopies   As Integer
                  dmDefaultSource   As Integer
                  dmPrintQuality   As Integer
                  dmColor   As Integer
                  dmDuplex   As Integer
                  dmYResolution   As Integer
                  dmTTOption   As Integer
                  dmCollate   As Integer
                  dmFormName   As String * CCHFORMNAME
                  dmUnusedPadding   As Integer
                  dmBitsPerPel   As Long
                  dmPelsWidth   As Long
                  dmPelsHeight   As Long
                  dmDisplayFlags   As Long
                  dmDisplayFrequency   As Long
  End Type
  Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (ByVal lpDevMode As Long, ByVal dwFlags As Long) As Long
  Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As Any) As Long
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
  Private Const SM_CXSCREEN = 0
  Private Const SM_CYSCREEN = 1
  Dim xx     As Integer
  Dim yy     As Integer
  '设置显示器分辨率的执行函数
  Private Function SetDisplayMode(Width As Integer, Height As Integer, Color As Integer) As Long                           ',   Freq   As   Long)   As   Long
          Dim pNewMode     As DEVMODE
          Dim pOldMode     As Long
          On Error GoTo ErrorHandler
          Const DM_PELSWIDTH = &H80000
          Const DM_PELSHEIGHT = &H100000
          Const DM_BITSPERPEL = &H40000
          Const DM_DISPLAYFLAGS = &H200000
          With pNewMode
                  .dmSize = Len(pNewMode)
                  If Color = 0 Then           'Color   =   0   时不更改屏幕颜色
                          .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
                  Else
                          .dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT                 'Or   DM_DISPLAYFREQUENCY'属性率的更改还是没办法,不过,不加入此DM_DISPLAYFREQUENCY这个参数,只要系统支持,应该不会更改刷新率的
                  End If
                  .dmPelsWidth = Width
                  .dmPelsHeight = Height
                  If Color <> 0 Then
                  .dmBitsPerPel = Color
                  End If
          End With
          pOldMode = lstrcpy(pNewMode, pNewMode)
          SetDisplayMode = ChangeDisplaySettings(pOldMode, 1)
          Exit Function
ErrorHandler:
          MsgBox Err.Description, vbCritical, "系统错误！"
  End Function



  Private Sub Form_Unload(Cancel As Integer)
  SetDisplayMode xx, yy, 32
  Set CPU = Nothing
  End Sub
Private Sub CandyButton1_Click()
Shell "C:\Program Files\KuGou\KuGou2008\KuGoo.exe", vbNormalFocus
End Sub
Private Sub CandyButton10_Click()
Shell "C:\Program Files\Thunder Network\Thunder\Program\Thunder.exe", vbNormalFocus
End Sub
Private Sub CandyButton11_Click()
Shell "C:\Program Files\GIF Movie Gear\movgear.exe", vbNormalFocus
End Sub
Private Sub CandyButton12_Click()
Shell "C:\Program Files\Tencent\QQ\Bin\QQ.exe", vbNormalFocus
End Sub
Private Sub CandyButton13_Click()
Shell "C:\CYYSoft\ScreenREC\CyyScreenREC.exe", vbNormalFocus
End Sub
Private Sub CandyButton14_Click()
Shell "D:\安装\Google Earth\googleearth.exe", vbNormalFocus
End Sub
Private Sub CandyButton15_Click()
Shell "C:\Program Files\一键清理垃圾\onekeyclear.exe", vbNormalFocus
End Sub
Private Sub CandyButton16_Click()
Shell "C:\Program Files\Super Rabbit\MagicSet\srgui9.exe", vbNormalFocus
End Sub
Private Sub CandyButton17_Click()
Form1.Height = 10800
Dim IE As String
Text3 = "cf.qq.com/index.shtml"
WebBrowser1.Navigate "http://cf.qq.com/index.shtml"
CandyButton28.Visible = False
End Sub
Private Sub CandyButton18_Click()
Set IE = CreateObject("internetexplorer.application")
IE.Visible = True
IE.Navigate "D:\vb\x.html"
End Sub
Private Sub CandyButton19_Click()
Dim wd
wd = Text1
Form1.Height = 10800
Dim IE As String
WebBrowser1.Navigate "http://www.google.cn/search?hl=zh-CN&q=" & wd
CandyButton28.Visible = False
End Sub
Private Sub CandyButton2_Click()
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE", vbNormalFocus
End Sub
Private Sub CandyButton20_Click()
Dim wd
wd = Text2
Form1.Height = 10800
Dim IE As String
WebBrowser1.Navigate "http://www.baidu.com/s?wd=" & wd
CandyButton28.Visible = False
End Sub


Private Sub CandyButton23_Click()
Shell "D:\软件\yy\Start.exe", vbNormalFocus
End Sub
Private Sub CandyButton24_Click()
CandyButton1.Style = 0
CandyButton2.Style = 0

CandyButton4.Style = 0
CandyButton5.Style = 0
CandyButton6.Style = 0
CandyButton7.Style = 0

CandyButton9.Style = 0
CandyButton10.Style = 0
CandyButton11.Style = 0
CandyButton12.Style = 0


CandyButton15.Style = 0
CandyButton16.Style = 0
CandyButton17.Style = 0
CandyButton18.Style = 0
CandyButton19.Style = 0
CandyButton20.Style = 0


CandyButton23.Style = 0
CandyButton24.Style = 0
CandyButton25.Style = 0
CandyButton26.Style = 0
CandyButton27.Style = 0
CandyButton28.Style = 0
CandyButton30.Style = 0
CandyButton29.Style = 0
CandyButton31.Style = 0
CandyButton32.Style = 0
CandyButton33.Style = 0
CandyButton34.Style = 0
CandyButton35.Style = 0
CandyButton36.Style = 0
End Sub
Private Sub CandyButton25_Click()
CandyButton1.Style = 2
CandyButton2.Style = 2

CandyButton4.Style = 2
CandyButton5.Style = 2
CandyButton6.Style = 2
CandyButton7.Style = 2

CandyButton9.Style = 2
CandyButton10.Style = 2
CandyButton11.Style = 2
CandyButton12.Style = 2


CandyButton15.Style = 2
CandyButton16.Style = 2
CandyButton17.Style = 2
CandyButton18.Style = 2
CandyButton19.Style = 2
CandyButton20.Style = 2
CandyButton27.Style = 2


CandyButton23.Style = 2
CandyButton24.Style = 2
CandyButton25.Style = 2
CandyButton26.Style = 2
CandyButton29.Style = 2
CandyButton28.Style = 2
CandyButton30.Style = 2
CandyButton31.Style = 2
CandyButton32.Style = 2
CandyButton33.Style = 2
CandyButton34.Style = 2
CandyButton35.Style = 2
CandyButton36.Style = 2
End Sub
Private Sub CandyButton26_Click()
CandyButton1.Style = 3
CandyButton2.Style = 3

CandyButton4.Style = 3
CandyButton5.Style = 3
CandyButton6.Style = 3
CandyButton7.Style = 3

CandyButton9.Style = 3
CandyButton10.Style = 3
CandyButton11.Style = 3
CandyButton12.Style = 3


CandyButton15.Style = 3
CandyButton16.Style = 3
CandyButton17.Style = 3
CandyButton18.Style = 3
CandyButton19.Style = 3
CandyButton20.Style = 3


CandyButton23.Style = 3
CandyButton24.Style = 3
CandyButton25.Style = 3
CandyButton26.Style = 3
CandyButton28.Style = 3
CandyButton27.Style = 3
CandyButton30.Style = 3
CandyButton29.Style = 3
CandyButton31.Style = 3
CandyButton32.Style = 3
CandyButton33.Style = 3
CandyButton34.Style = 3
CandyButton35.Style = 3
CandyButton36.Style = 3
End Sub
Private Sub CandyButton27_Click()
Dim wd
wd = Text4
Form1.Height = 10800
Dim IE As String
WebBrowser1.Navigate "http://www.google.cn/dictionary?q=" & wd & "&langpair=en|zh&hl=zh-CN&ei=sGi_StyBD86fkQWj-91W&sa=X&oi=translation&ct=result"
CandyButton28.Visible = False
End Sub
Private Sub CandyButton28_Click()
Form1.Height = 10800
CandyButton28.Visible = False
End Sub
Private Sub CandyButton29_Click()
CandyButton1.Style = 5
CandyButton2.Style = 5

CandyButton4.Style = 5
CandyButton5.Style = 5
CandyButton6.Style = 5
CandyButton7.Style = 5

CandyButton9.Style = 5
CandyButton10.Style = 5
CandyButton11.Style = 5
CandyButton12.Style = 5


CandyButton15.Style = 5
CandyButton16.Style = 5
CandyButton17.Style = 5
CandyButton18.Style = 5
CandyButton19.Style = 5
CandyButton20.Style = 5


CandyButton23.Style = 5
CandyButton24.Style = 5
CandyButton25.Style = 5
CandyButton26.Style = 5
CandyButton28.Style = 5
CandyButton30.Style = 5
CandyButton29.Style = 5
CandyButton31.Style = 5
CandyButton32.Style = 5
CandyButton27.Style = 5
CandyButton33.Style = 5
CandyButton34.Style = 5
CandyButton35.Style = 5
CandyButton36.Style = 5
End Sub

Private Sub CandyButton30_Click()
Form1.Height = 4965
CandyButton28.Visible = True
End Sub
Private Sub CandyButton31_Click()
Dim IE As String
IE = Text3
WebBrowser1.Navigate IE + ""
End Sub
Private Sub CndyButton35_Click()
Shell "explorer"
End Sub
Private Sub CandyButton32_Click()
Text3 = "http://www.qq.com"
Form1.Height = 10800
Dim IE As String
IE = Text3
WebBrowser1.Navigate "http://www.qq.com"
CandyButton28.Visible = False
End Sub
Private Sub CandyButton33_Click()
Text3 = "dnf.qq.com/index.shtml"
Form1.Height = 10800
Dim IE As String
IE = Text3
WebBrowser1.Navigate "http://dnf.qq.com/index.shtml"
CandyButton28.Visible = False
End Sub
Private Sub CandyButton34_Click()
Text3 = "www.2345.com"
Form1.Height = 10800
Dim IE As String
IE = Text3
WebBrowser1.Navigate "http://www.2345.com"
CandyButton28.Visible = False
End Sub

Private Sub CandyButton35_Click()
Text3 = "tool.114la.com/urlconvert.html"
Form1.Height = 10800
Dim IE As String
IE = Text3
WebBrowser1.Navigate "http://tool.114la.com/urlconvert.html"
CandyButton28.Visible = False
End Sub

Private Sub CandyButton36_Click()
Shell "D:\vb\d.bat"
End Sub
Private Sub CandyButton37_Click()
Shell "D:\vb\一键运行工具\一键运行工具.vbp", vbNormalFocus
End Sub
Private Sub CandyButton5_Click()
Shell "D:\vb\wdwd.bat"
End Sub
Private Sub CandyButton4_Click()
Shell "D:\vb\ACE_btjy.bat"
End Sub
Private Sub CandyButton6_Click()
Shell "D:\vb\ptlxcfc.bat"
End Sub
Private Sub CandyButton7_Click()
Shell "D:\vb\yzhlxcfc.bat"
End Sub
Private Sub CandyButton8_Click()
Shell "C:\Program Files\Fraps\fraps.exe", vbNormalFocus
End Sub
Private Sub CandyButton9_Click()
Shell "C:\Program Files\FreeTime\FormatFactory\FormatFactory.exe", vbNormalFocus
End Sub
Private Sub Form_Load()
    Set CPU = New clsCPUUsage
Dim wmiObjSet As SWbemObjectSet
Dim obj As SWbemObject
Dim Msg As String
Dim dtb As String
Dim d As String
Dim t As String
Dim bias As Long
On Local Error Resume Next
Set wmiObjSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_OperatingSystem")
For Each obj In wmiObjSet
Label5.Caption = "你当前使用的系统是 " & obj.Caption
Next
If Screen.Width / Screen.TwipsPerPixelX + Screen.Height / Screen.TwipsPerPixelY < 1792 Then
MsgBox "请把屏幕分辨率调到1024×768以上后再运行本程序"
Unload Me
Else
 If App.PrevInstance = True Then
MsgBox "本程序已运行"
Unload Me
End If
End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dim wd
wd = Text1
Form1.Height = 10800
Dim IE As String
WebBrowser1.Navigate "http://www.google.cn/search?hl=zh-CN&q=" & wd
CandyButton28.Visible = False
End If
End Sub
Private Sub Text2_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dim wd
wd = Text2
Form1.Height = 10800
Dim IE As String
WebBrowser1.Navigate "http://www.baidu.com/s?wd=" & wd
CandyButton28.Visible = False
End If
End Sub




Private Sub Text3_Keydown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dim IE As String
IE = Text3
WebBrowser1.Navigate IE + ""
End If
End Sub
Private Sub Timer1_Timer()
Label1.Caption = "CPU占用率: " & CPU.Usage & "%"
Label1.ForeColor = &HFF00&
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Dim wd
wd = Text4
Form1.Height = 10800
Dim IE As String
WebBrowser1.Navigate "http://www.google.cn/dictionary?q=" & wd & "&langpair=en|zh&hl=zh-CN&ei=sGi_StyBD86fkQWj-91W&sa=X&oi=translation&ct=result"
CandyButton28.Visible = False
End If
End Sub
Private Sub Timer2_Timer()
Label3.Caption = Now()
Label3.ToolTipText = Now()
  xx = Screen.Width / Screen.TwipsPerPixelX
  yy = Screen.Height / Screen.TwipsPerPixelY
     Label4.Caption = "分辨率：" & Screen.Width / Screen.TwipsPerPixelX & "×" & Screen.Height / Screen.TwipsPerPixelY
     Label4.ToolTipText = "分辨率：" & Screen.Width / Screen.TwipsPerPixelX & "×" & Screen.Height / Screen.TwipsPerPixelY
End Sub
