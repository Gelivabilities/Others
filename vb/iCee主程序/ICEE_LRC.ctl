VERSION 5.00
Begin VB.UserControl ICEE_LRC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00241D0A&
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   ControlContainer=   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   10170
   Begin VB.PictureBox fraLrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00241D0A&
      BorderStyle     =   0  'None
      Height          =   12000
      Left            =   960
      ScaleHeight     =   12000
      ScaleWidth      =   30000
      TabIndex        =   0
      Top             =   480
      Width           =   30000
      Begin VB.PictureBox fraMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00241D0A&
         BorderStyle     =   0  'None
         Height          =   12000
         Left            =   1800
         ScaleHeight     =   12000
         ScaleWidth      =   30000
         TabIndex        =   2
         Top             =   1200
         Width           =   30000
         Begin VB.Label lblCurLrc 
            BackColor       =   &H001F1F1F&
            BackStyle       =   0  'Transparent
            Caption         =   "当前播放的歌词"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0044DFE3&
            Height          =   300
            Left            =   0
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   0
            UseMnemonic     =   0   'False
            Width           =   7065
         End
      End
      Begin VB.Label lblEachLrc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "每一句歌词"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00989287&
         Height          =   540
         Index           =   0
         Left            =   -120
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   600
         Width           =   4905
      End
   End
End
Attribute VB_Name = "ICEE_LRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private m_Karaoke       As Boolean      '是否卡拉OK模式
Public DE_SIZE As Long
Dim LES

Property Let Karaoke(NewValue As Boolean)
    m_Karaoke = NewValue
End Property

Property Get Karaoke() As Boolean
    Karaoke = m_Karaoke
End Property

'设置前景色
Property Let ShowColor(Color As OLE_COLOR)
    lblCurLrc.FOREColor = Color
End Property
'得到前景色
Property Get ShowColor() As OLE_COLOR
    FOREColor = lblCurLrc.FOREColor
End Property

'得到字体
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

'设置字体
Public Property Set Font(ByVal New_Font As Font)
On Error GoTo ERR
    Dim i   As Integer
    
    Set UserControl.Font = New_Font
    Set lblCurLrc.Font = New_Font
    For i = 0 To iLrcRows
        Set lblEachLrc(i).Font = New_Font
    Next
    PropertyChanged "Font"
ERR:
    Call ResizeLrc                  '重新排列控件
End Property

'读取文件
Public Function ReadFile(LrcFileName As String)
    Call ClearLrc
    Call 歌词模块.ReadFile(LrcFileName)
    
    If iLrcRows > 0 Then
        Call LoadLrcLabel       '创建 lblEachLrc()控件
    End If
End Function

Public Sub ClearLrc()
    Dim i           As Integer
On Error Resume Next
    Erase myLrc()               '清空以前的歌词信息
    For i = 1 To iLrcRows       '卸载动态加载的控件
        Unload lblEachLrc(i)
    Next i
    lblEachLrc(0).Caption = ""
    
    Call 歌词模块.ClearLrc
End Sub

Public Sub StopLrc()
    iCurPlay = 0
    fraMask.Width = 0
End Sub

Private Sub fraLrc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub fraMask_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print X, Y

End Sub

Private Sub fraMask_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub lblCurLrc_Change()
If lblCurLrc.Caption = "中文歌词库 www.CnLyric.com" Then lblCurLrc.Caption = ""
End Sub

Private Sub lblCurLrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCurLrc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub lblEachLrc_Change(Index As Integer)
If lblEachLrc(Index).Caption = "中文歌词库 www.CnLyric.com" Then lblEachLrc(Index).Caption = ""

End Sub

Private Sub lblEachLrc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
lblCurLrc.Caption = ""
lblEachLrc(0).Caption = ""
lblCurLrc.FontName = "微软雅黑"
lblEachLrc(0).FontName = "微软雅黑"
DE_SIZE = 12
lblEachLrc(0).FontSize = DE_SIZE
lblCurLrc.FontSize = DE_SIZE
ReDim Preserve myLrc(0)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblEachLRC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub fraLrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


Public Sub SeekLrc(sTime As Double)    '设置当前歌词
    If Not 歌词模块.SeekLrc(sTime) Then Exit Sub
    On Error Resume Next
    If iCurPlay = -1 Then
        fraLrc.Top = (UserControl.Height) / 2                   '上下居中，并减去本句播放的部分
        fraMask.Width = 0
        Exit Sub
    End If
    
    Dim sLrcShowWidth           As Single               '本句歌词的显示宽度
    Dim sLrcShowHeight          As Single               '本句歌词的显示高度 //确认说每句高度都一样
    Dim sLrcShowLeft            As Single               '本句歌词左边位置
    Dim iCurTimer               As Double               '本句歌词的播放长度,用于计算滚动比例
    
    If iCurPlay = iLrcRows Then
        iCurTimer = 10
    Else
        iCurTimer = myLrc(iCurPlay + 1).lrcTime - myLrc(iCurPlay).lrcTime ' - 0.1
    End If
    sLrcShowWidth = UserControl.TextWidth(myLrc(iCurPlay).lrcString)    '本句歌词总宽度
    sLrcShowHeight = UserControl.TextHeight("Ag")                       '本句歌词总高度
    '调整歌词位置
    fraLrc.Top = (UserControl.Height) / 2 - sLrcShowHeight * iCurPlay - _
                sLrcShowHeight * (sTime - myLrc(iCurPlay).lrcTime) / iCurTimer    '上下居中，并减去本句播放的部分
    lblCurLrc.Caption = myLrc(iCurPlay).lrcString     '得到当前歌词
    fraMask.Move 0, lblEachLrc(iCurPlay).Top, _
            (fraLrc.Width - sLrcShowWidth) / 2 + sLrcShowWidth * IIf(m_Karaoke, (sTime - myLrc(iCurPlay).lrcTime) / iCurTimer, 1), sLrcShowHeight
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_Terminate()
    Erase myLrc()
End Sub

'将歌词信息加载到 lblEachLrc() 控件
Public Function LoadLrcLabel()
    Dim i           As Integer
On Error Resume Next
    For i = 0 To iLrcRows
        If i Then
            Load lblEachLrc(i)
        End If
        lblEachLrc(i).BackStyle = 0             '透明
        lblEachLrc(i).Alignment = 2             '居中
        lblEachLrc(i).UseMnemonic = False       '不转换 & 为下划线
        lblEachLrc(i).Caption = myLrc(i).lrcString
        lblEachLrc(i).AUTOSIZE = False
    Next i
    Call ResizeLrc
End Function

'重新设置控件大小
Private Sub ResizeLrc()
On Error Resume Next
    Dim i       As Integer
    Dim sWidth  As Single, sHeight  As Single
    Dim sMax    As Single
    
    With UserControl
        sWidth = 0
        sHeight = .TextHeight("Ag")
        For i = 0 To iLrcRows                       '找到歌词的最大宽度
            sMax = .TextWidth(myLrc(i).lrcString)
            If sMax > sWidth Then sWidth = sMax
        Next
'        sWidth = (sWidth \ Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
        lblCurLrc.Width = sWidth
        fraMask.Width = sWidth
        lblCurLrc.Alignment = 2
        fraLrc.Move (.Width - sWidth) / 2, .Height, sWidth, iLrcRows * sHeight             '容器大小及位置 /在控件中水平居中,垂直不显示
    End With
    For i = 0 To iLrcRows                       '调整位置等..
        lblEachLrc(i).Move 0, i * sHeight, sWidth                   '设置等宽(实际效果是居中显示)
        lblEachLrc(i).Visible = True                                '显示
    Next
    
    lRet = SetInitEntry("PLAYER", "LRC_SIZE", DE_SIZE)
End Sub

'背景色
Property Let BackColor(Color As Long)
    UserControl.BackColor = Color
    fraLrc.BackColor = Color
    fraMask.BackColor = Color
End Property
'背景色
Property Get BackColor() As Long
    BackColor = UserControl.BackColor
End Property
'设置前景色
Property Let FOREColor(Color As Long)
    Dim i As Integer
    For i = 0 To lblEachLrc.UBound
        lblEachLrc(i).FOREColor = Color
    Next i
End Property
'得到前景色
Property Get FOREColor() As Long
    FOREColor = lblEachLrc(0).FOREColor
End Property

'得到每行的高度
Property Get LineHeight() As Long
    LineHeight = lblEachLrc(0).Height
End Property
'得到当前移动行的Width
Public Function GetCurrWidth(Index As Integer) As Single
    On Error Resume Next
    GetCurrWidth = lblEachLrc(Index).Width
End Function
'得到当前移动行的Width
Public Function GetLineTop(Index As Integer) As Single
    On Error Resume Next
    GetLineTop = lblEachLrc(Index).Top
End Function
'得到当前移动行的Width
Public Function GetLineLeft(Index As Integer) As Single
    On Error Resume Next
    GetLineLeft = lblEachLrc(Index).Left
End Function
'得到当前的歌词内容
Public Function GetCurrLrc(Index As Integer) As String
    GetCurrLrc = lblEachLrc(Index).Caption
End Function

Private Sub UserControl_Resize()
    ResizeLrc
End Sub


Sub SETFONTSIZE(Size As Long)
Dim i As Integer
For i = 0 To lblEachLrc.Count - 1
lblEachLrc(i).FontSize = Size
Next
DE_SIZE = Size
lblCurLrc.Font = Size
If Size <= 6 Then Size = 6: DE_SIZE = 6
Call ResizeLrc
lRet = SetInitEntry("LRC", "SIZE", DE_SIZE)

End Sub

Sub SETPIC(HD As Object, X As Single, Y As Single)
LES = BitBlt(UserControl.hdc, 0, 0, UserControl.Width, UserControl.Height, HD.hdc, X, Y, &HCC0020)
UserControl.Refresh
LES = BitBlt(UserControl.hdc, 0, 0, fraLrc.Width, fraLrc.Height, UserControl.hdc, fraLrc.Left, fraLrc.Top, &HCC0020)
fraLrc.Refresh
LES = BitBlt(fraMask.hdc, 0, 0, fraMask.Width, fraMask.Height, UserControl.hdc, fraMask.Left, fraMask.Top, &HCC0020)
fraMask.Refresh
End Sub
