VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeskLrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim frm     As FRMSHOW
Private m_Karaoke       As Boolean      '是否卡拉OK模式

Private Sub Class_Initialize()
    Set frm = New FRMSHOW
    frm.Show
End Sub
'
Private Sub Class_Terminate()
    Unload frm
    Set frm = Nothing
End Sub

Property Let Karaoke(NewValue As Boolean)
    m_Karaoke = NewValue
End Property

Property Get Karaoke() As Boolean
    Karaoke = m_Karaoke
End Property

Public Sub ShowText(ByVal Text As String)
    Call frm.DrawText(Text, 0)
End Sub

Public Sub ReDraw()
    Call frm.ReDrawText
End Sub

Public Property Get FontName() As String
    FontName = frm.LrcFontName
End Property

Public Property Let FontName(ByVal NewName As String)
    frm.LrcFontName = NewName
End Property

Public Property Get FontSize() As Single
    FontSize = frm.LrcFontSize
End Property

Public Property Let FontSize(ByVal NewSize As Single)
    frm.LrcFontSize = NewSize
End Property

Public Property Get FontBold() As Boolean
    FontBold = frm.LrcFontBold
End Property

Public Property Let FontBold(ByVal Bold As Boolean)
    frm.LrcFontBold = Bold
End Property

Public Property Get BackColor1() As OLE_COLOR
    BackColor1 = frm.BackColor1
End Property

Public Property Let BackColor1(ByVal NewColor As OLE_COLOR)
    frm.BackColor1 = NewColor
End Property

Public Property Get BackColor2() As OLE_COLOR
    BackColor2 = frm.BackColor2
End Property

Public Property Let BackColor2(ByVal NewColor As OLE_COLOR)
    frm.BackColor2 = NewColor
End Property

Public Property Get ForeColor1() As OLE_COLOR
    ForeColor1 = frm.ForeColor1
End Property

Public Property Let ForeColor1(ByVal NewColor As OLE_COLOR)
    frm.ForeColor1 = NewColor
End Property

Public Property Get ForeColor2() As OLE_COLOR
    ForeColor2 = frm.ForeColor2
End Property

Public Property Let ForeColor2(ByVal NewColor As OLE_COLOR)
    frm.ForeColor2 = NewColor
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = frm.LineColor
End Property

Public Property Let LineColor(ByVal NewColor As OLE_COLOR)
    frm.LineColor = NewColor
End Property


Public Sub SeekLrc(sTime As Double, Optional ReSeek As Boolean = False)   '设置当前歌词
On Error Resume Next
    If ReSeek Then
        Call 歌词模块.SeekLrc(sTime)
    End If
    
    Dim sLrcShowWidth           As Single               '本句歌词的显示宽度
    Dim iCurTimer               As Double               '本句歌词的播放长度,用于计算滚动比例
    Dim iWidth                  As Long                 '卡拉OK歌词宽度
    
    If m_Karaoke Then
        If iCurPlay = iLrcRows Then
            iCurTimer = 10
        Else
'            iCurTimer = myLrc(iCurPlay + 1).lrcTime - myLrc(iCurPlay).lrcTime ' - 0.1
        End If
'        sLrcShowWidth = frm.TextWidth(myLrc(iCurPlay).lrcString)    '本句歌词总宽度
        
        '已唱的时间 / 总时间 * 总宽度
'        iWidth = (sTime - myLrc(iCurPlay).lrcTime) / iCurTimer * sLrcShowWidth
    Else
        iWidth = 0
    End If
    
    If iCurPlay = -1 Then
        frm.DrawText "", iWidth
    Else
        frm.DrawText myLrc(iCurPlay).lrcString, iWidth
    End If
End Sub

