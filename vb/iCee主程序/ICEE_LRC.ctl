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
            Caption         =   "��ǰ���ŵĸ��"
            BeginProperty Font 
               Name            =   "����"
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
         Caption         =   "ÿһ����"
         BeginProperty Font 
            Name            =   "����"
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
Private m_Karaoke       As Boolean      '�Ƿ���OKģʽ
Public DE_SIZE As Long
Dim LES

Property Let Karaoke(NewValue As Boolean)
    m_Karaoke = NewValue
End Property

Property Get Karaoke() As Boolean
    Karaoke = m_Karaoke
End Property

'����ǰ��ɫ
Property Let ShowColor(Color As OLE_COLOR)
    lblCurLrc.FOREColor = Color
End Property
'�õ�ǰ��ɫ
Property Get ShowColor() As OLE_COLOR
    FOREColor = lblCurLrc.FOREColor
End Property

'�õ�����
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

'��������
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
    Call ResizeLrc                  '�������пؼ�
End Property

'��ȡ�ļ�
Public Function ReadFile(LrcFileName As String)
    Call ClearLrc
    Call ���ģ��.ReadFile(LrcFileName)
    
    If iLrcRows > 0 Then
        Call LoadLrcLabel       '���� lblEachLrc()�ؼ�
    End If
End Function

Public Sub ClearLrc()
    Dim i           As Integer
On Error Resume Next
    Erase myLrc()               '�����ǰ�ĸ����Ϣ
    For i = 1 To iLrcRows       'ж�ض�̬���صĿؼ�
        Unload lblEachLrc(i)
    Next i
    lblEachLrc(0).Caption = ""
    
    Call ���ģ��.ClearLrc
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
If lblCurLrc.Caption = "���ĸ�ʿ� www.CnLyric.com" Then lblCurLrc.Caption = ""
End Sub

Private Sub lblCurLrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCurLrc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub lblEachLrc_Change(Index As Integer)
If lblEachLrc(Index).Caption = "���ĸ�ʿ� www.CnLyric.com" Then lblEachLrc(Index).Caption = ""

End Sub

Private Sub lblEachLrc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
lblCurLrc.Caption = ""
lblEachLrc(0).Caption = ""
lblCurLrc.FontName = "΢���ź�"
lblEachLrc(0).FontName = "΢���ź�"
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


Public Sub SeekLrc(sTime As Double)    '���õ�ǰ���
    If Not ���ģ��.SeekLrc(sTime) Then Exit Sub
    On Error Resume Next
    If iCurPlay = -1 Then
        fraLrc.Top = (UserControl.Height) / 2                   '���¾��У�����ȥ���䲥�ŵĲ���
        fraMask.Width = 0
        Exit Sub
    End If
    
    Dim sLrcShowWidth           As Single               '�����ʵ���ʾ���
    Dim sLrcShowHeight          As Single               '�����ʵ���ʾ�߶� //ȷ��˵ÿ��߶ȶ�һ��
    Dim sLrcShowLeft            As Single               '���������λ��
    Dim iCurTimer               As Double               '�����ʵĲ��ų���,���ڼ����������
    
    If iCurPlay = iLrcRows Then
        iCurTimer = 10
    Else
        iCurTimer = myLrc(iCurPlay + 1).lrcTime - myLrc(iCurPlay).lrcTime ' - 0.1
    End If
    sLrcShowWidth = UserControl.TextWidth(myLrc(iCurPlay).lrcString)    '�������ܿ��
    sLrcShowHeight = UserControl.TextHeight("Ag")                       '�������ܸ߶�
    '�������λ��
    fraLrc.Top = (UserControl.Height) / 2 - sLrcShowHeight * iCurPlay - _
                sLrcShowHeight * (sTime - myLrc(iCurPlay).lrcTime) / iCurTimer    '���¾��У�����ȥ���䲥�ŵĲ���
    lblCurLrc.Caption = myLrc(iCurPlay).lrcString     '�õ���ǰ���
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

'�������Ϣ���ص� lblEachLrc() �ؼ�
Public Function LoadLrcLabel()
    Dim i           As Integer
On Error Resume Next
    For i = 0 To iLrcRows
        If i Then
            Load lblEachLrc(i)
        End If
        lblEachLrc(i).BackStyle = 0             '͸��
        lblEachLrc(i).Alignment = 2             '����
        lblEachLrc(i).UseMnemonic = False       '��ת�� & Ϊ�»���
        lblEachLrc(i).Caption = myLrc(i).lrcString
        lblEachLrc(i).AUTOSIZE = False
    Next i
    Call ResizeLrc
End Function

'�������ÿؼ���С
Private Sub ResizeLrc()
On Error Resume Next
    Dim i       As Integer
    Dim sWidth  As Single, sHeight  As Single
    Dim sMax    As Single
    
    With UserControl
        sWidth = 0
        sHeight = .TextHeight("Ag")
        For i = 0 To iLrcRows                       '�ҵ���ʵ������
            sMax = .TextWidth(myLrc(i).lrcString)
            If sMax > sWidth Then sWidth = sMax
        Next
'        sWidth = (sWidth \ Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
        lblCurLrc.Width = sWidth
        fraMask.Width = sWidth
        lblCurLrc.Alignment = 2
        fraLrc.Move (.Width - sWidth) / 2, .Height, sWidth, iLrcRows * sHeight             '������С��λ�� /�ڿؼ���ˮƽ����,��ֱ����ʾ
    End With
    For i = 0 To iLrcRows                       '����λ�õ�..
        lblEachLrc(i).Move 0, i * sHeight, sWidth                   '���õȿ�(ʵ��Ч���Ǿ�����ʾ)
        lblEachLrc(i).Visible = True                                '��ʾ
    Next
    
    lRet = SetInitEntry("PLAYER", "LRC_SIZE", DE_SIZE)
End Sub

'����ɫ
Property Let BackColor(Color As Long)
    UserControl.BackColor = Color
    fraLrc.BackColor = Color
    fraMask.BackColor = Color
End Property
'����ɫ
Property Get BackColor() As Long
    BackColor = UserControl.BackColor
End Property
'����ǰ��ɫ
Property Let FOREColor(Color As Long)
    Dim i As Integer
    For i = 0 To lblEachLrc.UBound
        lblEachLrc(i).FOREColor = Color
    Next i
End Property
'�õ�ǰ��ɫ
Property Get FOREColor() As Long
    FOREColor = lblEachLrc(0).FOREColor
End Property

'�õ�ÿ�еĸ߶�
Property Get LineHeight() As Long
    LineHeight = lblEachLrc(0).Height
End Property
'�õ���ǰ�ƶ��е�Width
Public Function GetCurrWidth(Index As Integer) As Single
    On Error Resume Next
    GetCurrWidth = lblEachLrc(Index).Width
End Function
'�õ���ǰ�ƶ��е�Width
Public Function GetLineTop(Index As Integer) As Single
    On Error Resume Next
    GetLineTop = lblEachLrc(Index).Top
End Function
'�õ���ǰ�ƶ��е�Width
Public Function GetLineLeft(Index As Integer) As Single
    On Error Resume Next
    GetLineLeft = lblEachLrc(Index).Left
End Function
'�õ���ǰ�ĸ������
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
