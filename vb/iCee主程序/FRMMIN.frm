VERSION 5.00
Begin VB.Form FRMMIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00241D0A&
   BorderStyle     =   0  'None
   Caption         =   "文件信息"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   ForeColor       =   &H00000000&
   Icon            =   "FRMMIN.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PERR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00241D0A&
      BorderStyle     =   0  'None
      Height          =   8460
      Left            =   7080
      ScaleHeight     =   564
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   518
      TabIndex        =   16
      Top             =   8400
      Visible         =   0   'False
      Width           =   7770
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   8640
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   1440
      Picture         =   "FRMMIN.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2280
      Picture         =   "FRMMIN.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   240
      Picture         =   "FRMMIN.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox TXTPATH 
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   8280
      Width           =   7335
   End
   Begin VB.PictureBox PO 
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   3
      Left            =   600
      ScaleHeight     =   3255
      ScaleWidth      =   3495
      TabIndex        =   9
      Top             =   4440
      Width           =   3495
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label LBTS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件信息"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   2
      Left            =   4080
      ScaleHeight     =   3255
      ScaleWidth      =   3615
      TabIndex        =   7
      Top             =   4440
      Width           =   3615
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   120
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label LBTS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件信息"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   1
      Left            =   4200
      ScaleHeight     =   3015
      ScaleWidth      =   3495
      TabIndex        =   5
      Top             =   1320
      Width           =   3495
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   120
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label LBTS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件信息"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox PO 
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   0
      Left            =   600
      ScaleHeight     =   3015
      ScaleWidth      =   3735
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Text            =   "无"
         Top             =   2445
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   120
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label LBTS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件信息"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   720
      End
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   21
      Top             =   8640
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   22
      Top             =   8640
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   3
      Left            =   3960
      TabIndex        =   23
      Top             =   8640
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
   End
   Begin VB.Image IU 
      Height          =   705
      Left            =   7185
      Picture         =   "FRMMIN.frx":0636
      ToolTipText     =   "关闭"
      Top             =   15
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "歌名"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "专辑"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   4800
      Width           =   360
   End
End
Attribute VB_Name = "FRMMIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Type Mp3tag    'mp3的ID31结构
  Title(29) As Byte    '标题
  Artist(29) As Byte   '艺员
  Album(29) As Byte    '专辑
  Year(3) As Byte      '年代
  Comment(29) As Byte  '注释
  Genre As Byte        '风格
End Type

Private Type ID3Header 'mp3的ID3头结构
  id As String * 3
  Version As Integer   '版本
  flag As Byte         '标志
  Size(3) As Byte      '大小
End Type

Private Type wmaExtend  'wma扩展标签结构
  ObjectID(15) As Byte  '对象ID
  ObjectSize As Long    '对象大小
  vain As Long          '空字节
  fSum As Integer       '帧总数
End Type

Private Type wmaContent 'wma标准标签结构
  ObjectID(15) As Byte  '对象ID
  ObjectSize As Long    '对象大小
  vain As Long          '空字节
  L(4) As Integer       '项长度
End Type

Private Const tag1ID = "3326B2758E66CF11A6D900AA0062CE6C" 'wma标准标签对象ID
Private Const tag2ID = "40A4D0D207E3D21197F000A0C95EA850" 'wma扩展标签对象ID

Dim WithEvents CD As VBControlExtender
Attribute CD.VB_VarHelpID = -1
Dim OpenName As String, SaveName As String
Dim audioData() As Byte   '音频文件数据
Dim bjplay As Boolean     '播放标记
Dim bjTag1 As Boolean     'mp3的ID3V1或wma的标准标签写盘标记
Dim bjTag2 As Boolean     'mp3的ID3V2或wma的扩展标签写盘标记
Dim bjType1 As Boolean    '音频类型标记.1-mp3，0-wma
Dim bjType2 As Boolean    '同上

Dim Wm(7) As String       'wma扩展标签的帧名称
Dim wmaHeader(29) As Byte 'wma头数据
Dim HeaderLen As Long     'wma顶级头对象大小
Dim ObjectSum As Byte     'wma顶级头对象中的子对象数量

Dim ID3V2Info() As Byte   'mp3的ID3V2信息

Private Sub Form_Activate()
Me.Cls
Me.BackColor = COLOR_NOR
Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is PictureBox Then
PBOX.Cls
PBOX.Refresh
PBOX.BackColor = Me.BackColor
End If
If TypeOf PBOX Is TextBox Then PBOX.BackColor = Me.BackColor
Next

Call PaintPng(App.Path & "\SKIN\CTS.PNG", PERR.hdc, 160, 216)
Dim i As Integer
For i = 0 To ICM.Count - 1
ICM(i).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\S_T.PNG", Me.hdc, 8, 8)

End Sub

Private Sub Form_Load()
PERR.Move 8, 56

ICM(0).SETTXT "保存文件信息"
ICM(1).SETTXT "关闭"
ICM(2).SETTXT "打开文件位置"
ICM(3).SETTXT "分享音乐文件"

Dim i As Integer, st() As String, Z As String, MYLF, MYTOP
MYLF = GetInitEntry("INFO", "LEFT", (Screen.Width - Me.Width) / 2)
MYTOP = GetInitEntry("INFO", "TOP", (Screen.Height - Me.Height) / 2)
Me.ScaleMode = 1

Call SeekMe(Me)

For i = 1 To 6: Load Label1(i): Label1(i).Move 180, 350 * i + 1680: Label1(i).Visible = True: Next
For i = 1 To 7: Load Label2(i): Label2(i).Move 180, 350 * i + 4800: Label2(i).Visible = True: Next
For i = 1 To 5: Load Text1(i): Text1(i).Move 120, 345 * i + 360: Text1(i).Visible = True: Next
For i = 1 To 7: Load Text2(i): Text2(i).Move 120, 345 * i + 360: Text2(i).Visible = True: Next
For i = 1 To 6: Load Text3(i): Text3(i).Move 120, 345 * i + 360: Text3(i).Visible = True: Next
For i = 1 To 7: Load Text4(i): Text4(i).Move 120, 345 * i + 360: Text4(i).Visible = True: Next

设置 0
st = Split("WM/AlbumTitle|WM/Track|WM/TrackNumber|WM/AlbumArtist|WM/Writer|WM/Composer|WM/Year|WM/Mood", "|")
For i = 0 To 7: Wm(i) = st(i): Next '0专辑,1曲号,2曲数,3歌手,4作词,5作曲,6日期,7气氛

Z = "无|布鲁斯|古典摇滚|乡村|舞曲|迪斯科|伤感爵士|垃圾摇滚|饶舌|爵士|金属|前卫|怀旧|其他|流行|" & _
  "摇滚布鲁斯|说唱|雷盖扭摆舞|摇滚|电子流行乐|工业|非主流|斯卡|重金属|恶作剧|电影配乐|现代电子乐|" & _
  "环绕|轻音乐|口头|电贝司爵士|合成音乐|迷幻舞曲|古典|器乐|讽刺|电子舞曲|野性|声音剪辑|福音|喧闹|" & _
  "非主流摇滚|低音|灵歌|颓废|空间|沉思|流行器乐|器乐摇滚|民族|粗鄙|暗潮|现代电子|电子|民歌|欧洲舞曲|" & _
  "梦幻|南部摇滚|喜剧|狂热|冈斯特说唱|顶峰40 |基督教说唱|流行摇滚|丛林|美国本土|酒馆|新wav|特色音乐|" & _
  "狂欢|演出收听|广告|低保真|面具舞|讽刺颓废|讽刺爵士|波尔卡|复古|喜剧|摇滚舞|硬摇滚|民间|民俗摇滚|" & _
  "民乐|摇摆|快速融合|比博普爵士乐|拉丁舞|复兴|凯尔特|蓝色牧场|激进|哥特摇滚|前卫摇滚|迷幻摇滚|交响摇滚|" & _
  "慢摇滚|爵士乐团|合唱|轻松的流行乐|原声|幽默|演说|小调|歌剧|室内乐|奏鸣曲|交响曲|低音波特|苏格兰|" & _
  "色情音乐|讽刺|慢即兴演奏|俱乐部|探戈|桑巴|民俗|小曲|豆瓣民谣|节奏灵歌|自由式|二重唱|庞克摇滚鼓|" & _
  "鼓独奏|无伴奏|欧式电子舞曲|舞厅|果阿|低音鼓|俱乐部电子舞|硬核摇滚|惊悚|独立音乐|英式摇滚|黑人庞克|" & _
  "波兰庞克|踢踏舞|基督教黑帮說唱|重金属摇滚|黑金属摇滚|跨界音乐|钢琴|基督教摇滚|梅伦格舞|" & _
  "莎莎拉丁|废金属|动漫|日本流行音乐|电子合成器流行音乐"
st = Split(Z, "|")
For i = 0 To UBound(st)
Combo1.AddItem st(i)
Next

End Sub

Private Sub 设置(Index As Integer)
Dim i As Integer, st1() As String, st2() As String, Z As String
Select Case Index
  Case 1
    LBTS(0).Caption = "ID3V1 信息"
    LBTS(2).Caption = "ID3V2 信息"
    LBTS(3).Caption = "ID3V1 备用"
    LBTS(4).Caption = "ID3V2 备用"
    
    st1 = Split("歌名|歌手|专辑|曲号|年份|注释|风格", "|"): st2 = Split("歌名|歌手|专辑|曲号|日期|注释|风格|备注", "|")
    Z = "长度不超过30字节"
    For i = 0 To 5: Text1(i) = "": Text1(i).ToolTipText = Z: Text3(i).ToolTipText = Z: Next
    Z = "数值为1—255"
    Text1(3).ToolTipText = Z: Text3(3).ToolTipText = Z
    Z = "长度不超过4字节"
    Text1(4).ToolTipText = Z: Text3(4).ToolTipText = Z
    For i = 0 To 7: Text2(i) = "": Next
  Case 0
    LBTS(0).Caption = "标准标签信息"
    LBTS(2).Caption = "扩展标签信息"
    LBTS(3).Caption = "标准备用"
    LBTS(4).Caption = "扩展备用"
    
    st1 = Split("歌名|歌手|版权|注释|风格|无效|列表", "|"): st2 = Split("专辑|曲号|曲数|歌手|作词|作曲|日期|气氛", "|")
    For i = 0 To 5: Text1(i) = "": Text1(i).ToolTipText = "": Text3(i).ToolTipText = "": Next
    For i = 0 To 7: Text2(i) = "": Next
End Select
For i = 0 To 6: Label1(i) = st1(i): Next
For i = 0 To 7: Label2(i) = st2(i): Next
bjType2 = bjType1
End Sub

Sub SeeIt(filename As String)  '打开
On Error GoTo 100
OpenName = filename
txtPath.Text = filename
If UCase(Split(txtPath.Text, ":")(0)) = "HTTP" Then PERR.Visible = True: Exit Sub
列表框处理
100
End Sub
Private Sub 列表框处理()
Dim Z As String, i As Integer
bjType1 = (LCase(Right(OpenName, 3)) = "mp3")
If bjType1 <> bjType2 Then 设置 Abs(bjType1)
Z = Dir(OpenName)
If bjType1 Then mp3信息处理 Else wma信息处理
End Sub

Private Sub mp3信息处理()
On Error GoTo 100
Dim ID3v As String * 3, L1 As Byte, L2 As Byte, L3 As Byte, ID3Len As Long
Dim ID3V1Info As Mp3tag, i As Integer, FileLen As Long
For i = 0 To 5: Text1(i) = "": Next
For i = 0 To 7: Text2(i) = "": Next
Caption = Dir(OpenName): Text3(0) = Left(Caption, Len(Caption) - 4)
bjTag2 = False: bjTag1 = False

Open OpenName For Binary As #1
FileLen = LOF(1)

Get #1, FileLen - 127, ID3v
If ID3v = "TAG" Then '如果有ID3V1
  bjTag1 = True
  Get #1, , ID3V1Info
End If

Get #1, 1, ID3v
If ID3v = "ID3" Then '如果有ID3V2
  bjTag2 = True
  Get #1, 8, L1
  Get #1, , L2
  Get #1, , L3
  ID3Len = L1
  ID3Len = ID3Len * &H4000 + L2 * &H80 + L3
  ReDim ID3V2Info(ID3Len - 1)
  Get #1, , ID3V2Info
End If

ReDim audioData(FileLen + bjTag1 * 128 + bjTag2 * (ID3Len + 10) - 1)
If bjTag2 Then
  Get #1, , audioData
Else
  Get #1, 1, audioData
End If

If bjTag1 Then 获取ID3V1信息 ID3V1Info
If bjTag2 Then 获取ID3V2信息
PERR.Visible = False
100
Close #1

If ERR.Number > 0 Then PERR.Visible = True: Call SHOWWRONG("读入文件时出错,错误号:" & ERR.Number, 2)
End Sub

Private Sub 获取ID3V1信息(ID3V1 As Mp3tag)
With ID3V1
  ID3V1处理 .Title, 0   '歌名
  ID3V1处理 .Artist, 1  '艺员
  ID3V1处理 .Album, 2   '专辑
  If .Comment(28) = 0 And .Comment(29) > 0 And Len(Text1(2)) > 0 Then Text1(3) = .Comment(29): .Comment(29) = 0
  ID3V1处理 .Comment, 5 '注释
  Text1(4) = StrConv(.Year, vbUnicode)
  If .Genre < 149 Then Combo1.ListIndex = .Genre + 1
End With
End Sub

Private Sub ID3V1处理(tem() As Byte, K As Integer)
If IsTextUTF8(tem) Then
  Text1(K) = UTF_8ToTxt(tem)
Else
  Text1(K) = StrConv(tem, vbUnicode)
End If
End Sub

Private Sub 获取ID3V2信息()
ID3V2处理 "TIT2", 0 '歌名
ID3V2处理 "TPE1", 1 '艺员
ID3V2处理 "TALB", 2 '专辑
ID3V2处理 "COMM", 5 '注释
ID3V2处理 "TYER", 4 '年份
ID3V2处理 "TRCK", 3 '曲号
ID3V2处理 "TXXX", 7 '用户文本
ID3V2处理 "TCON", 6 '风格
If Len(Text2(6)) > 0 Then
  If InStr("( （", Left(Text2(6), 1)) > 0 Then
    Dim K As Integer
    K = Val(Mid(Text2(6), 2))
    If K < 149 Then Text2(6) = Combo1.List(K + 1)
  Else
    If Val(Text2(6)) Then Text2(6) = Combo1.List(Val(Text2(6)) + 1)
  End If
End If
End Sub

Private Sub ID3V2处理(st As String, K As Integer)
Dim Length As Integer, Place As Long, p As Long, i As Long, tem() As Byte, bj As Boolean
Text2(K).ToolTipText = ""
tem = StrConv(st, vbFromUnicode)
p = InStrB(ID3V2Info, tem)
If p > 0 Then
  Length = ID3V2Info(p + 5) * &H80 + ID3V2Info(p + 6) - 1: If Length < 1 Then Exit Sub
  Place = p + 9
  If ID3V2Info(Place) = 1 Then Place = Place + 3: Length = Length - 3: bj = True: If Length < 1 Then Exit Sub 'UTF-16LE编码(Unicode编码)
  ReDim tem(Length)
  For i = Place To Place + Length: tem(i - Place) = ID3V2Info(i): Next
  If bj Then
    Text2(K) = tem
  Else
    If IsTextUTF8(tem) Then 'UTF-8编码
      Text2(K) = Replace(UTF_8ToTxt(tem), Chr(0), "")
    Else
      Text2(K) = Replace(StrConv(tem, vbUnicode), Chr(0), "")
    End If
  End If
  Text2(K).ToolTipText = Text2(K)
End If
End Sub

Private Sub wma信息处理()
On Error GoTo 100
Dim i As Long, K As Long
Caption = Dir(OpenName): Text3(0) = Left(Caption, Len(Caption) - 4)

Open OpenName For Binary As #1
ReDim audioData(LOF(1) - 1)
Get #1, , audioData
Close #1

HeaderLen = audioData(16) + audioData(17) * 256 + audioData(18) * 65536 '计算顶级头对象大小
获取标准标签信息
获取扩展标签信息

ObjectSum = audioData(24) '获取对象数量
K = UBound(audioData)
For i = 0 To 29: wmaHeader(i) = audioData(i): Next '分离出前30字节
For i = 0 To K - 30: audioData(i) = audioData(i + 30): Next '数据前移
ReDim Preserve audioData(K - 30)

Exit Sub
100
Close
End Sub

Private Sub 获取标准标签信息()
On Error GoTo 100
Dim ObjectID(15) As Byte, i As Integer, k1 As Long, k2 As Long, k3 As Long
Dim Ltag(4) As Integer

For i = 0 To 15: ObjectID(i) = Val("&H" & Mid(tag1ID, i * 2 + 1, 2)): Next
For i = 0 To 4: Text1(i) = "": Text1(i).ToolTipText = "": Text3(i).ToolTipText = "": Next

k2 = InStrB(audioData, ObjectID)
If k2 > 0 Then '如果有标准标签
  k3 = k2 - 1  'k3是对象ID的起始位置
  k2 = k2 + 23: k1 = k2
  For i = 0 To 4
    Ltag(i) = audioData(k1 + i * 2) + audioData(k1 + i * 2 + 1) * 256 '获取各项长度
    标准标签信息处理 Ltag(i) - 2, k2 + 10, i
    k2 = k2 + Ltag(i)
  Next
  k1 = audioData(k3 + 16) + audioData(k3 + 17) * 256 '获取标准标签的大小
  For k2 = k3 To UBound(audioData) - k1: audioData(k2) = audioData(k2 + k1): Next '数据前移
  ReDim Preserve audioData(UBound(audioData) - k1)   '从原数据中去掉标准标签
  HeaderLen = HeaderLen - k1 '计算去掉标准标签后的顶级头对象大小
  audioData(24) = audioData(24) - 1 '计算去掉标准标签后的对象数量
End If

100
End Sub

Private Sub 标准标签信息处理(S1 As Integer, S2 As Long, n As Integer) 's1-项长度；s2-项位置；n-文本框编号
Dim j As Integer, i As Long, tem() As Byte
If S1 > 1 Then
  ReDim tem(S1 - 1)
  For i = S2 To S2 + S1 - 1: tem(j) = audioData(i): j = j + 1: Next
  Text1(n) = tem
End If
End Sub

Private Sub 获取扩展标签信息()
On Error GoTo 100
Dim ObjectID(15) As Byte, i As Integer, k1 As Long, k2 As Long, k3 As Long

For i = 0 To 15: ObjectID(i) = Val("&H" & Mid(tag2ID, i * 2 + 1, 2)): Next

k2 = InStrB(audioData, ObjectID)
If k2 > 0 Then '如果有标准标签
  k3 = k2 - 1  'k3是对象ID的起始位置
  For i = 0 To 7
    Text2(i) = ""
    扩展标签信息处理 i
  Next
  k1 = audioData(k3 + 16) + audioData(k3 + 17) * 256 '获取扩展标签的大小
  For k2 = k3 To UBound(audioData) - k1: audioData(k2) = audioData(k2 + k1): Next '数据前移
  ReDim Preserve audioData(UBound(audioData) - k1)   '从原数据中去掉扩展标签
  HeaderLen = HeaderLen - k1 '计算去掉扩展标签后的顶级头对象大小
  audioData(24) = audioData(24) - 1 '计算去掉扩展标签后的对象数量
End If

100
End Sub

Private Sub 扩展标签信息处理(n As Integer) 'n-文本框编号
On Error GoTo 100
Dim tem1() As Byte, tem2() As Byte, j As Integer, K As Long, i As Long, L1 As Long, L2 As Long
tem1 = Wm(n)
K = InStrB(audioData, tem1)
If K > 0 Then '如果有这个帧
  L1 = audioData(K - 3) + audioData(K - 4) * 256 '帧名称长度
  L2 = audioData(K + L1 + 1) + audioData(K + L1 + 2) * 256 '帧内容长度
  If L2 > 3 Then
    L1 = L1 + K + 3 '帧内容起始字节
    ReDim tem2(L2 - 3)
    For i = L1 To L1 + L2 - 3: tem2(j) = audioData(i): j = j + 1: Next '取出帧内容，同时去掉字符串最后的2个空字符
    Text2(n) = tem2
  End If
End If
100
End Sub

Private Sub 保存() '保存
On Error GoTo 100
Dim st1 As String, st2 As String
If Len(SaveName) > 6 Then st1 = Left(SaveName, InStrRev(SaveName, "\")) & Dir(OpenName) Else st1 = OpenName
st2 = "保存全部信息(*.mp3)" & Chr(0) & "mp3"
SaveName = OpenName
Call saveMP3
Me.Hide
Call FRMMIN.SeeIt(frmma.PLIST.URL(frmma.PLIST.ListIndex))
100
End Sub
Private Function 写入标准标签信息() As Integer
On Error GoTo 100
Dim t1 As wmaContent, st As String, i As Integer, tem() As Byte
Dim s As String

With t1
  For i = 0 To 4
    If Len(Text1(i)) > 0 Then .L(i) = LenB(Text1(i)) + 2
  Next
  .ObjectSize = .L(0) + .L(1) + .L(2) + .L(3) + .L(4) + 34
  For i = 0 To 15: .ObjectID(i) = Val("&H" & Mid(tag1ID, i * 2 + 1, 2)): Next
End With

Open SaveName For Binary As #1
Seek #1, LOF(1) + 1
Put #1, , t1
For i = 0 To 4
  If Len(Text1(i)) > 0 Then
    tem = Text1(i) & Chr(0)
    Put #1, , tem
  End If
Next
100
Close #1
写入标准标签信息 = ERR.Number
End Function

Private Function 写入扩展标签信息() As Integer
On Error GoTo 100
Dim t2 As wmaExtend, i As Integer, m(7) As Integer, n(7) As Integer, tem1() As Byte, tem2() As Byte

With t2
  For i = 0 To 7
    If Len(Text2(i)) > 0 Then
      m(i) = LenB(Wm(i)) + 2
      n(i) = LenB(Text2(i)) + 2
      .ObjectSize = .ObjectSize + m(i) + n(i) + 6
      .fSum = .fSum + 1
    End If
  Next
  For i = 0 To 15: .ObjectID(i) = Val("&H" & Mid(tag2ID, i * 2 + 1, 2)): Next
 .ObjectSize = .ObjectSize + 26
End With

Open SaveName For Binary As #1
Seek #1, LOF(1) + 1
Put #1, , t2
For i = 0 To 7
  If Len(Text2(i)) > 0 Then
    tem1 = Wm(i) & String(2, 0)
    tem2 = Text2(i) & Chr(0)
    Put #1, , m(i)
    Put #1, , tem1
    Put #1, , n(i)
    Put #1, , tem2
  End If
Next

100
Close #1
写入扩展标签信息 = ERR.Number
End Function

Private Sub saveMP3() '保存mp3
On Error GoTo 100
Dim i As Integer, K As Long
GoSub 200: GoSub 300: bjTag1 = True: bjTag2 = True '写入全部
i = 0
If bjTag2 Then i = 写入ID3V2
If i = 0 Then
  Open SaveName For Binary As #1
  If bjTag2 Then Seek #1, LOF(1) + 1
  Put #1, , audioData
  Close #1
End If
If bjTag1 Then If i = 0 Then i = 写入ID3V1
If i = 0 Then Debug.Print "保存成功" Else Debug.Print "保存失败,错误号＝" & i
100
Exit Sub
200
i = examine: If i > 0 Then Call SHOWWRONG("ID3V1信息中第" & i & "个文本框字数超出", 0): Exit Sub
Return
300
For i = 0 To 7: K = K + Len(Text2(i)): Next
If K = 0 Then Call SHOWWRONG("要写入ID3V2信息,不能所有文本框都为空!", 0): Exit Sub
Return
End Sub

Private Function examine() As Integer
If lstrlen(Text1(0)) > 30 Then examine = 1: Exit Function
If lstrlen(Text1(1)) > 30 Then examine = 2: Exit Function
If lstrlen(Text1(2)) > 30 Then examine = 3: Exit Function
If lstrlen(Text1(5)) > 30 Then examine = 6: Exit Function
If lstrlen(Text1(4)) > 4 Then examine = 5
End Function

Private Function 写入ID3V1() As Integer
On Error GoTo 100
Dim ID3V1Info As Mp3tag, i As Integer, tem() As Byte, Tag As String * 3
Tag = "TAG"

With ID3V1Info
  tem = StrConv(Text1(0), vbFromUnicode)
  For i = 0 To UBound(tem): .Title(i) = tem(i): Next   '歌名

  tem = StrConv(Text1(1), vbFromUnicode)
  For i = 0 To UBound(tem): .Artist(i) = tem(i): Next  '艺员

  tem = StrConv(Text1(2), vbFromUnicode)
  For i = 0 To UBound(tem): .Album(i) = tem(i): Next   '专辑

  tem = StrConv(Text1(5), vbFromUnicode)
  For i = 0 To UBound(tem): .Comment(i) = tem(i): Next '注释

  tem = StrConv(Left(Text1(4) & String(4, 0), 4), vbFromUnicode) '年份
  For i = 0 To 3: .Year(i) = tem(i): Next

  i = Val(Text1(3)): If i > 255 Then i = 255
  If Len(Text1(2)) > 0 And i > 0 Then .Comment(28) = 0: .Comment(29) = i '曲号

  For i = 0 To Combo1.ListCount - 1 '风格
    If Combo1.List(i) = Combo1.Text Then Exit For
  Next
  If i = 0 Or i = Combo1.ListCount Then i = 256
  .Genre = i - 1

End With
'Debug.Print I, Tag
Open SaveName For Binary As #1
Seek #1, LOF(1) + 1
Put #1, , Tag
Put #1, , ID3V1Info
100
Close #1
写入ID3V1 = ERR.Number
End Function

Private Function 写入ID3V2() As Integer
On Error GoTo 100
Dim ID3V2 As ID3Header
Dim FrameID() As String '帧标识符
Dim Size(3) As Byte     '帧内容长度
Dim flags As Integer    '标志
Dim Data As String      '帧内容
Dim L(7) As Integer, v2Len As Integer, i As Integer, s As String

s = "TIT2|TPE1|TALB|TRCK|TYER|COMM|TCON|TXXX"
FrameID = Split(s, "|")

For i = 0 To 7
 If Len(Text2(i)) > 0 Then L(i) = lstrlen(Text2(i)) + 1: v2Len = v2Len + L(i) + 10
Next

With ID3V2
  .id = "ID3"
  .Version = 3
  .Size(2) = v2Len \ 128
  .Size(3) = v2Len Mod 128
End With

Open SaveName For Binary As #1
Put #1, , ID3V2

For i = 0 To 7
  If L(i) > 0 Then
    Size(2) = L(i) \ 128
    Size(3) = L(i) Mod 128
    Data = Chr(0) & Text2(i)
    
    Put #1, , FrameID(i)
    Put #1, , Size
    Put #1, , flags
    Put #1, , Data
  End If
Next

100
Close #1
写入ID3V2 = ERR.Number
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE <> Me.X1.PICTURE Then IU.PICTURE = Me.X1.PICTURE

End Sub

Private Sub Form_Unload(Cancel As Integer)
lRet = SetInitEntry("INFO", "LEFT", Me.Left)
lRet = SetInitEntry("INFO", "TOP", Me.Top)
End Sub

Private Sub ICM_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Call 保存
Case 1
Me.Hide
Case 2
If Dir(txtPath.Text) = "" Then Exit Sub
Shell "explorer.exe /select," & txtPath.Text, vbNormalFocus
Case 3
Call frmma.SHAREIT(txtPath.Text)
End Select
End Sub

Private Sub IU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IU.PICTURE = Me.X2.PICTURE Then IU.PICTURE = Me.X3.PICTURE
End Sub
Private Sub IU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE = Me.X1.PICTURE Then IU.PICTURE = Me.X2.PICTURE
End Sub
Private Sub IU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IU.PICTURE = Me.X3.PICTURE Then IU.PICTURE = Me.X1.PICTURE
Me.Hide
End Sub


Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LBTS_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PERR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 2 Then If Len(Text1(Index)) = 0 Then Text1(Index) = Clipboard.GetText Else Clipboard.SetText Text1(Index)
End Sub

Private Sub Text2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 2 Then If Len(Text2(Index)) = 0 Then Text2(Index) = Clipboard.GetText Else Clipboard.SetText Text2(Index)
End Sub

Private Sub Text3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 2 Then If Len(Text3(Index)) = 0 Then Text3(Index) = Clipboard.GetText Else Clipboard.SetText Text3(Index)
End Sub

Private Sub Text4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 2 Then If Len(Text3(Index)) = 0 Then Text3(Index) = Clipboard.GetText Else Clipboard.SetText Text3(Index)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0: Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text2_GotFocus(Index As Integer)
Text2(Index).SelStart = 0: Text2(Index).SelLength = Len(Text2(Index))
End Sub

Private Sub Text3_GotFocus(Index As Integer)
Text3(Index).SelStart = 0: Text3(Index).SelLength = Len(Text3(Index))
End Sub

Private Sub Text4_GotFocus(Index As Integer)
Text4(Index).SelStart = 0: Text4(Index).SelLength = Len(Text4(Index))
End Sub

Private Sub Text4_DblClick(Index As Integer)
If bjType1 Then
  If Index = 4 Then Text4(4) = Date & " " & WeekdayName(WeekDay(Date, 1)) & " " & TimE: Text2(4) = Text4(4): Text1(4) = Left(Text4(4), 4): Text3(4) = Text1(4)
Else
  If Index = 6 Then Text4(6) = Date & " " & WeekdayName(WeekDay(Date, 1)) & " " & TimE: Text2(6) = Text4(6)
End If
End Sub

Private Sub Combo1_Click()
If bjType1 Then Text3(6) = Combo1.Text: Text4(6) = Text3(6): Text2(6) = Text3(6) Else Text1(4) = Combo1.Text: Text3(4) = Text1(4)
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 100
OpenName = Data.files.Item(1)
列表框处理
100
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error GoTo 100
If InStr("mp3,wma", LCase(Right(Data.files.Item(1), 3))) Then Effect = vbDropEffectCopy And Effect Else Effect = vbDropEffectNone
100
End Sub

Private Function UTF_8ToTxt(bytSrc() As Byte) As String 'UTF_8编码转换为普通文本
On Error GoTo 100
Dim tem() As Byte, L As Integer, K As Integer, i As Integer
K = UBound(bytSrc)
ReDim tem(K * 2) As Byte
For i = 0 To K
  If bytSrc(i) < 128 Then
    tem(L) = bytSrc(i)
  Else
    tem(L + 1) = ((bytSrc(i) And 15) * 16 + (bytSrc(i + 1) And 60) / 4)
    tem(L) = (bytSrc(i + 1) And 3) * 64 + (bytSrc(i + 2) And 63)
    i = i + 2
  End If
  L = L + 2
Next
ReDim Preserve tem(L - 1) As Byte
UTF_8ToTxt = tem
100
End Function

Private Function IsTextUTF8(bytSrc() As Byte) As Boolean '判断是否UTF-8编码
Dim i As Integer, AscN As Integer, n As Integer
n = UBound(bytSrc)

Do While i <= n
  If bytSrc(i) < 128 Then 'Ascii字符
    i = i + 1: AscN = AscN + 1
  ElseIf (bytSrc(i) And &HF0) = &HE0 Then '3个字节的UTF-8
    If (bytSrc(i + 1) And &HC0) = &H80 Then
      If (bytSrc(i + 2) And &HC0) = &H80 Or (bytSrc(i + 2) And &HC0) = 0 Then i = i + 3 Else Exit Function
    Else
      Exit Function
    End If
  Else
    Exit Function
  End If
Loop
IsTextUTF8 = (AscN <> n + 1)
End Function

Private Sub txtPath_Change()
On Error Resume Next
If PathFileExists(txtPath.Text) = 0 Then PERR.Visible = True
End Sub
