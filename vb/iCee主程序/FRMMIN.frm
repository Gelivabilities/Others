VERSION 5.00
Begin VB.Form FRMMIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00241D0A&
   BorderStyle     =   0  'None
   Caption         =   "�ļ���Ϣ"
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
   StartUpPosition =   3  '����ȱʡ
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
            Name            =   "����"
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
         Caption         =   "�ļ���Ϣ"
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
            Name            =   "����"
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
         Caption         =   "�ļ���Ϣ"
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
            Name            =   "����"
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
         Caption         =   "�ļ���Ϣ"
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
         Text            =   "��"
         Top             =   2445
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ļ���Ϣ"
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
      ToolTipText     =   "�ر�"
      Top             =   15
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
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
      Caption         =   "ר��"
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

Private Type Mp3tag    'mp3��ID31�ṹ
  Title(29) As Byte    '����
  Artist(29) As Byte   '��Ա
  Album(29) As Byte    'ר��
  Year(3) As Byte      '���
  Comment(29) As Byte  'ע��
  Genre As Byte        '���
End Type

Private Type ID3Header 'mp3��ID3ͷ�ṹ
  id As String * 3
  Version As Integer   '�汾
  flag As Byte         '��־
  Size(3) As Byte      '��С
End Type

Private Type wmaExtend  'wma��չ��ǩ�ṹ
  ObjectID(15) As Byte  '����ID
  ObjectSize As Long    '�����С
  vain As Long          '���ֽ�
  fSum As Integer       '֡����
End Type

Private Type wmaContent 'wma��׼��ǩ�ṹ
  ObjectID(15) As Byte  '����ID
  ObjectSize As Long    '�����С
  vain As Long          '���ֽ�
  L(4) As Integer       '���
End Type

Private Const tag1ID = "3326B2758E66CF11A6D900AA0062CE6C" 'wma��׼��ǩ����ID
Private Const tag2ID = "40A4D0D207E3D21197F000A0C95EA850" 'wma��չ��ǩ����ID

Dim WithEvents CD As VBControlExtender
Attribute CD.VB_VarHelpID = -1
Dim OpenName As String, SaveName As String
Dim audioData() As Byte   '��Ƶ�ļ�����
Dim bjplay As Boolean     '���ű��
Dim bjTag1 As Boolean     'mp3��ID3V1��wma�ı�׼��ǩд�̱��
Dim bjTag2 As Boolean     'mp3��ID3V2��wma����չ��ǩд�̱��
Dim bjType1 As Boolean    '��Ƶ���ͱ��.1-mp3��0-wma
Dim bjType2 As Boolean    'ͬ��

Dim Wm(7) As String       'wma��չ��ǩ��֡����
Dim wmaHeader(29) As Byte 'wmaͷ����
Dim HeaderLen As Long     'wma����ͷ�����С
Dim ObjectSum As Byte     'wma����ͷ�����е��Ӷ�������

Dim ID3V2Info() As Byte   'mp3��ID3V2��Ϣ

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

ICM(0).SETTXT "�����ļ���Ϣ"
ICM(1).SETTXT "�ر�"
ICM(2).SETTXT "���ļ�λ��"
ICM(3).SETTXT "���������ļ�"

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

���� 0
st = Split("WM/AlbumTitle|WM/Track|WM/TrackNumber|WM/AlbumArtist|WM/Writer|WM/Composer|WM/Year|WM/Mood", "|")
For i = 0 To 7: Wm(i) = st(i): Next '0ר��,1����,2����,3����,4����,5����,6����,7����

Z = "��|��³˹|�ŵ�ҡ��|���|����|��˹��|�˸о�ʿ|����ҡ��|����|��ʿ|����|ǰ��|����|����|����|" & _
  "ҡ����³˹|˵��|�׸�Ť����|ҡ��|����������|��ҵ|������|˹��|�ؽ���|������|��Ӱ����|�ִ�������|" & _
  "����|������|��ͷ|�籴˾��ʿ|�ϳ�����|�Ի�����|�ŵ�|����|���|��������|Ұ��|��������|����|����|" & _
  "������ҡ��|����|���|�Ƿ�|�ռ�|��˼|��������|����ҡ��|����|�ֱ�|����|�ִ�����|����|���|ŷ������|" & _
  "�λ�|�ϲ�ҡ��|ϲ��|����|��˹��˵��|����40 |������˵��|����ҡ��|����|��������|�ƹ�|��wav|��ɫ����|" & _
  "��|�ݳ�����|���|�ͱ���|�����|����Ƿ�|��̾�ʿ|������|����|ϲ��|ҡ����|Ӳҡ��|���|����ҡ��|" & _
  "����|ҡ��|�����ں�|�Ȳ��վ�ʿ��|������|����|������|��ɫ����|����|����ҡ��|ǰ��ҡ��|�Ի�ҡ��|����ҡ��|" & _
  "��ҡ��|��ʿ����|�ϳ�|���ɵ�������|ԭ��|��Ĭ|��˵|С��|���|������|������|������|��������|�ո���|" & _
  "ɫ������|���|����������|���ֲ�|̽��|ɣ��|����|С��|������ҥ|�������|����ʽ|���س�|�ӿ�ҡ����|" & _
  "�Ķ���|�ް���|ŷʽ��������|����|����|������|���ֲ�������|Ӳ��ҡ��|���|��������|Ӣʽҡ��|�����ӿ�|" & _
  "�����ӿ�|��̤��|�����̺ڰ��f��|�ؽ���ҡ��|�ڽ���ҡ��|�������|����|������ҡ��|÷�׸���|" & _
  "ɯɯ����|�Ͻ���|����|�ձ���������|���Ӻϳ�����������"
st = Split(Z, "|")
For i = 0 To UBound(st)
Combo1.AddItem st(i)
Next

End Sub

Private Sub ����(Index As Integer)
Dim i As Integer, st1() As String, st2() As String, Z As String
Select Case Index
  Case 1
    LBTS(0).Caption = "ID3V1 ��Ϣ"
    LBTS(2).Caption = "ID3V2 ��Ϣ"
    LBTS(3).Caption = "ID3V1 ����"
    LBTS(4).Caption = "ID3V2 ����"
    
    st1 = Split("����|����|ר��|����|���|ע��|���", "|"): st2 = Split("����|����|ר��|����|����|ע��|���|��ע", "|")
    Z = "���Ȳ�����30�ֽ�"
    For i = 0 To 5: Text1(i) = "": Text1(i).ToolTipText = Z: Text3(i).ToolTipText = Z: Next
    Z = "��ֵΪ1��255"
    Text1(3).ToolTipText = Z: Text3(3).ToolTipText = Z
    Z = "���Ȳ�����4�ֽ�"
    Text1(4).ToolTipText = Z: Text3(4).ToolTipText = Z
    For i = 0 To 7: Text2(i) = "": Next
  Case 0
    LBTS(0).Caption = "��׼��ǩ��Ϣ"
    LBTS(2).Caption = "��չ��ǩ��Ϣ"
    LBTS(3).Caption = "��׼����"
    LBTS(4).Caption = "��չ����"
    
    st1 = Split("����|����|��Ȩ|ע��|���|��Ч|�б�", "|"): st2 = Split("ר��|����|����|����|����|����|����|����", "|")
    For i = 0 To 5: Text1(i) = "": Text1(i).ToolTipText = "": Text3(i).ToolTipText = "": Next
    For i = 0 To 7: Text2(i) = "": Next
End Select
For i = 0 To 6: Label1(i) = st1(i): Next
For i = 0 To 7: Label2(i) = st2(i): Next
bjType2 = bjType1
End Sub

Sub SeeIt(filename As String)  '��
On Error GoTo 100
OpenName = filename
txtPath.Text = filename
If UCase(Split(txtPath.Text, ":")(0)) = "HTTP" Then PERR.Visible = True: Exit Sub
�б����
100
End Sub
Private Sub �б����()
Dim Z As String, i As Integer
bjType1 = (LCase(Right(OpenName, 3)) = "mp3")
If bjType1 <> bjType2 Then ���� Abs(bjType1)
Z = Dir(OpenName)
If bjType1 Then mp3��Ϣ���� Else wma��Ϣ����
End Sub

Private Sub mp3��Ϣ����()
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
If ID3v = "TAG" Then '�����ID3V1
  bjTag1 = True
  Get #1, , ID3V1Info
End If

Get #1, 1, ID3v
If ID3v = "ID3" Then '�����ID3V2
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

If bjTag1 Then ��ȡID3V1��Ϣ ID3V1Info
If bjTag2 Then ��ȡID3V2��Ϣ
PERR.Visible = False
100
Close #1

If ERR.Number > 0 Then PERR.Visible = True: Call SHOWWRONG("�����ļ�ʱ����,�����:" & ERR.Number, 2)
End Sub

Private Sub ��ȡID3V1��Ϣ(ID3V1 As Mp3tag)
With ID3V1
  ID3V1���� .Title, 0   '����
  ID3V1���� .Artist, 1  '��Ա
  ID3V1���� .Album, 2   'ר��
  If .Comment(28) = 0 And .Comment(29) > 0 And Len(Text1(2)) > 0 Then Text1(3) = .Comment(29): .Comment(29) = 0
  ID3V1���� .Comment, 5 'ע��
  Text1(4) = StrConv(.Year, vbUnicode)
  If .Genre < 149 Then Combo1.ListIndex = .Genre + 1
End With
End Sub

Private Sub ID3V1����(tem() As Byte, K As Integer)
If IsTextUTF8(tem) Then
  Text1(K) = UTF_8ToTxt(tem)
Else
  Text1(K) = StrConv(tem, vbUnicode)
End If
End Sub

Private Sub ��ȡID3V2��Ϣ()
ID3V2���� "TIT2", 0 '����
ID3V2���� "TPE1", 1 '��Ա
ID3V2���� "TALB", 2 'ר��
ID3V2���� "COMM", 5 'ע��
ID3V2���� "TYER", 4 '���
ID3V2���� "TRCK", 3 '����
ID3V2���� "TXXX", 7 '�û��ı�
ID3V2���� "TCON", 6 '���
If Len(Text2(6)) > 0 Then
  If InStr("( ��", Left(Text2(6), 1)) > 0 Then
    Dim K As Integer
    K = Val(Mid(Text2(6), 2))
    If K < 149 Then Text2(6) = Combo1.List(K + 1)
  Else
    If Val(Text2(6)) Then Text2(6) = Combo1.List(Val(Text2(6)) + 1)
  End If
End If
End Sub

Private Sub ID3V2����(st As String, K As Integer)
Dim Length As Integer, Place As Long, p As Long, i As Long, tem() As Byte, bj As Boolean
Text2(K).ToolTipText = ""
tem = StrConv(st, vbFromUnicode)
p = InStrB(ID3V2Info, tem)
If p > 0 Then
  Length = ID3V2Info(p + 5) * &H80 + ID3V2Info(p + 6) - 1: If Length < 1 Then Exit Sub
  Place = p + 9
  If ID3V2Info(Place) = 1 Then Place = Place + 3: Length = Length - 3: bj = True: If Length < 1 Then Exit Sub 'UTF-16LE����(Unicode����)
  ReDim tem(Length)
  For i = Place To Place + Length: tem(i - Place) = ID3V2Info(i): Next
  If bj Then
    Text2(K) = tem
  Else
    If IsTextUTF8(tem) Then 'UTF-8����
      Text2(K) = Replace(UTF_8ToTxt(tem), Chr(0), "")
    Else
      Text2(K) = Replace(StrConv(tem, vbUnicode), Chr(0), "")
    End If
  End If
  Text2(K).ToolTipText = Text2(K)
End If
End Sub

Private Sub wma��Ϣ����()
On Error GoTo 100
Dim i As Long, K As Long
Caption = Dir(OpenName): Text3(0) = Left(Caption, Len(Caption) - 4)

Open OpenName For Binary As #1
ReDim audioData(LOF(1) - 1)
Get #1, , audioData
Close #1

HeaderLen = audioData(16) + audioData(17) * 256 + audioData(18) * 65536 '���㶥��ͷ�����С
��ȡ��׼��ǩ��Ϣ
��ȡ��չ��ǩ��Ϣ

ObjectSum = audioData(24) '��ȡ��������
K = UBound(audioData)
For i = 0 To 29: wmaHeader(i) = audioData(i): Next '�����ǰ30�ֽ�
For i = 0 To K - 30: audioData(i) = audioData(i + 30): Next '����ǰ��
ReDim Preserve audioData(K - 30)

Exit Sub
100
Close
End Sub

Private Sub ��ȡ��׼��ǩ��Ϣ()
On Error GoTo 100
Dim ObjectID(15) As Byte, i As Integer, k1 As Long, k2 As Long, k3 As Long
Dim Ltag(4) As Integer

For i = 0 To 15: ObjectID(i) = Val("&H" & Mid(tag1ID, i * 2 + 1, 2)): Next
For i = 0 To 4: Text1(i) = "": Text1(i).ToolTipText = "": Text3(i).ToolTipText = "": Next

k2 = InStrB(audioData, ObjectID)
If k2 > 0 Then '����б�׼��ǩ
  k3 = k2 - 1  'k3�Ƕ���ID����ʼλ��
  k2 = k2 + 23: k1 = k2
  For i = 0 To 4
    Ltag(i) = audioData(k1 + i * 2) + audioData(k1 + i * 2 + 1) * 256 '��ȡ�����
    ��׼��ǩ��Ϣ���� Ltag(i) - 2, k2 + 10, i
    k2 = k2 + Ltag(i)
  Next
  k1 = audioData(k3 + 16) + audioData(k3 + 17) * 256 '��ȡ��׼��ǩ�Ĵ�С
  For k2 = k3 To UBound(audioData) - k1: audioData(k2) = audioData(k2 + k1): Next '����ǰ��
  ReDim Preserve audioData(UBound(audioData) - k1)   '��ԭ������ȥ����׼��ǩ
  HeaderLen = HeaderLen - k1 '����ȥ����׼��ǩ��Ķ���ͷ�����С
  audioData(24) = audioData(24) - 1 '����ȥ����׼��ǩ��Ķ�������
End If

100
End Sub

Private Sub ��׼��ǩ��Ϣ����(S1 As Integer, S2 As Long, n As Integer) 's1-��ȣ�s2-��λ�ã�n-�ı�����
Dim j As Integer, i As Long, tem() As Byte
If S1 > 1 Then
  ReDim tem(S1 - 1)
  For i = S2 To S2 + S1 - 1: tem(j) = audioData(i): j = j + 1: Next
  Text1(n) = tem
End If
End Sub

Private Sub ��ȡ��չ��ǩ��Ϣ()
On Error GoTo 100
Dim ObjectID(15) As Byte, i As Integer, k1 As Long, k2 As Long, k3 As Long

For i = 0 To 15: ObjectID(i) = Val("&H" & Mid(tag2ID, i * 2 + 1, 2)): Next

k2 = InStrB(audioData, ObjectID)
If k2 > 0 Then '����б�׼��ǩ
  k3 = k2 - 1  'k3�Ƕ���ID����ʼλ��
  For i = 0 To 7
    Text2(i) = ""
    ��չ��ǩ��Ϣ���� i
  Next
  k1 = audioData(k3 + 16) + audioData(k3 + 17) * 256 '��ȡ��չ��ǩ�Ĵ�С
  For k2 = k3 To UBound(audioData) - k1: audioData(k2) = audioData(k2 + k1): Next '����ǰ��
  ReDim Preserve audioData(UBound(audioData) - k1)   '��ԭ������ȥ����չ��ǩ
  HeaderLen = HeaderLen - k1 '����ȥ����չ��ǩ��Ķ���ͷ�����С
  audioData(24) = audioData(24) - 1 '����ȥ����չ��ǩ��Ķ�������
End If

100
End Sub

Private Sub ��չ��ǩ��Ϣ����(n As Integer) 'n-�ı�����
On Error GoTo 100
Dim tem1() As Byte, tem2() As Byte, j As Integer, K As Long, i As Long, L1 As Long, L2 As Long
tem1 = Wm(n)
K = InStrB(audioData, tem1)
If K > 0 Then '��������֡
  L1 = audioData(K - 3) + audioData(K - 4) * 256 '֡���Ƴ���
  L2 = audioData(K + L1 + 1) + audioData(K + L1 + 2) * 256 '֡���ݳ���
  If L2 > 3 Then
    L1 = L1 + K + 3 '֡������ʼ�ֽ�
    ReDim tem2(L2 - 3)
    For i = L1 To L1 + L2 - 3: tem2(j) = audioData(i): j = j + 1: Next 'ȡ��֡���ݣ�ͬʱȥ���ַ�������2�����ַ�
    Text2(n) = tem2
  End If
End If
100
End Sub

Private Sub ����() '����
On Error GoTo 100
Dim st1 As String, st2 As String
If Len(SaveName) > 6 Then st1 = Left(SaveName, InStrRev(SaveName, "\")) & Dir(OpenName) Else st1 = OpenName
st2 = "����ȫ����Ϣ(*.mp3)" & Chr(0) & "mp3"
SaveName = OpenName
Call saveMP3
Me.Hide
Call FRMMIN.SeeIt(frmma.PLIST.URL(frmma.PLIST.ListIndex))
100
End Sub
Private Function д���׼��ǩ��Ϣ() As Integer
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
д���׼��ǩ��Ϣ = ERR.Number
End Function

Private Function д����չ��ǩ��Ϣ() As Integer
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
д����չ��ǩ��Ϣ = ERR.Number
End Function

Private Sub saveMP3() '����mp3
On Error GoTo 100
Dim i As Integer, K As Long
GoSub 200: GoSub 300: bjTag1 = True: bjTag2 = True 'д��ȫ��
i = 0
If bjTag2 Then i = д��ID3V2
If i = 0 Then
  Open SaveName For Binary As #1
  If bjTag2 Then Seek #1, LOF(1) + 1
  Put #1, , audioData
  Close #1
End If
If bjTag1 Then If i = 0 Then i = д��ID3V1
If i = 0 Then Debug.Print "����ɹ�" Else Debug.Print "����ʧ��,����ţ�" & i
100
Exit Sub
200
i = examine: If i > 0 Then Call SHOWWRONG("ID3V1��Ϣ�е�" & i & "���ı�����������", 0): Exit Sub
Return
300
For i = 0 To 7: K = K + Len(Text2(i)): Next
If K = 0 Then Call SHOWWRONG("Ҫд��ID3V2��Ϣ,���������ı���Ϊ��!", 0): Exit Sub
Return
End Sub

Private Function examine() As Integer
If lstrlen(Text1(0)) > 30 Then examine = 1: Exit Function
If lstrlen(Text1(1)) > 30 Then examine = 2: Exit Function
If lstrlen(Text1(2)) > 30 Then examine = 3: Exit Function
If lstrlen(Text1(5)) > 30 Then examine = 6: Exit Function
If lstrlen(Text1(4)) > 4 Then examine = 5
End Function

Private Function д��ID3V1() As Integer
On Error GoTo 100
Dim ID3V1Info As Mp3tag, i As Integer, tem() As Byte, Tag As String * 3
Tag = "TAG"

With ID3V1Info
  tem = StrConv(Text1(0), vbFromUnicode)
  For i = 0 To UBound(tem): .Title(i) = tem(i): Next   '����

  tem = StrConv(Text1(1), vbFromUnicode)
  For i = 0 To UBound(tem): .Artist(i) = tem(i): Next  '��Ա

  tem = StrConv(Text1(2), vbFromUnicode)
  For i = 0 To UBound(tem): .Album(i) = tem(i): Next   'ר��

  tem = StrConv(Text1(5), vbFromUnicode)
  For i = 0 To UBound(tem): .Comment(i) = tem(i): Next 'ע��

  tem = StrConv(Left(Text1(4) & String(4, 0), 4), vbFromUnicode) '���
  For i = 0 To 3: .Year(i) = tem(i): Next

  i = Val(Text1(3)): If i > 255 Then i = 255
  If Len(Text1(2)) > 0 And i > 0 Then .Comment(28) = 0: .Comment(29) = i '����

  For i = 0 To Combo1.ListCount - 1 '���
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
д��ID3V1 = ERR.Number
End Function

Private Function д��ID3V2() As Integer
On Error GoTo 100
Dim ID3V2 As ID3Header
Dim FrameID() As String '֡��ʶ��
Dim Size(3) As Byte     '֡���ݳ���
Dim flags As Integer    '��־
Dim Data As String      '֡����
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
д��ID3V2 = ERR.Number
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
Call ����
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
�б����
100
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error GoTo 100
If InStr("mp3,wma", LCase(Right(Data.files.Item(1), 3))) Then Effect = vbDropEffectCopy And Effect Else Effect = vbDropEffectNone
100
End Sub

Private Function UTF_8ToTxt(bytSrc() As Byte) As String 'UTF_8����ת��Ϊ��ͨ�ı�
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

Private Function IsTextUTF8(bytSrc() As Byte) As Boolean '�ж��Ƿ�UTF-8����
Dim i As Integer, AscN As Integer, n As Integer
n = UBound(bytSrc)

Do While i <= n
  If bytSrc(i) < 128 Then 'Ascii�ַ�
    i = i + 1: AscN = AscN + 1
  ElseIf (bytSrc(i) And &HF0) = &HE0 Then '3���ֽڵ�UTF-8
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
