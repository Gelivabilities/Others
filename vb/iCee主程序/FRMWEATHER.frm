VERSION 5.00
Begin VB.Form FRMWEATHER 
   AutoRedraw      =   -1  'True
   BackColor       =   &H005C6105&
   BorderStyle     =   0  'None
   Caption         =   "����Ԥ��"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin iCee.I_COMBO CBTOWN 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   420
   End
End
Attribute VB_Name = "FRMWEATHER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Sub LOADCITY()
On Error Resume Next
Dim ST() As String, Z As String
READYLOAD = True
 Z = "����|�Ϻ�|���|����|���|����|������|�������|ĵ����| " & vbCrLf & _
 "����|����|˫Ѽɽ|�׸�|����|��ľ˹|��̨��|�ں�|�绯|���˰���| " & vbCrLf & _
 "����|�Ӽ�|����|��ɽ|�׳�|��ƽ|��ԭ|��Դ|��|ͨ��|����| " & vbCrLf & _
 "����|��«��|�̽�|��Ϫ|��˳|����|����|Ӫ��|����|����|����|����|��ɽ| " & vbCrLf & _
    "���ͺ���|���ױ���|���ֺ���|��ͷ|���|������|�ں�|������˹|ͨ��| " & vbCrLf & _
    "ʯ��ׯ|��ɽ|�żҿ�|�ȷ�|��̨|����|����|��ˮ|�е�|����|�ػʵ�| " & vbCrLf & _
    "֣��|����|����|ƽ��ɽ|����|�ױ�|����|����|���|���|���| " & vbCrLf & _
    "����Ͽ|����|����|����|�ܿ�|פ���| " & vbCrLf & _
    "����|�ൺ|�Ͳ�|����|����|����|��̨|��ׯ|�ĳ�| " & vbCrLf & _
    "����|����|̩��|����|��Ӫ|����|����|����|Ϋ��| " & vbCrLf & _
    "̫ԭ|��Ȫ|����|����|�ٷ�|�˳�|����|˷��|����|��ͬ|����| " & vbCrLf & _
    "�Ͼ�|����|��ɽ|��ͨ|̫��|����|����|����|��|����|����| " & vbCrLf & _
    "�γ�|̩��|����|���Ƹ�|����|����|��Ǩ| " & vbCrLf & _
    "�Ϸ�|����|����|����|����|����|��ɽ|����|����| " & vbCrLf & _
    "ͭ��|����|�ߺ�|����|����|����|����| " & vbCrLf & _
    "����|����|����|����|����|����|����|μ��|����|ͭ��|�Ӱ�| " & vbCrLf & _
    "����|��ԭ|����|ʯ��ɽ|����| " & vbCrLf & _
    "����|����|����|��Ȫ|��ˮ|����|��Ҵ|����|����|ƽ��|����|���| " & vbCrLf & _
    "����|����|����|����|����|����|����|����| " & vbCrLf & _
    "�人|�˲�|�Ƹ�|��ʩ|����|��ũ��|ʮ��|����|����|Т��|����|��ʯ|����|����| " & vbCrLf & _
    "��ɳ|����|����|����|����|����|¦��|��̶|����|����|����|����|����|��ɽ|�żҽ�| " & vbCrLf & _
    "����|����|��|����|��ˮ|����|����|����|̨��|��ɽ|����| " & vbCrLf & _
    "�ϲ�|Ƽ��|�Ž�|����|����|����|ӥ̶|�˴�|����|������|����| " & vbCrLf & _
    "����|����|����|��ƽ|����|����|Ȫ��|����|����|����|��˳|��ˮ|����|ͭ��|����ˮ|�Ͻ�|����|����|�ɶ�|����|�ڽ�|��ɽ|����|����|��Ԫ|��ɽ|����|����|��֦��|�Ű�|�˱�|�Թ�|������|����|����|�㰲|����|üɽ|�ϳ�|����|����|����|�ع�|տ��|����|��Զ|��ݸ|����|ï��|����|��β|��Դ|����|÷��|��ɽ|����|����|�Ƹ�|�麣|��ͷ|��ɽ|����|����|��˷|����|����|����|��ƽ|����|����|���|���Ǹ�|��ɫ|����|�ӳ�|����|����|����|��ɽ|����|�º�|���|�ٲ�|ŭ��|����|˼é|��ɽ|��Ϫ|��ͨ|����|����|����|����|����|��ɽ|ͨʲ|�Ĳ�|��³ľ��|����̩|������|����|����|����|��ʲ|��������|ʯ����|����|�����|��³��|����|����|����|����|����|�տ���|ɽ��|��֥|̨��|����"
ST = Split(Z, "|")
For i = 0 To UBound(ST)
CBTOWN.AddItem Trim(ST(i))
Next
CBTOWN.ListIndex = GetInitEntry("SYSTEM", "WEATHER", 0)

End Sub

Sub GETWEATHER()
W_P_URL = "http://www.weather.com.cn/data/cityinfo/" & W_P_CODE & ".html"
Text1.Text = Replace(ReadinteFile(W_P_URL), """", "")
Text1.Text = Replace(Text1.Text, ",", " ")
Text1.Text = Replace(Text1.Text, ":", " ")
Text1.Text = Replace(Text1.Text, "{", "")
Text1.Text = Replace(Text1.Text, "}", "")
Text1.Text = Replace(Text1.Text, "weatherinfo", "")
Text1.Text = Replace(Text1.Text, "city", "|")
Text1.Text = Replace(Text1.Text, W_P_CODE, "")
Text1.Text = Replace(Text1.Text, "weather ", "|")
Text1.Text = Replace(Text1.Text, "temp2", "|")
Text1.Text = Replace(Text1.Text, "temp1", "|")
Text1.Text = Replace(Text1.Text, "id", "")
Text1.Text = Replace(Text1.Text, "img", "|")
Text1.Text = Replace(Text1.Text, " ", "")
LA(3).Caption = Split(Text1.Text, "|")(5)  '
LA(2).Caption = Split(Text1.Text, "|")(4) '
LA(1).Caption = Split(Text1.Text, "|")(3) '
LA(0).Caption = Split(Text1.Text, "|")(1) '
End Sub


Private Sub CBTOWN_Change()
Select Case CBTOWN.Text
Case "����"
W_P_CODE = "101010100"
Case "�Ϻ�"
W_P_CODE = "101020100"
Case "���"
W_P_CODE = "101030100"
Case "����"
W_P_CODE = "101040100"
Case "���"
W_P_CODE = "101320101"
Case "����"
W_P_CODE = "101330101"

 Case "������"
   W_P_CODE = "101050101"
 Case "�������"

  W_P_CODE = "101050201"
  
  Case "ĵ����"

  W_P_CODE = "101050301"
  
  Case "����"

  W_P_CODE = "101050901"
  
  Case "����"

  W_P_CODE = "101050801"
  
  Case "˫Ѽɽ"

  W_P_CODE = "101051301"
  
  Case "�׸�"

  W_P_CODE = "101051201"
  
  Case "����"

  W_P_CODE = "101051101"
  
  Case "��ľ˹"

  W_P_CODE = "101050401"
  
  Case "��̨��"

  W_P_CODE = "101051002"
  
  Case "�ں�"

  W_P_CODE = "101050601"
  
  Case "�绯"

  W_P_CODE = "101050501"
  Case "���˰���"

  W_P_CODE = "101050701"

 Case "����"
    W_P_CODE = "101060101"
    Case "�Ӽ�"

  W_P_CODE = "101060301"
  
  Case "����"

  W_P_CODE = "101060201"
  
  Case "��ɽ"

  W_P_CODE = "101060901"
  
  Case "�׳�"

  W_P_CODE = "101060601"
  
  Case "��ƽ"

  W_P_CODE = "101060401"
  
  Case "��ԭ"

  W_P_CODE = "101060801"
  
  Case "��Դ"

  W_P_CODE = "101060701"
  
  Case "��"

  W_P_CODE = "101060603"
  
  Case "ͨ��"

  W_P_CODE = "101060501"

 Case "����"
  W_P_CODE = "101070101"
  Case "����"

  W_P_CODE = "101070201"
  Case "��«��"

  W_P_CODE = "101071401"
  Case "�̽�"

  W_P_CODE = "101071301"
  Case "��Ϫ"

  W_P_CODE = "101070501"
  Case "��˳"

  W_P_CODE = "101070401"
  Case "����"

  W_P_CODE = "101071101"
  Case "����"

  W_P_CODE = "101071001"
  Case "Ӫ��"

  W_P_CODE = "101070801"
  Case "����"

  W_P_CODE = "101070901"
  Case "����"

  W_P_CODE = "101071201"
  Case "����"

  W_P_CODE = "101070701"
  Case "����"
 
  W_P_CODE = "101070601"
  Case "��ɽ"

  W_P_CODE = "101070301"

 Case "���ͺ���"
 W_P_CODE = "101080101"
 Case "���ױ���"

 W_P_CODE = "101081000"
 
 Case "���ֺ���"

 W_P_CODE = "101080901"
 
 Case "��ͷ"

 W_P_CODE = "101080201"
 
 Case "���"

 W_P_CODE = "101080601"
 
 Case "������"

 W_P_CODE = "101081001"
 
 Case "�ں�"

 W_P_CODE = "101080301"
 
 Case "������˹"

 W_P_CODE = "101080701"
 
 Case "ͨ��"
 
 W_P_CODE = "101080501"

 Case "ʯ��ׯ"
  W_P_CODE = "101090101"
  Case "��ɽ"
  W_P_CODE = "101090101"
  Case "�żҿ�"

  W_P_CODE = "101090301"
  Case "�ȷ�"

  W_P_CODE = "101090601"
  Case "��̨"
 
  W_P_CODE = "101090901"
  Case "����"

  W_P_CODE = "101091001"
  Case "����"

  W_P_CODE = "101090701"
  Case "��ˮ"
 
  W_P_CODE = "101090801"
  Case "�е�"

  W_P_CODE = "101090402"
  Case "����"

  W_P_CODE = "101090201"
  
  Case "�ػʵ�"

  W_P_CODE = "101091101"


 Case "֣��"
      W_P_CODE = "101180101"
Case "����"

 W_P_CODE = "101180801"
 Case "����"

 W_P_CODE = "101180901"
 Case "ƽ��ɽ"

 W_P_CODE = "101180501"
 Case "����"

 W_P_CODE = "101181101"
 Case "�ױ�"

 W_P_CODE = "101181201"
 Case "����"

 W_P_CODE = "101180301"
 Case "����"

 W_P_CODE = "101180201"
 Case "���"

 W_P_CODE = "101181301"
 Case "���"

 W_P_CODE = "101180401"
 Case "���"

 W_P_CODE = "101181501"
 Case "����Ͽ"

 W_P_CODE = "101181701"
 Case "����"

 W_P_CODE = "101180701"
 Case "����"

 W_P_CODE = "101181001"
 Case "����"

 W_P_CODE = "101180601"
 Case "�ܿ�"
 
 W_P_CODE = "101181401"
 Case "פ���"

 W_P_CODE = "101181601"

 Case "����"
  W_P_CODE = "101120101"
 Case "�ൺ"
  W_P_CODE = "101120201"
  Case "�Ͳ�"

  W_P_CODE = "101120301"
  Case "����"

  W_P_CODE = "101121301"
  Case "����"

  W_P_CODE = "101120710"
  Case "����"
 
  W_P_CODE = "101120901"
  Case "��̨"

  W_P_CODE = "101120501"
  Case "��ׯ"

  W_P_CODE = "101121401"
  Case "�ĳ�"

  W_P_CODE = "101121701"
  Case "����"

  W_P_CODE = "101120701"
  Case "����"

  W_P_CODE = "101121001"
  Case "̩��"

  W_P_CODE = "101120801"
  Case "����"

  W_P_CODE = "101121501"
  Case "��Ӫ"

  W_P_CODE = "101121201"
  Case "����"

  W_P_CODE = "101120401"
  Case "����"

  W_P_CODE = "101121101"
  Case "����"
  
  W_P_CODE = "101121601"
  Case "Ϋ��"

  W_P_CODE = "101120601"

Case "̫ԭ"
 W_P_CODE = "101100101"

Case "��Ȫ"
  W_P_CODE = "101100301"
  Case "����"

  W_P_CODE = "101100601"
  Case "����"

  W_P_CODE = "101100401"
  Case "�ٷ�"
 
  W_P_CODE = "101100701"
  Case "�˳�"

  W_P_CODE = "101100801"
  Case "����"

  W_P_CODE = "101100501"
  Case "˷��"

  W_P_CODE = "101100901"
  Case "����"

  W_P_CODE = "101101001"
  Case "��ͬ"

  W_P_CODE = "101100201"
  Case "����"

  W_P_CODE = "101101101"


Case "�Ͼ�"
W_P_CODE = "101190101"
Case "����"

  W_P_CODE = "101190401"
  Case "��ɽ"

  W_P_CODE = "101190404"
  Case "��ͨ"

  W_P_CODE = "101190501"
  Case "̫��"

  W_P_CODE = "101190408"
  Case "����"

  W_P_CODE = "101190406"
  Case "����"

  W_P_CODE = "101190801"
  Case "����"

  W_P_CODE = "101190203"
  Case "��"
 
  W_P_CODE = "101190301"
  Case "����"

  W_P_CODE = "101190901"
  Case "����"

  W_P_CODE = "101190402"
  Case "�γ�"

  W_P_CODE = "101190701"
  Case "̩��"

  W_P_CODE = "101191201"
  Case "����"

  W_P_CODE = "101190201"
  Case "���Ƹ�"

  W_P_CODE = "101191001"
  Case "����"

  W_P_CODE = "101190601"
  Case "����"
 
  W_P_CODE = "101191101"
  Case "��Ǩ"

  W_P_CODE = "101191301"

 Case "�Ϸ�"
      W_P_CODE = "101220101"
Case "����"

  W_P_CODE = "101221601"
  Case "����"

  W_P_CODE = "101220201"
  Case "����"

  W_P_CODE = "101220601"
  Case "����"

  W_P_CODE = "101221501"
  Case "����"

  W_P_CODE = "101221101"
  Case "��ɽ"

  W_P_CODE = "101220501"
  Case "����"
  
  W_P_CODE = "101220801"
  Case "����"

  W_P_CODE = "101221401"
  Case "ͭ��"

  W_P_CODE = "101221301"
  Case "����"

  W_P_CODE = "101221201"
  Case "�ߺ�"

  W_P_CODE = "101220301"
  Case "����"

  W_P_CODE = "101220901"
  Case "����"

  W_P_CODE = "101220701"
  Case "����"

  W_P_CODE = "101220401"
  Case "����"

  W_P_CODE = "101221701"

 Case "����"
 W_P_CODE = "101110101"
 Case "����"
  
  W_P_CODE = "101110510"
  Case "����"

  W_P_CODE = "101110701"
  Case "����"

  W_P_CODE = "101110801"
  Case "����"

  W_P_CODE = "101110901"
  Case "����"

  W_P_CODE = "101110200"
  
  Case "����"

  W_P_CODE = "101110401"
  Case "μ��"

  W_P_CODE = "101110501"
  Case "����"

  W_P_CODE = "101110601"
  Case "ͭ��"

  W_P_CODE = "101111001"
  Case "�Ӱ�"

  W_P_CODE = "101110300"


 Case "����"
 W_P_CODE = "101170101"
 Case "��ԭ"

 W_P_CODE = "101170401"
 Case "����"

 W_P_CODE = "101170501"
 Case "ʯ��ɽ"

 W_P_CODE = "101170201"
 Case "����"
 
 W_P_CODE = "101170301"

 Case "����"
  W_P_CODE = "101160101"
  Case "����"

  W_P_CODE = "101161301"
  Case "����"

  W_P_CODE = "101160401"
  Case "��Ȫ"

  W_P_CODE = "101160801"
  Case "��ˮ"

  W_P_CODE = "101160901"
  Case "����"

  W_P_CODE = "101160501"
  Case "��Ҵ"

  W_P_CODE = "101160701"
  Case "����"

  W_P_CODE = "101050204"
  Case "����"

  W_P_CODE = "101161101"
  Case "ƽ��"

  W_P_CODE = "101160301"
  Case "����"

  W_P_CODE = "101160201"
  Case "���"

  W_P_CODE = "101160601"

  Case "����"
     W_P_CODE = "101150101"
Case "����"
 
 W_P_CODE = "101150801"
 Case "����"

 W_P_CODE = "101150701"
 Case "����"

 W_P_CODE = "101150301"
 Case "����"

 W_P_CODE = "101150501"
 Case "����"

 W_P_CODE = "101150601"
 Case "����"

 W_P_CODE = "101150201"
 Case "����"

 W_P_CODE = "101150401"


Case "�人"
W_P_CODE = "101200101"

Case "�˲�"

 W_P_CODE = "101200901"
 
 Case "�Ƹ�"

 W_P_CODE = "101200501"
 
 Case "��ʩ"

 W_P_CODE = "101201001"
 
 Case "����"
 W_P_CODE = "101200801"
 
 Case "��ũ��"

 W_P_CODE = "101201201"
 
 Case "ʮ��"

 W_P_CODE = "101201101"
 
 Case "����"

 W_P_CODE = "101200701"
 
 Case "����"

 W_P_CODE = "101200201"
 
 Case "Т��"

 W_P_CODE = "101200401"
 
 Case "����"

 W_P_CODE = "101201301"
 Case "��ʯ"

 W_P_CODE = "101200601"
 Case "����"

 W_P_CODE = "101201401"
 Case "����"
 
 W_P_CODE = "101200301"

 Case "��ɳ"
  W_P_CODE = "101250101"
  
  Case "����"

  W_P_CODE = "101250901"
  
  Case "����"

  W_P_CODE = "101250601"
  
  Case "����"

  W_P_CODE = "101250501"
  
  Case "����"

  W_P_CODE = "101251501"
  
  Case "����"

  W_P_CODE = "101250301"
  
  Case "¦��"
 
  W_P_CODE = "101250801"
  
  Case "��̶"

  W_P_CODE = "101250201"
  
  Case "����"

  W_P_CODE = "101250701"
  
  Case "����"

  W_P_CODE = "101251401"
  
  Case "����"

  W_P_CODE = "101251001"
  
  Case "����"

  W_P_CODE = "101250401"
  
  Case "����"

  W_P_CODE = "101251201"
  
  Case "��ɽ"

  W_P_CODE = "101250202"
  
  Case "�żҽ�"

  W_P_CODE = "101251101"

Case "����"
W_P_CODE = "101210101"
Case "����"

  W_P_CODE = "101210201"
  
  Case "��"
  
  W_P_CODE = "101210901"
  
  Case "����"

  W_P_CODE = "101210401"
  
  Case "��ˮ"

  W_P_CODE = "101210801"
  
  Case "����"

  W_P_CODE = "101210501"
  
  Case "����"

  W_P_CODE = "101211001"
  
  Case "����"

  W_P_CODE = "101210301"
  
  Case "̨��"

  W_P_CODE = "101210601"
  
  Case "��ɽ"

  W_P_CODE = "101211101"
  Case "����"

  W_P_CODE = "101210701"


Case "�ϲ�"
W_P_CODE = "101240101"
Case "Ƽ��"

  W_P_CODE = "101240901"
  
  Case "�Ž�"

  W_P_CODE = "101240201"
  
  Case "����"

  W_P_CODE = "101240301"
  
  Case "����"
 
  W_P_CODE = "101240401"
  
  Case "����"

  W_P_CODE = "101240601"
  
  Case "ӥ̶"

  W_P_CODE = "101241101"
  
  Case "�˴�"
 
  W_P_CODE = "101240501"
  
  Case "����"

  W_P_CODE = "101241001"
  Case "������"

  W_P_CODE = "101240801"
  Case "����"

  W_P_CODE = "101240701"

 Case "����"
 W_P_CODE = "101230101"
 Case "����"

  W_P_CODE = "101230201"
  
  Case "����"

  W_P_CODE = "101230701"
  
  Case "��ƽ"

  W_P_CODE = "101230901"
  
  Case "����"

  W_P_CODE = "101230301"
  
  Case "����"

  W_P_CODE = "101230401"
  
  Case "Ȫ��"

  W_P_CODE = "101230501"
  
  Case "����"

  W_P_CODE = "101230801"
  
  Case "����"
 
  W_P_CODE = "101230601"


 Case "����"
   W_P_CODE = "101260101"
   Case "��˳"

  W_P_CODE = "101260301"
  Case "��ˮ"

  W_P_CODE = "101260208"
  Case "����"

  W_P_CODE = "101260201"
  Case "ͭ��"

  W_P_CODE = "101260601"
  Case "����ˮ"

  W_P_CODE = "101260801"
  Case "�Ͻ�"

  W_P_CODE = "101260701"
  Case "����"

  W_P_CODE = "101260501"
  Case "����"

  W_P_CODE = "101260401"


Case "�ɶ�"
W_P_CODE = "101270101"
Case "����"

  W_P_CODE = "101271001"
  
  Case "�ڽ�"

  W_P_CODE = "101271201"
  
  Case "��ɽ"

  W_P_CODE = "101271601"
  
  Case "����"

  W_P_CODE = "101271901"
  
  Case "����"

  W_P_CODE = "101270901"
  
  Case "��Ԫ"

  W_P_CODE = "101272101"
  Case "��ɽ"
 
  W_P_CODE = "101271401"
  
  Case "����"

  W_P_CODE = "101270401"
  Case "����"

  W_P_CODE = "101272001"
  Case "��֦��"

  W_P_CODE = "101270201"
  
  Case "�Ű�"

  W_P_CODE = "101271701"
  
  Case "�˱�"

  W_P_CODE = "101271101"
  
  Case "�Թ�"

  W_P_CODE = "101270301"
  
  Case "������"
 
  W_P_CODE = "101271801"
  
  Case "����"

  W_P_CODE = "101270601"
  
  Case "����"

  W_P_CODE = "101271301"
  
  Case "�㰲"

  W_P_CODE = "101270801"
  
  Case "����"

  W_P_CODE = "101270701"
  
  Case "üɽ"

  W_P_CODE = "101271501"
  
  Case "�ϳ�"

  W_P_CODE = "101270501"


  Case "����"
   W_P_CODE = "101280101"
 
 Case "����"
 W_P_CODE = "101280601"

 
 Case "����"
 W_P_CODE = "101281501"

 
 Case "�ع�"
 W_P_CODE = "101280201"

 
 Case "տ��"

  W_P_CODE = "101281001"

 Case "����"

  W_P_CODE = "101280301"

 Case "��Զ"
 W_P_CODE = "101281301"

 
 Case "��ݸ"
 W_P_CODE = "101281601"

 
 Case "����"
 W_P_CODE = "101281101"

 
 Case "ï��"
  W_P_CODE = "101282001"

 
 Case "����"

  W_P_CODE = "101280901"

 Case "��β"
 W_P_CODE = "101282101"

 
 Case "��Դ"
 W_P_CODE = "101281201"

 
 Case "����"
 W_P_CODE = "101281901"

 
 Case "÷��"

  W_P_CODE = "101280401"

 Case "��ɽ"
 W_P_CODE = "101281701"

 Case "����"

 W_P_CODE = "101280905"
 
 Case "����"
 W_P_CODE = "101281801"

  Case "�Ƹ�"
 W_P_CODE = "101281401"
  Case "�麣"
 W_P_CODE = "101280701"
  Case "��ͷ"
 W_P_CODE = "101280501"
 Case "��ɽ"

 W_P_CODE = "101280800"

Case "����"
 W_P_CODE = "101300101"
 Case "����"

  W_P_CODE = "101300501"
  
  Case "��˷"

  W_P_CODE = "101300510"
  
  Case "����"

  W_P_CODE = "101300301"
  
  Case "����"

  W_P_CODE = "101300601"
  
  Case "����"

  W_P_CODE = "101300901"
  
  Case "��ƽ"

  W_P_CODE = "101300802"
  
  Case "����"

  W_P_CODE = "101300701"
  
  Case "����"

  W_P_CODE = "101301101"
  
  Case "���"
 
  W_P_CODE = "101300801"
  
  Case "���Ǹ�"

  W_P_CODE = "101301401"
  
  Case "��ɫ"

  W_P_CODE = "101301001"
  
  Case "����"

  W_P_CODE = "101301301"
  
  Case "�ӳ�"

  W_P_CODE = "101301201"
  Case "����"

  W_P_CODE = "101300401"
  Case "����"

  W_P_CODE = "101300201"
Case "����"
 W_P_CODE = "101290101"
 Case "��ɽ"

  W_P_CODE = "101290501"
  
  Case "����"

  W_P_CODE = "101290801"
  
  Case "�º�"

  W_P_CODE = "101291501"
  
  Case "���"

  W_P_CODE = "101290301"
  
  Case "�ٲ�"

  W_P_CODE = "101291101"
  
  Case "ŭ��"

  W_P_CODE = "101291201"
  
  Case "����"

  W_P_CODE = "101290401"
  
  Case "˼é"

  W_P_CODE = "101290901"
  
  Case "��ɽ"

  W_P_CODE = "101290601"
  Case "��Ϫ"
 
  W_P_CODE = "101290701"
  
  Case "��ͨ"

  W_P_CODE = "101291001"
  Case "����"

  W_P_CODE = "101291401"
  Case "����"

  W_P_CODE = "101290201"
Case "����"
 W_P_CODE = "101310101"
 Case "����"

 W_P_CODE = "101310201"
 Case "����"

 W_P_CODE = "101310205"
 Case "��ɽ"

 W_P_CODE = "101310102"
 Case "ͨʲ"

 W_P_CODE = "101310222"
 Case "�Ĳ�"

 W_P_CODE = "101310212"
Case "��³ľ��"
W_P_CODE = "101130101"
Case "����̩"

  W_P_CODE = "101131401"
  Case "������"

  W_P_CODE = "101130801"
  Case "����"

  W_P_CODE = "101130401"
  Case "����"

  W_P_CODE = "101131201"
  Case "����"

  W_P_CODE = "101131301"
  Case "��ʲ"

  W_P_CODE = "101130901"
  Case "��������"

  W_P_CODE = "101130201"
  Case "ʯ����"

  W_P_CODE = "101130301"
  Case "����"

  W_P_CODE = "101131101"
  Case "�����"

  W_P_CODE = "101130601"
  Case "��³��"

  W_P_CODE = "101130501"
  Case "����"
  
  W_P_CODE = "101131001"
Case "����"
 W_P_CODE = "101140101"
 Case "����"
  W_P_CODE = "101140701"
Case "����"
 W_P_CODE = "101140501"
 Case "����"
 W_P_CODE = "101140601"
Case "�տ���"
 W_P_CODE = "101140201"
Case "ɽ��"
 W_P_CODE = "101140301"
Case "��֥"
 W_P_CODE = "101140401"
 Case "̨��"
 W_P_CODE = "101340102"
 Case "����"
W_P_CODE = "101340201"
End Select
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then Exit Sub
If READYLOAD = False Then Call GETWEATHER
lRet = SetInitEntry("SYSTEM", "WEATHER", CBTOWN.ListIndex)
End Sub
Private Sub Form_Load()
Call LOADCITY
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H808080, B
MakeTransparent Me.hWnd, 250
End Sub
