VERSION 5.00
Begin VB.Form FRMWEATHER 
   AutoRedraw      =   -1  'True
   BackColor       =   &H005C6105&
   BorderStyle     =   0  'None
   Caption         =   "天气预报"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "天气情况"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "最低气温"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "最高气温"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "城市"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
 Z = "北京|上海|天津|重庆|香港|澳门|哈尔滨|齐齐哈尔|牡丹江| " & vbCrLf & _
 "大庆|伊春|双鸭山|鹤岗|鸡西|佳木斯|七台河|黑河|绥化|大兴安岭| " & vbCrLf & _
 "长春|延吉|吉林|白山|白城|四平|松原|辽源|大安|通化|沈阳| " & vbCrLf & _
 "大连|葫芦岛|盘锦|本溪|抚顺|铁岭|辽阳|营口|阜新|朝阳|锦州|丹东|鞍山| " & vbCrLf & _
    "呼和浩特|呼伦贝尔|锡林浩特|包头|赤峰|海拉尔|乌海|鄂尔多斯|通辽| " & vbCrLf & _
    "石家庄|唐山|张家口|廊坊|邢台|邯郸|沧州|衡水|承德|保定|秦皇岛| " & vbCrLf & _
    "郑州|开封|洛阳|平顶山|焦作|鹤壁|新乡|安阳|濮阳|许昌|漯河| " & vbCrLf & _
    "三门峡|南阳|商丘|信阳|周口|驻马店| " & vbCrLf & _
    "济南|青岛|淄博|威海|曲阜|临沂|烟台|枣庄|聊城| " & vbCrLf & _
    "济宁|菏泽|泰安|日照|东营|德州|滨州|莱芜|潍坊| " & vbCrLf & _
    "太原|阳泉|晋城|晋中|临汾|运城|长治|朔州|忻州|大同|吕梁| " & vbCrLf & _
    "南京|苏州|昆山|南通|太仓|吴县|徐州|宜兴|镇江|淮安|常熟| " & vbCrLf & _
    "盐城|泰州|无锡|连云港|扬州|常州|宿迁| " & vbCrLf & _
    "合肥|巢湖|蚌埠|安庆|六安|滁州|马鞍山|阜阳|宣城| " & vbCrLf & _
    "铜陵|淮北|芜湖|毫州|宿州|淮南|池州| " & vbCrLf & _
    "西安|韩城|安康|汉中|宝鸡|咸阳|榆林|渭南|商洛|铜川|延安| " & vbCrLf & _
    "银川|固原|中卫|石嘴山|吴忠| " & vbCrLf & _
    "兰州|白银|庆阳|酒泉|天水|武威|张掖|甘南|临夏|平凉|定西|金昌| " & vbCrLf & _
    "西宁|海北|海西|黄南|果洛|玉树|海东|海南| " & vbCrLf & _
    "武汉|宜昌|黄冈|恩施|荆州|神农架|十堰|咸宁|襄阳|孝感|随州|黄石|荆门|鄂州| " & vbCrLf & _
    "长沙|邵阳|常德|郴州|吉首|株洲|娄底|湘潭|益阳|永州|岳阳|衡阳|怀化|韶山|张家界| " & vbCrLf & _
    "杭州|湖州|金华|宁波|丽水|绍兴|衢州|嘉兴|台州|舟山|温州| " & vbCrLf & _
    "南昌|萍乡|九江|上饶|抚州|吉安|鹰潭|宜春|新余|景德镇|赣州| " & vbCrLf & _
    "福州|厦门|龙岩|南平|宁德|莆田|泉州|三明|漳州|贵阳|安顺|赤水|遵义|铜仁|六盘水|毕节|凯里|都匀|成都|泸州|内江|凉山|阿坝|巴中|广元|乐山|绵阳|德阳|攀枝花|雅安|宜宾|自贡|甘孜州|达州|资阳|广安|遂宁|眉山|南充|广州|深圳|潮州|韶关|湛江|惠州|清远|东莞|江门|茂名|肇庆|汕尾|河源|揭阳|梅州|中山|德庆|阳江|云浮|珠海|汕头|佛山|南宁|桂林|阳朔|柳州|梧州|玉林|桂平|贺州|钦州|贵港|防城港|百色|北海|河池|来宾|崇左|昆明|保山|楚雄|德宏|红河|临沧|怒江|曲靖|思茅|文山|玉溪|昭通|丽江|大理|海口|三亚|儋州|琼山|通什|文昌|乌鲁木齐|阿勒泰|阿克苏|昌吉|哈密|和田|喀什|克拉玛依|石河子|塔城|库尔勒|吐鲁番|伊宁|拉萨|阿里|昌都|那曲|日喀则|山南|林芝|台北|高雄"
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
Case "北京"
W_P_CODE = "101010100"
Case "上海"
W_P_CODE = "101020100"
Case "天津"
W_P_CODE = "101030100"
Case "重庆"
W_P_CODE = "101040100"
Case "香港"
W_P_CODE = "101320101"
Case "澳门"
W_P_CODE = "101330101"

 Case "哈尔滨"
   W_P_CODE = "101050101"
 Case "齐齐哈尔"

  W_P_CODE = "101050201"
  
  Case "牡丹江"

  W_P_CODE = "101050301"
  
  Case "大庆"

  W_P_CODE = "101050901"
  
  Case "伊春"

  W_P_CODE = "101050801"
  
  Case "双鸭山"

  W_P_CODE = "101051301"
  
  Case "鹤岗"

  W_P_CODE = "101051201"
  
  Case "鸡西"

  W_P_CODE = "101051101"
  
  Case "佳木斯"

  W_P_CODE = "101050401"
  
  Case "七台河"

  W_P_CODE = "101051002"
  
  Case "黑河"

  W_P_CODE = "101050601"
  
  Case "绥化"

  W_P_CODE = "101050501"
  Case "大兴安岭"

  W_P_CODE = "101050701"

 Case "长春"
    W_P_CODE = "101060101"
    Case "延吉"

  W_P_CODE = "101060301"
  
  Case "吉林"

  W_P_CODE = "101060201"
  
  Case "白山"

  W_P_CODE = "101060901"
  
  Case "白城"

  W_P_CODE = "101060601"
  
  Case "四平"

  W_P_CODE = "101060401"
  
  Case "松原"

  W_P_CODE = "101060801"
  
  Case "辽源"

  W_P_CODE = "101060701"
  
  Case "大安"

  W_P_CODE = "101060603"
  
  Case "通化"

  W_P_CODE = "101060501"

 Case "沈阳"
  W_P_CODE = "101070101"
  Case "大连"

  W_P_CODE = "101070201"
  Case "葫芦岛"

  W_P_CODE = "101071401"
  Case "盘锦"

  W_P_CODE = "101071301"
  Case "本溪"

  W_P_CODE = "101070501"
  Case "抚顺"

  W_P_CODE = "101070401"
  Case "铁岭"

  W_P_CODE = "101071101"
  Case "辽阳"

  W_P_CODE = "101071001"
  Case "营口"

  W_P_CODE = "101070801"
  Case "阜新"

  W_P_CODE = "101070901"
  Case "朝阳"

  W_P_CODE = "101071201"
  Case "锦州"

  W_P_CODE = "101070701"
  Case "丹东"
 
  W_P_CODE = "101070601"
  Case "鞍山"

  W_P_CODE = "101070301"

 Case "呼和浩特"
 W_P_CODE = "101080101"
 Case "呼伦贝尔"

 W_P_CODE = "101081000"
 
 Case "锡林浩特"

 W_P_CODE = "101080901"
 
 Case "包头"

 W_P_CODE = "101080201"
 
 Case "赤峰"

 W_P_CODE = "101080601"
 
 Case "海拉尔"

 W_P_CODE = "101081001"
 
 Case "乌海"

 W_P_CODE = "101080301"
 
 Case "鄂尔多斯"

 W_P_CODE = "101080701"
 
 Case "通辽"
 
 W_P_CODE = "101080501"

 Case "石家庄"
  W_P_CODE = "101090101"
  Case "唐山"
  W_P_CODE = "101090101"
  Case "张家口"

  W_P_CODE = "101090301"
  Case "廊坊"

  W_P_CODE = "101090601"
  Case "邢台"
 
  W_P_CODE = "101090901"
  Case "邯郸"

  W_P_CODE = "101091001"
  Case "沧州"

  W_P_CODE = "101090701"
  Case "衡水"
 
  W_P_CODE = "101090801"
  Case "承德"

  W_P_CODE = "101090402"
  Case "保定"

  W_P_CODE = "101090201"
  
  Case "秦皇岛"

  W_P_CODE = "101091101"


 Case "郑州"
      W_P_CODE = "101180101"
Case "开封"

 W_P_CODE = "101180801"
 Case "洛阳"

 W_P_CODE = "101180901"
 Case "平顶山"

 W_P_CODE = "101180501"
 Case "焦作"

 W_P_CODE = "101181101"
 Case "鹤壁"

 W_P_CODE = "101181201"
 Case "新乡"

 W_P_CODE = "101180301"
 Case "安阳"

 W_P_CODE = "101180201"
 Case "濮阳"

 W_P_CODE = "101181301"
 Case "许昌"

 W_P_CODE = "101180401"
 Case "漯河"

 W_P_CODE = "101181501"
 Case "三门峡"

 W_P_CODE = "101181701"
 Case "南阳"

 W_P_CODE = "101180701"
 Case "商丘"

 W_P_CODE = "101181001"
 Case "信阳"

 W_P_CODE = "101180601"
 Case "周口"
 
 W_P_CODE = "101181401"
 Case "驻马店"

 W_P_CODE = "101181601"

 Case "济南"
  W_P_CODE = "101120101"
 Case "青岛"
  W_P_CODE = "101120201"
  Case "淄博"

  W_P_CODE = "101120301"
  Case "威海"

  W_P_CODE = "101121301"
  Case "曲阜"

  W_P_CODE = "101120710"
  Case "临沂"
 
  W_P_CODE = "101120901"
  Case "烟台"

  W_P_CODE = "101120501"
  Case "枣庄"

  W_P_CODE = "101121401"
  Case "聊城"

  W_P_CODE = "101121701"
  Case "济宁"

  W_P_CODE = "101120701"
  Case "菏泽"

  W_P_CODE = "101121001"
  Case "泰安"

  W_P_CODE = "101120801"
  Case "日照"

  W_P_CODE = "101121501"
  Case "东营"

  W_P_CODE = "101121201"
  Case "德州"

  W_P_CODE = "101120401"
  Case "滨州"

  W_P_CODE = "101121101"
  Case "莱芜"
  
  W_P_CODE = "101121601"
  Case "潍坊"

  W_P_CODE = "101120601"

Case "太原"
 W_P_CODE = "101100101"

Case "阳泉"
  W_P_CODE = "101100301"
  Case "晋城"

  W_P_CODE = "101100601"
  Case "晋中"

  W_P_CODE = "101100401"
  Case "临汾"
 
  W_P_CODE = "101100701"
  Case "运城"

  W_P_CODE = "101100801"
  Case "长治"

  W_P_CODE = "101100501"
  Case "朔州"

  W_P_CODE = "101100901"
  Case "忻州"

  W_P_CODE = "101101001"
  Case "大同"

  W_P_CODE = "101100201"
  Case "吕梁"

  W_P_CODE = "101101101"


Case "南京"
W_P_CODE = "101190101"
Case "苏州"

  W_P_CODE = "101190401"
  Case "昆山"

  W_P_CODE = "101190404"
  Case "南通"

  W_P_CODE = "101190501"
  Case "太仓"

  W_P_CODE = "101190408"
  Case "吴县"

  W_P_CODE = "101190406"
  Case "徐州"

  W_P_CODE = "101190801"
  Case "宜兴"

  W_P_CODE = "101190203"
  Case "镇江"
 
  W_P_CODE = "101190301"
  Case "淮安"

  W_P_CODE = "101190901"
  Case "常熟"

  W_P_CODE = "101190402"
  Case "盐城"

  W_P_CODE = "101190701"
  Case "泰州"

  W_P_CODE = "101191201"
  Case "无锡"

  W_P_CODE = "101190201"
  Case "连云港"

  W_P_CODE = "101191001"
  Case "扬州"

  W_P_CODE = "101190601"
  Case "常州"
 
  W_P_CODE = "101191101"
  Case "宿迁"

  W_P_CODE = "101191301"

 Case "合肥"
      W_P_CODE = "101220101"
Case "巢湖"

  W_P_CODE = "101221601"
  Case "蚌埠"

  W_P_CODE = "101220201"
  Case "安庆"

  W_P_CODE = "101220601"
  Case "六安"

  W_P_CODE = "101221501"
  Case "滁州"

  W_P_CODE = "101221101"
  Case "马鞍山"

  W_P_CODE = "101220501"
  Case "阜阳"
  
  W_P_CODE = "101220801"
  Case "宣城"

  W_P_CODE = "101221401"
  Case "铜陵"

  W_P_CODE = "101221301"
  Case "淮北"

  W_P_CODE = "101221201"
  Case "芜湖"

  W_P_CODE = "101220301"
  Case "毫州"

  W_P_CODE = "101220901"
  Case "宿州"

  W_P_CODE = "101220701"
  Case "淮南"

  W_P_CODE = "101220401"
  Case "池州"

  W_P_CODE = "101221701"

 Case "西安"
 W_P_CODE = "101110101"
 Case "韩城"
  
  W_P_CODE = "101110510"
  Case "安康"

  W_P_CODE = "101110701"
  Case "汉中"

  W_P_CODE = "101110801"
  Case "宝鸡"

  W_P_CODE = "101110901"
  Case "咸阳"

  W_P_CODE = "101110200"
  
  Case "榆林"

  W_P_CODE = "101110401"
  Case "渭南"

  W_P_CODE = "101110501"
  Case "商洛"

  W_P_CODE = "101110601"
  Case "铜川"

  W_P_CODE = "101111001"
  Case "延安"

  W_P_CODE = "101110300"


 Case "银川"
 W_P_CODE = "101170101"
 Case "固原"

 W_P_CODE = "101170401"
 Case "中卫"

 W_P_CODE = "101170501"
 Case "石嘴山"

 W_P_CODE = "101170201"
 Case "吴忠"
 
 W_P_CODE = "101170301"

 Case "兰州"
  W_P_CODE = "101160101"
  Case "白银"

  W_P_CODE = "101161301"
  Case "庆阳"

  W_P_CODE = "101160401"
  Case "酒泉"

  W_P_CODE = "101160801"
  Case "天水"

  W_P_CODE = "101160901"
  Case "武威"

  W_P_CODE = "101160501"
  Case "张掖"

  W_P_CODE = "101160701"
  Case "甘南"

  W_P_CODE = "101050204"
  Case "临夏"

  W_P_CODE = "101161101"
  Case "平凉"

  W_P_CODE = "101160301"
  Case "定西"

  W_P_CODE = "101160201"
  Case "金昌"

  W_P_CODE = "101160601"

  Case "西宁"
     W_P_CODE = "101150101"
Case "海北"
 
 W_P_CODE = "101150801"
 Case "海西"

 W_P_CODE = "101150701"
 Case "黄南"

 W_P_CODE = "101150301"
 Case "果洛"

 W_P_CODE = "101150501"
 Case "玉树"

 W_P_CODE = "101150601"
 Case "海东"

 W_P_CODE = "101150201"
 Case "海南"

 W_P_CODE = "101150401"


Case "武汉"
W_P_CODE = "101200101"

Case "宜昌"

 W_P_CODE = "101200901"
 
 Case "黄冈"

 W_P_CODE = "101200501"
 
 Case "恩施"

 W_P_CODE = "101201001"
 
 Case "荆州"
 W_P_CODE = "101200801"
 
 Case "神农架"

 W_P_CODE = "101201201"
 
 Case "十堰"

 W_P_CODE = "101201101"
 
 Case "咸宁"

 W_P_CODE = "101200701"
 
 Case "襄阳"

 W_P_CODE = "101200201"
 
 Case "孝感"

 W_P_CODE = "101200401"
 
 Case "随州"

 W_P_CODE = "101201301"
 Case "黄石"

 W_P_CODE = "101200601"
 Case "荆门"

 W_P_CODE = "101201401"
 Case "鄂州"
 
 W_P_CODE = "101200301"

 Case "长沙"
  W_P_CODE = "101250101"
  
  Case "邵阳"

  W_P_CODE = "101250901"
  
  Case "常德"

  W_P_CODE = "101250601"
  
  Case "郴州"

  W_P_CODE = "101250501"
  
  Case "吉首"

  W_P_CODE = "101251501"
  
  Case "株洲"

  W_P_CODE = "101250301"
  
  Case "娄底"
 
  W_P_CODE = "101250801"
  
  Case "湘潭"

  W_P_CODE = "101250201"
  
  Case "益阳"

  W_P_CODE = "101250701"
  
  Case "永州"

  W_P_CODE = "101251401"
  
  Case "岳阳"

  W_P_CODE = "101251001"
  
  Case "衡阳"

  W_P_CODE = "101250401"
  
  Case "怀化"

  W_P_CODE = "101251201"
  
  Case "韶山"

  W_P_CODE = "101250202"
  
  Case "张家界"

  W_P_CODE = "101251101"

Case "杭州"
W_P_CODE = "101210101"
Case "湖州"

  W_P_CODE = "101210201"
  
  Case "金华"
  
  W_P_CODE = "101210901"
  
  Case "宁波"

  W_P_CODE = "101210401"
  
  Case "丽水"

  W_P_CODE = "101210801"
  
  Case "绍兴"

  W_P_CODE = "101210501"
  
  Case "衢州"

  W_P_CODE = "101211001"
  
  Case "嘉兴"

  W_P_CODE = "101210301"
  
  Case "台州"

  W_P_CODE = "101210601"
  
  Case "舟山"

  W_P_CODE = "101211101"
  Case "温州"

  W_P_CODE = "101210701"


Case "南昌"
W_P_CODE = "101240101"
Case "萍乡"

  W_P_CODE = "101240901"
  
  Case "九江"

  W_P_CODE = "101240201"
  
  Case "上饶"

  W_P_CODE = "101240301"
  
  Case "抚州"
 
  W_P_CODE = "101240401"
  
  Case "吉安"

  W_P_CODE = "101240601"
  
  Case "鹰潭"

  W_P_CODE = "101241101"
  
  Case "宜春"
 
  W_P_CODE = "101240501"
  
  Case "新余"

  W_P_CODE = "101241001"
  Case "景德镇"

  W_P_CODE = "101240801"
  Case "赣州"

  W_P_CODE = "101240701"

 Case "福州"
 W_P_CODE = "101230101"
 Case "厦门"

  W_P_CODE = "101230201"
  
  Case "龙岩"

  W_P_CODE = "101230701"
  
  Case "南平"

  W_P_CODE = "101230901"
  
  Case "宁德"

  W_P_CODE = "101230301"
  
  Case "莆田"

  W_P_CODE = "101230401"
  
  Case "泉州"

  W_P_CODE = "101230501"
  
  Case "三明"

  W_P_CODE = "101230801"
  
  Case "漳州"
 
  W_P_CODE = "101230601"


 Case "贵阳"
   W_P_CODE = "101260101"
   Case "安顺"

  W_P_CODE = "101260301"
  Case "赤水"

  W_P_CODE = "101260208"
  Case "遵义"

  W_P_CODE = "101260201"
  Case "铜仁"

  W_P_CODE = "101260601"
  Case "六盘水"

  W_P_CODE = "101260801"
  Case "毕节"

  W_P_CODE = "101260701"
  Case "凯里"

  W_P_CODE = "101260501"
  Case "都匀"

  W_P_CODE = "101260401"


Case "成都"
W_P_CODE = "101270101"
Case "泸州"

  W_P_CODE = "101271001"
  
  Case "内江"

  W_P_CODE = "101271201"
  
  Case "凉山"

  W_P_CODE = "101271601"
  
  Case "阿坝"

  W_P_CODE = "101271901"
  
  Case "巴中"

  W_P_CODE = "101270901"
  
  Case "广元"

  W_P_CODE = "101272101"
  Case "乐山"
 
  W_P_CODE = "101271401"
  
  Case "绵阳"

  W_P_CODE = "101270401"
  Case "德阳"

  W_P_CODE = "101272001"
  Case "攀枝花"

  W_P_CODE = "101270201"
  
  Case "雅安"

  W_P_CODE = "101271701"
  
  Case "宜宾"

  W_P_CODE = "101271101"
  
  Case "自贡"

  W_P_CODE = "101270301"
  
  Case "甘孜州"
 
  W_P_CODE = "101271801"
  
  Case "达州"

  W_P_CODE = "101270601"
  
  Case "资阳"

  W_P_CODE = "101271301"
  
  Case "广安"

  W_P_CODE = "101270801"
  
  Case "遂宁"

  W_P_CODE = "101270701"
  
  Case "眉山"

  W_P_CODE = "101271501"
  
  Case "南充"

  W_P_CODE = "101270501"


  Case "广州"
   W_P_CODE = "101280101"
 
 Case "深圳"
 W_P_CODE = "101280601"

 
 Case "潮州"
 W_P_CODE = "101281501"

 
 Case "韶关"
 W_P_CODE = "101280201"

 
 Case "湛江"

  W_P_CODE = "101281001"

 Case "惠州"

  W_P_CODE = "101280301"

 Case "清远"
 W_P_CODE = "101281301"

 
 Case "东莞"
 W_P_CODE = "101281601"

 
 Case "江门"
 W_P_CODE = "101281101"

 
 Case "茂名"
  W_P_CODE = "101282001"

 
 Case "肇庆"

  W_P_CODE = "101280901"

 Case "汕尾"
 W_P_CODE = "101282101"

 
 Case "河源"
 W_P_CODE = "101281201"

 
 Case "揭阳"
 W_P_CODE = "101281901"

 
 Case "梅州"

  W_P_CODE = "101280401"

 Case "中山"
 W_P_CODE = "101281701"

 Case "德庆"

 W_P_CODE = "101280905"
 
 Case "阳江"
 W_P_CODE = "101281801"

  Case "云浮"
 W_P_CODE = "101281401"
  Case "珠海"
 W_P_CODE = "101280701"
  Case "汕头"
 W_P_CODE = "101280501"
 Case "佛山"

 W_P_CODE = "101280800"

Case "南宁"
 W_P_CODE = "101300101"
 Case "桂林"

  W_P_CODE = "101300501"
  
  Case "阳朔"

  W_P_CODE = "101300510"
  
  Case "柳州"

  W_P_CODE = "101300301"
  
  Case "梧州"

  W_P_CODE = "101300601"
  
  Case "玉林"

  W_P_CODE = "101300901"
  
  Case "桂平"

  W_P_CODE = "101300802"
  
  Case "贺州"

  W_P_CODE = "101300701"
  
  Case "钦州"

  W_P_CODE = "101301101"
  
  Case "贵港"
 
  W_P_CODE = "101300801"
  
  Case "防城港"

  W_P_CODE = "101301401"
  
  Case "百色"

  W_P_CODE = "101301001"
  
  Case "北海"

  W_P_CODE = "101301301"
  
  Case "河池"

  W_P_CODE = "101301201"
  Case "来宾"

  W_P_CODE = "101300401"
  Case "崇左"

  W_P_CODE = "101300201"
Case "昆明"
 W_P_CODE = "101290101"
 Case "保山"

  W_P_CODE = "101290501"
  
  Case "楚雄"

  W_P_CODE = "101290801"
  
  Case "德宏"

  W_P_CODE = "101291501"
  
  Case "红河"

  W_P_CODE = "101290301"
  
  Case "临沧"

  W_P_CODE = "101291101"
  
  Case "怒江"

  W_P_CODE = "101291201"
  
  Case "曲靖"

  W_P_CODE = "101290401"
  
  Case "思茅"

  W_P_CODE = "101290901"
  
  Case "文山"

  W_P_CODE = "101290601"
  Case "玉溪"
 
  W_P_CODE = "101290701"
  
  Case "昭通"

  W_P_CODE = "101291001"
  Case "丽江"

  W_P_CODE = "101291401"
  Case "大理"

  W_P_CODE = "101290201"
Case "海口"
 W_P_CODE = "101310101"
 Case "三亚"

 W_P_CODE = "101310201"
 Case "儋州"

 W_P_CODE = "101310205"
 Case "琼山"

 W_P_CODE = "101310102"
 Case "通什"

 W_P_CODE = "101310222"
 Case "文昌"

 W_P_CODE = "101310212"
Case "乌鲁木齐"
W_P_CODE = "101130101"
Case "阿勒泰"

  W_P_CODE = "101131401"
  Case "阿克苏"

  W_P_CODE = "101130801"
  Case "昌吉"

  W_P_CODE = "101130401"
  Case "哈密"

  W_P_CODE = "101131201"
  Case "和田"

  W_P_CODE = "101131301"
  Case "喀什"

  W_P_CODE = "101130901"
  Case "克拉玛依"

  W_P_CODE = "101130201"
  Case "石河子"

  W_P_CODE = "101130301"
  Case "塔城"

  W_P_CODE = "101131101"
  Case "库尔勒"

  W_P_CODE = "101130601"
  Case "吐鲁番"

  W_P_CODE = "101130501"
  Case "伊宁"
  
  W_P_CODE = "101131001"
Case "拉萨"
 W_P_CODE = "101140101"
 Case "阿里"
  W_P_CODE = "101140701"
Case "昌都"
 W_P_CODE = "101140501"
 Case "那曲"
 W_P_CODE = "101140601"
Case "日喀则"
 W_P_CODE = "101140201"
Case "山南"
 W_P_CODE = "101140301"
Case "林芝"
 W_P_CODE = "101140401"
 Case "台北"
 W_P_CODE = "101340102"
 Case "高雄"
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
