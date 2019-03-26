VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type SolarHolidayStruct
    Month As Long
    Day As Long
    Recess As Long
    HolidayName As String
End Type

Private Type LunarHolidayStruct
    Month As Long
    Day As Long
    Recess As Long
    HolidayName As String
End Type

Private Type WeekHolidayStruct
    Month As Long
    WeekAtMonth As Long
    WeekDay As Long
    HolidayName As String
End Type

'保持属性值的局部变量
Private mvarsYear As Long                                                       '局部复制
Private mvarsMonth As Long                                                      '局部复制
Private mvarsDay As Long                                                        '局部复制
Private mvarlYear As Long                                                       '局部复制
Private mvarlMonth As Long                                                      '局部复制
Private mvarlDay As Long                                                        '局部复制
Private mvarIsLeap As Boolean                                                   '局部复制

'Private Declare Function BitRight32 Lib "Bit4VB.DLL" (ByVal x As Long, ByVal num As Long) As Long
'此定义已不使用
'定义类内部用公用变量

Private SolarMonth As Variant
Private Gan As Variant
Private Zhi As Variant
Private Animals As Variant
Private SolarTerm As Variant
Private sTermInfo As Variant
Private nStr1 As Variant
Private nStr2 As Variant
Private MonthName As Variant
Private LunarInfo(150) As Long
Private LunarYearDays(150) As Long
Private sHolidayInfo() As SolarHolidayStruct
Private lHolidayInfo() As LunarHolidayStruct
Private wHolidayInfo() As WeekHolidayStruct

Private mvarDate As Date                                                        '内部使用标准的日期变量

Private BitPower(31) As Long                                                    '0-31

'位测试 ,测试位为1 返回真
Private Function mvarBitTest32(Number As Long, Bit As Long) As Boolean
    If Bit < 0 Or Bit > 31 Then '不是整数位
        mvarBitTest32 = False
    Else
        If Number And BitPower(Bit) Then
          mvarBitTest32 = True
        Else
          mvarBitTest32 = False
        End If
    End If
End Function

Private Sub Class_Initialize()
   Dim tempArray As Variant
   Dim i As Long
   Dim b As Long
   Dim sFtv As Variant
   Dim lFtv As Variant
   Dim wFtv As Variant
   
    
   '根据VB的位计算特点,故扩充原有的数据位,将其变成32位
   tempArray = Array( _
        &H104BD8, &H104AE0, &H10A570, &H1054D5, &H10D260, &H10D950, &H116554, &H1056A0, &H109AD0, &H1055D2, _
        &H104AE0, &H10A5B6, &H10A4D0, &H10D250, &H11D255, &H10B540, &H10D6A0, &H10ADA2, &H1095B0, &H114977, _
        &H104970, &H10A4B0, &H10B4B5, &H106A50, &H106D40, &H11AB54, &H102B60, &H109570, &H1052F2, &H104970, _
        &H106566, &H10D4A0, &H10EA50, &H106E95, &H105AD0, &H102B60, &H1186E3, &H1092E0, &H11C8D7, &H10C950, _
        &H10D4A0, &H11D8A6, &H10B550, &H1056A0, &H11A5B4, &H1025D0, &H1092D0, &H10D2B2, &H10A950, &H10B557, _
        &H106CA0, &H10B550, &H115355, &H104DA0, &H10A5D0, &H114573, &H1052D0, &H10A9A8, &H10E950, &H106AA0, _
        &H10AEA6, &H10AB50, &H104B60, &H10AAE4, &H10A570, &H105260, &H10F263, &H10D950, &H105B57, &H1056A0, _
        &H1096D0, &H104DD5, &H104AD0, &H10A4D0, &H10D4D4, &H10D250, &H10D558, &H10B540, &H10B5A0, &H1195A6, _
        &H1095B0, &H1049B0, &H10A974, &H10A4B0, &H10B27A, &H106A50, &H106D40, &H10AF46, &H10AB60, &H109570, _
        &H104AF5, &H104970, &H1064B0, &H1074A3, &H10EA50, &H106B58, &H1055C0, &H10AB60, &H1096D5, &H1092E0, _
        &H10C960, &H10D954, &H10D4A0, &H10DA50, &H107552, &H1056A0, &H10ABB7, &H1025D0, &H1092D0, &H10CAB5, _
        &H10A950, &H10B4A0, &H10BAA4, &H10AD50, &H1055D9, &H104BA0, &H10A5B0, &H115176, &H1052B0, &H10A930, _
        &H107954, &H106AA0, &H10AD50, &H105B52, &H104B60, &H10A6E6, &H10A4E0, &H10D260, &H10EA65, &H10D530, _
        &H105AA0, &H1076A3, &H1096D0, &H104BD7, &H104AD0, &H10A4D0, &H11D0B6, &H10D250, &H10D520, &H10DD45, _
        &H10B5A0, &H1056D0, &H1055B2, &H1049B0, &H10A577, &H10A4B0, &H10AA50, &H11B255, &H106D20, &H10ADA0)
   For i = 0 To 149
     LunarInfo(i) = tempArray(i)
   Next
   
   tempArray = Array( _
                384, 354, 355, 383, 354, 355, 384, 354, 355, 384, _
                354, 384, 354, 354, 384, 354, 355, 384, 355, 384, _
                354, 354, 384, 354, 354, 385, 354, 355, 384, 354, _
                383, 354, 355, 384, 355, 354, 384, 354, 384, 354, _
                354, 384, 355, 354, 385, 354, 354, 384, 354, 384, _
                354, 355, 384, 354, 355, 384, 354, 383, 355, 354, _
                384, 355, 354, 384, 355, 353, 384, 355, 384, 354, _
                355, 384, 354, 354, 384, 354, 384, 354, 355, 384, _
                355, 354, 384, 354, 384, 354, 354, 384, 355, 355, _
                384, 354, 354, 383, 355, 384, 354, 355, 384, 354, _
                354, 384, 354, 355, 384, 354, 385, 354, 354, 384, _
                354, 354, 384, 355, 384, 354, 355, 384, 354, 354, _
                384, 354, 355, 384, 354, 384, 354, 354, 384, 355, _
                354, 384, 355, 384, 354, 354, 384, 354, 354, 384, _
                355, 355, 384, 354, 384, 354, 354, 384, 354, 355)
   
   For i = 0 To 149
     LunarYearDays(i) = tempArray(i)
   Next
   
   SolarMonth = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
   Gan = Array("甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸")
   Zhi = Array("子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥")
   Animals = Array("鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊", "猴", "鸡", "狗", "猪")
   SolarTerm = Array("小寒", "大寒", "立春", "雨水", "惊蛰", "春分", "清明", "谷雨", "立夏", "小满", "芒种", "夏至", "小暑", "大暑", "立秋", "处暑", "白露", "秋分", "寒露", "霜降", "立冬", "小雪", "大雪", "冬至")
   sTermInfo = Array(0, 21208, 42467, 63836, 85337, 107014, 128867, 150921, 173149, 195551, 218072, 240693, 263343, 285989, 308563, 331033, 353350, 375494, 397447, 419210, 440795, 462224, 483532, 504758)
   nStr1 = Array("日", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
   nStr2 = Array("初", "十", "廿", "卅", "　")
   MonthName = Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
   
   '国历节日 *表示放假日
   sFtv = Array( _
        1, 1, 1, "元旦", _
        2, 2, 0, "世界湿地日", 2, 14, 0, "情人节", 2, 10, 0, "国际气象节", _
        3, 1, 0, "国际海豹日", 3, 5, 0, "学雷锋纪念日", 3, 8, 0, "妇女节", 3, 12, 0, "植树节 孙中山逝世纪念日", 3, 14, 0, "国际警察日", 3, 15, 0, "消费者权益日", 3, 17, 0, "中国国医节 国际航海日", 3, 21, 0, "世界森林日 消除种族歧视国际日 世界儿歌日", 3, 22, 0, "世界水日", 3, 24, 0, "世界防治结核病日", _
        4, 1, 0, "愚人节", 4, 28, 0, "08'胶济铁路火车相撞", 4, 7, 0, "世界卫生日", 4, 22, 0, "世界地球日", _
        5, 1, 1, "劳动节", 5, 4, 0, "青年节", 5, 12, 0, "08'汶川地震 护士节", 5, 31, 0, "世界无烟日", 5, 8, 0, "世界红十字日", 5, 12, 0, "国际护士节", _
        6, 1, 0, "儿童节", 6, 5, 0, "世界环境保护日", 6, 26, 0, "国际禁毒日", _
        7, 1, 0, "建党节 香港回归纪念 世界建筑日", 7, 11, 0, "世界人口日", _
        8, 1, 0, "建军节", 8, 8, 0, "中国男子节 父亲节", 8, 15, 0, "抗日战争胜利纪念", _
        9, 9, 0, "毛泽东逝世纪念", 9, 10, 0, "教师节", 9, 18, 0, "九·一八事变纪念日", 9, 28, 0, "孔子诞辰", 9, 20, 0, "国际爱牙日", 9, 27, 0, "世界旅游日", _
        10, 1, 0, "国庆节 国际音乐日", 10, 2, 1, "国庆节假日", 10, 3, 1, "国庆节假日", 10, 6, 0, "老人节", 10, 24, 0, "联合国日", _
        11, 12, 0, "孙中山诞辰纪念", 11, 3, 0, "", 11, 1, 0, "国际旅游电影节", 11, 8, 0, "中国记者日", 11, 10, 0, "世界青年日", 11, 14, 0, "世界糖尿病日", 11, 17, 0, "世界学生节", _
        12, 1, 0, "世界艾滋病日", 12, 3, 0, "世界残疾人日", 12, 20, 0, "澳门回归纪念", 12, 24, 0, "平安夜", 12, 25, 0, "圣诞节", 12, 26, 0, "毛泽东诞辰纪念", 12, 9, 0, "世界足球日", 12, 10, 0, "世界人权日", 12, 13, 0, "南京大屠杀纪念日", _
        9, 3, 0, "抗日战争胜利日", 9, 17, 0, "国际和平日", 9, 20, 0, "全国爱牙日", 10, 8, 0, "全国高血压日", 10, 15, 0, "国际盲人节(白手杖日)", 10, 16, 0, "世界粮食日", 10, 31, 0, "万圣节")
        
   b = UBound(sFtv) + 1
   ReDim sHolidayInfo(b / 4)
   For i = 0 To (b / 4) - 1
       sHolidayInfo(i).Month = sFtv(i * 4)
       sHolidayInfo(i).Day = sFtv(i * 4 + 1)
       sHolidayInfo(i).Recess = sFtv(i * 4 + 2)
       sHolidayInfo(i).HolidayName = sFtv(i * 4 + 3)
   Next
   
   '农历节日 *表示放假日
   lFtv = Array( _
        1, 1, 1, "春节", _
        1, 15, 0, "元宵节", _
        5, 5, 1, "端午节", _
        7, 7, 0, "七夕情人节", _
        7, 15, 0, "中元节 盂兰盆节", _
        8, 15, 1, "中秋节", _
        9, 9, 0, "重阳节", _
        12, 8, 0, "腊八节", _
        12, 24, 0, "小年")
        '12, 31, 0, "除夕")  '注意除夕需要其它方法进行计算
   
   b = UBound(lFtv) + 1
   ReDim lHolidayInfo(b / 4)
   For i = 0 To (b / 4) - 1
       lHolidayInfo(i).Month = lFtv(i * 4)
       lHolidayInfo(i).Day = lFtv(i * 4 + 1)
       lHolidayInfo(i).Recess = lFtv(i * 4 + 2)
       lHolidayInfo(i).HolidayName = lFtv(i * 4 + 3)
   Next
   
   '某月的第几个星期几
   wFtv = Array( _
       5, 2, 1, "国际母亲节", _
       5, 3, 1, "全国助残日", _
       6, 3, 1, "父亲节", _
       9, 3, 3, "国际和平日", _
       9, 4, 1, "国际聋人节", _
       10, 1, 2, "国际住房日", _
       10, 1, 4, "国际减轻自然灾害日", _
       11, 4, 5, "感恩节")
   b = UBound(wFtv) + 1
   ReDim wHolidayInfo(b / 4)
   For i = 0 To (b / 4) - 1
       wHolidayInfo(i).Month = wFtv(i * 4)
       wHolidayInfo(i).WeekAtMonth = wFtv(i * 4 + 1)
       wHolidayInfo(i).WeekDay = wFtv(i * 4 + 2) '1 代表星期天
       wHolidayInfo(i).HolidayName = wFtv(i * 4 + 3)
   Next


   '位操作初使化模块函数 modBit4VB中定义
   For i = 0 To 30
      BitPower(i) = 2 ^ i
   Next
   BitPower(31) = &H80000000
End Sub

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////

'计算农历上的节气
Public Property Get lSolarTerm() As String

    '//===== 某年的第n个节气为几日(从0小寒起算)
    'function sTerm(y,n) {
    '   var offDate = new Date( ( 31556925974.7*(y-1900) + sTermInfo[n]*60000  ) + Date.UTC(1900,0,6,2,5) )
    '   return(offDate.getUTCDate())

    '//节气
    '   tmp1 = sTerm(y, m * 2) - 1
    
    Dim baseDateAndTime As Date
    Dim newdate As Date
    Dim num As Double
    Dim Y As Long
    Dim TempStr As String
    
    baseDateAndTime = #1/6/1900 2:05:00 AM#
    Y = mvarsYear
    TempStr = ""
    
    Dim i As Long
    For i = 1 To 24
       num = 525948.76 * (Y - 1900) + sTermInfo(i - 1)
       newdate = DateAdd("n", num, baseDateAndTime)  '按分钟计算,之所以不按秒钟计算，是因为会溢出
       If Abs(DateDiff("d", newdate, mvarDate)) = 0 Then
          TempStr = SolarTerm(i - 1)
          Exit For
       End If
    Next
  
    lSolarTerm = TempStr
End Property
'计算按第几周星期几计算的节日
Public Property Get wHoliday() As String
    Dim w As Long
    Dim i As Long
    Dim b As Long
    Dim FirstDay As Date
    Dim TempStr As String
    TempStr = ""
    b = UBound(wHolidayInfo)
    For i = 0 To b
      If wHolidayInfo(i).Month = mvarsMonth Then  '当月份相当时
         w = WeekDay(mvarDate)
         If wHolidayInfo(i).WeekDay = w Then  '仅当星期几也相等时
            FirstDay = mvarsMonth & "/" & 1 & "/" & mvarsYear '取当月第一天
            If (DateDiff("ww", FirstDay, mvarDate) = wHolidayInfo(i).WeekAtMonth) Then
                TempStr = wHolidayInfo(i).HolidayName
            End If
         End If
      End If
    Next
    
    
    wHoliday = TempStr
End Property

Public Property Get lHoliday() As String
    Dim i As Long
    Dim b As Long
    Dim TempStr As String
    Dim oy As Long
    Dim odate As Date
    Dim ndate As Date
    TempStr = ""
    b = UBound(lHolidayInfo)
    If mvarlMonth = 12 And (mvarlDay = 29 Or mvarlDay = 30) Then
       '保
       oy = mvarlYear '保存农历年数
       odate = mvarDate
       ndate = mvarDate + 1
       Call sInitDate(Year(ndate), Month(ndate), Day(ndate)) '计算第二天的属性
       If oy = mvarlYear - 1 Then '如果农历年数增加了1
          TempStr = "除夕"
          Call sInitDate(Year(odate), Month(odate), Day(odate)) '恢复到今天原有数据
          
       End If
    Else
        For i = 0 To b
           If (lHolidayInfo(i).Month = mvarlMonth) And _
              (lHolidayInfo(i).Day = mvarlDay) Then
               TempStr = lHolidayInfo(i).HolidayName
               Exit For
           End If
        Next
    End If
    lHoliday = TempStr
End Property
'求公历节日
Public Property Get sHoliday() As String
    Dim i As Long
    Dim b As Long
    Dim TempStr As String
    
    TempStr = ""
    b = UBound(sHolidayInfo)
    For i = 0 To b
       If (sHolidayInfo(i).Month = mvarsMonth) And _
          (sHolidayInfo(i).Day = mvarsDay) Then
           TempStr = sHolidayInfo(i).HolidayName
           Exit For
       End If
    Next
    sHoliday = TempStr
End Property
Public Property Get sHolidayRecess() As Boolean
    Dim i As Long
    Dim b As Long
    Dim TempStr As Boolean
    
    TempStr = False
    b = UBound(sHolidayInfo)
    For i = 0 To b
       If (sHolidayInfo(i).Month = mvarsMonth) And _
          (sHolidayInfo(i).Day = mvarsDay) Then
           TempStr = sHolidayInfo(i).Recess = 1
           Exit For
       End If
    Next
    sHolidayRecess = TempStr
End Property
'是否是农历的闰月

Public Property Get IsLeap() As Boolean
    IsLeap = mvarIsLeap
End Property

Public Property Get lDay() As Long
    lDay = mvarlDay
End Property

Public Property Get lMonth() As Long
    lMonth = mvarlMonth
End Property

Public Property Get lYear() As Long
    lYear = mvarlYear
End Property
Public Property Get sWeekDay() As Long
    sWeekDay = WeekDay(mvarDate)
End Property

'计算星期几中文字串
Public Property Get sWeekDayStr() As String
   Select Case WeekDay(mvarDate)
     Case vbSunday
        sWeekDayStr = "星期日"
     Case vbMonday
        sWeekDayStr = "星期一"
     Case vbTuesday
        sWeekDayStr = "星期二"
     Case vbWednesday
        sWeekDayStr = "星期三"
     Case vbThursday
        sWeekDayStr = "星期四"
     Case vbFriday
        sWeekDayStr = "星期五"
     Case vbSaturday
        sWeekDayStr = "星期六"
   End Select
End Property

Public Function Constellation2(m As Long, d As Long) As String
   Dim Y As Long
   Dim tempDate As Date
   Dim ConstellName As String
 
   Y = 2000
   tempDate = m & "/" & d & "/" & Y

   Select Case tempDate
      Case #3/21/2000# To #4/19/2000#
            ConstellName = "阳性.火象星座，守护行星:火星"
      Case #4/20/2000# To #5/20/2000#
            ConstellName = "阴性.地象星座，守护行星:金星"
      Case #5/21/2000# To #6/21/2000#
            ConstellName = "阳性.风象星座，守护行星:水星"
      Case #6/22/2000# To #7/22/2000#
            ConstellName = "阴性.水象星座守护行星:月亮"
      Case #7/23/2000# To #8/22/2000#
            ConstellName = "阳性.火象星座，守护行星:太阳"
      Case #8/23/2000# To #9/22/2000#
            ConstellName = "阴性.土象星座，守护行星:水星"
      Case #9/23/2000# To #10/23/2000#
            ConstellName = "阳性.风象星座，守护行星:金星"
      Case #10/24/2000# To #11/21/2000#
            ConstellName = "阴性.水象星座，守护行星:冥王星(传统上为火星)"
      Case #11/22/2000# To #12/21/2000#
            ConstellName = "阳性.火象星座，守护行星:木星"
      Case #12/22/2000# To #12/31/2000#
            ConstellName = "阴性.土象星座，守护行星:土星"
      Case #1/1/2000# To #1/19/2000#
            ConstellName = "阴性.土象星座，守护行星:土星"
      Case #1/20/2000# To #2/18/2000#
            ConstellName = "阳性.风象星座,守护行星:天王星(传统上为土星)"
      Case #2/19/2000# To #3/20/2000#
            ConstellName = "阴性.水象星座，守护行星:海王星"
      Case Else
            ConstellName = ""
   End Select
   Constellation2 = ConstellName
End Function

Public Property Get sDay() As Long
    sDay = mvarsDay
End Property

Public Property Get sMonth() As Long
    sMonth = mvarsMonth
End Property

Public Property Get sYear() As Long
    sYear = mvarsYear
End Property

'////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function IsToday(Y As Long, m As Long, d As Long) As Boolean
   
    If (Year(Date) = Y) And _
       (Month(Date) = m) And _
       (Day(Date) = d) Then
        IsToday = True
    Else
        IsToday = False
    End If
    
End Function


'根据年份不同计算当年属于什么朝代
Public Function Era(Y As Long) As String
   Dim TempStr As String
   
   If Y < 1874 Then
       TempStr = "未知"
   Else
       If Y <= 1908 Then
          TempStr = "清朝光绪"
          If Y = 1874 Then
            TempStr = TempStr & "元年"
          Else
            TempStr = TempStr & UpNumber(CStr(Y - 1874)) & "年"
          End If
       Else
          If Y <= 1910 Then
             TempStr = "清朝宣统"
             If Y = 1909 Then
                TempStr = TempStr & "元年"
             Else
                TempStr = TempStr & UpNumber(CStr(Y - 1909 + 1)) & "年"
             End If
          Else
             If Y < 1949 Then
                TempStr = "中华民国"
                If Y = 1912 Then
                   TempStr = TempStr & "元年"
                Else
                   TempStr = TempStr & UpNumber(CStr(Y - 1912 + 1)) & "年"
                End If
             Else
                TempStr = "中华人民共和国成立"
                If Y = 1949 Then
                   TempStr = TempStr & "了"
                Else
                   Select Case Y
                      Case 2000
                         TempStr = "千禧年"
                      Case Else
                         TempStr = TempStr & UpNumber(CStr(Y - 1949)) & "周年"
                   End Select
                End If
             End If
          End If
       End If
   End If
   
   Era = TempStr
End Function

' 传入 num 传回干支, 0=甲子
Public Function GanZhi(num As Long) As String
    Dim TempStr As String
    Dim i As Long
    i = (num - 1864) Mod 60 '计算干支
    TempStr = Gan(i Mod 10) & Zhi(i Mod 12)
    GanZhi = TempStr
End Function

'计算年的属相字串
Public Function YearAttribute(Y As Long) As String
    YearAttribute = Animals((Y - 1900) Mod 12)
End Function

'将数字汉化
Public Function UpNumber(Dxs As String) As String

 '检测为空时
 If Trim(Dxs) = "" Then
    UpNumber = ""
    Exit Function
 End If
  
 Dim Sw As Integer, SzUp As Integer, TempStr As String, DXStr As String
    Sw = Len(Trim(Dxs))
     
     Dim i As Integer
     For i = 1 To Sw
         TempStr = Right(Trim(Dxs), i)
         TempStr = Left(TempStr, 1)
         TempStr = Converts(TempStr)
         Select Case i
           Case 1
               If TempStr = "零" Then
                  TempStr = ""
                   Else
                  TempStr = TempStr + ""
               End If
           Case 2
               If TempStr = "零" Then
                  TempStr = "零"
                Else
                  TempStr = TempStr + "十"
               End If
           Case 3
               If TempStr = "零" Then
                  TempStr = "零"
                   Else
                  TempStr = TempStr + "百"
               End If
            Case 4
               If TempStr = "零" Then
                  TempStr = "零"
                   Else
                  TempStr = TempStr + "千"
               End If
            Case 5
               If TempStr = "零" Then
                  TempStr = "万"
                   Else
                  TempStr = TempStr + "万"
               End If
            Case 6
               If TempStr = "零" Then
                  TempStr = "零"
                   Else
                  TempStr = TempStr + "十"
               End If
            Case 7
               If TempStr = "零" Then
                  TempStr = "零"
                   Else
                  TempStr = TempStr + "百"
               End If
            Case 8
               If TempStr = "零" Then
                  TempStr = "零"
                   Else
                  TempStr = TempStr + "千"
               End If
            Case 9
               If TempStr = "零" Then
                  TempStr = "亿"
                   Else
                  TempStr = TempStr + "亿"
               End If
         End Select
    Dim TempA As String
        TempA = Left(Trim(DXStr), 1)
    If TempStr = "零" Then
      Select Case TempA
       Case "零"
            DXStr = DXStr
       Case "万"
            DXStr = DXStr
       Case "亿"
            DXStr = DXStr
       Case Else
            DXStr = TempStr + DXStr
      End Select
      Else
       DXStr = TempStr + DXStr
    End If
     Next
  
    UpNumber = DXStr
End Function
Private Function Converts(NumStr As String) As String
    Select Case Val(NumStr)
      Case 0
        Converts = "零"
      Case 1
        Converts = "一"
      Case 2
        Converts = "二"
      Case 3
        Converts = "三"
      Case 4
        Converts = "四"
      Case 5
        Converts = "五"
      Case 6
         Converts = "六"
      Case 7
         Converts = "七"
      Case 8
         Converts = "八"
      Case 9
         Converts = "九"
    End Select
End Function
'中文日期
Public Function CDayStr(d As Long) As String
   Dim s As String
   Select Case d
      Case 0
        s = ""
      Case 10
        s = "初十"
      Case 20
        s = "二十"
      Case 30
        s = "三十"
      Case Else
        s = nStr2(d \ 10)  '整数除法
        s = s & nStr1(d Mod 10)
   End Select
   CDayStr = s
End Function
'计算星座归属
Public Function Constellation(m As Long, d As Long) As String
   Dim Y As Long
   Dim tempDate As Date
   Dim ConstellName As String
 
   Y = 2000
   tempDate = m & "/" & d & "/" & Y

   Select Case tempDate
      Case #3/21/2000# To #4/19/2000#
            ConstellName = "山羊"
      Case #4/20/2000# To #5/20/2000#
            ConstellName = "金牛"
      Case #5/21/2000# To #6/21/2000#
            ConstellName = "双子"
      Case #6/22/2000# To #7/22/2000#
            ConstellName = "巨蟹"
      Case #7/23/2000# To #8/22/2000#
            ConstellName = "狮子"
      Case #8/23/2000# To #9/22/2000#
            ConstellName = "处女"
      Case #9/23/2000# To #10/23/2000#
            ConstellName = "天平"
      Case #10/24/2000# To #11/21/2000#
            ConstellName = "天蝎"
      Case #11/22/2000# To #12/21/2000#
            ConstellName = "射手"
      Case #12/22/2000# To #12/31/2000#
            ConstellName = "摩蝎"
      Case #1/1/2000# To #1/19/2000#
            ConstellName = "摩蝎"
      Case #1/20/2000# To #2/18/2000#
            ConstellName = "水瓶"
      Case #2/19/2000# To #3/20/2000#
            ConstellName = "双鱼"
      Case Else
            ConstellName = ""
   End Select
   Constellation = ConstellName
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
'以下为类内部使用的一些函数
'传回农历 y年的总天数
Private Function lYearDays(ByVal Y As Long) As Long
   
'   Dim i As Long
'   Dim f As Long
'   Dim sumDay As Long
'   Dim info As Long
   
'   sumDay = 348
'   i = &H8000
'   info = LunarInfo(y - 1900) And &H1000FFFF  '屏蔽高位,
'   Do
'     f = info And i
'     If f <> 0 Then
'       sumDay = sumDay + 1
'     End If
'     i = BitRight16(i, 1)
'   Loop Until i < &H10
'   lYearDays = sumDay + leapDays(y)
    
    lYearDays = LunarYearDays(Y - 1900)  '先计算出每年的天数,并形成数组,以减少以后的运算时间
End Function

'传回农历 y年m月的总天数
Private Function lMonthDays(ByVal Y As Long, ByVal m As Long) As Long
  'If (LunarInfo(y - 1900) And &H1000FFFF) And BitRight32(&H10000, m) Then
  
  If mvarBitTest32((LunarInfo(Y - 1900) And &H1000FFFF), 16 - m) Then
     lMonthDays = 30
  Else
     lMonthDays = 29
  End If
End Function

'传回农历 y年闰月的天数
Private Function leapDays(Y As Long) As Long
  If leapMonth(Y) Then
    If LunarInfo(Y - 1900) And &H10000 Then
      leapDays = 30
    Else
      leapDays = 29
    End If
  Else
    leapDays = 0
  End If
End Function

'传回农历 y年闰哪个月 1-12 , 没闰传回 0
Private Function leapMonth(Y As Long) As Long
  Dim i As Long
  i = LunarInfo(Y - 1900) And &HF
  If i > 12 Then
    Debug.Print Y
  End If
  leapMonth = i
End Function

'计算公历年月的天数
Private Function SolarDays(Y As Long, m As Long) As Long
   Dim d As Long
   
   If (Y Mod 4) = 0 Then   '闰年
      If m = 2 Then
        d = 29
      Else
        d = SolarMonth(m - 1)
      End If
   Else
      If m = 2 Then
        d = 28
      Else
        d = SolarMonth(m - 1)
      End If
   End If
   
   SolarDays = d
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////
'
'主要的函数,用公历年月日对日期对象进行初使化,在此函数内部完成对私有对象属性的设置
'
'//////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub sInitDate(ByVal Y As Long, ByVal m As Long, ByVal d As Long)
    Dim i As Long
    Dim leap As Long
    Dim temp As Long
    Dim offset As Long
    
    mvarDate = DateSerial(Y, m, d)
    mvarsYear = Y
    mvarsMonth = m
    mvarsDay = d
    
    '农历日期计算部分
    leap = 0
    temp = 0
    
    offset = mvarDate - #1/30/1900#  '计算两天的基本差距
    
    For i = 1900 To 2049
        temp = lYearDays(i)  '求当年农历年天数
        
        offset = offset - temp
        If offset < 1 Then Exit For
    Next
    
    offset = offset + temp
    mvarlYear = i
    
    leap = leapMonth(i) '闰哪个月
    
    mvarIsLeap = False
    For i = 1 To 12
        '闰月
        If leap > 0 And i = (leap + 1) And mvarIsLeap = False Then
           mvarIsLeap = True
           i = i - 1
           temp = leapDays(mvarlYear)   '计算闰月天数
        Else
           mvarIsLeap = False
           temp = lMonthDays(mvarlYear, i) '计算非闰月天数
        End If
    
        offset = offset - temp
        If offset <= 0 Then Exit For
    Next
    
    offset = offset + temp
    mvarlMonth = i
    mvarlDay = offset
        
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////
'
'主要的函数,用农历年月日对日期对象进行初使化,在此函数内部完成对私有对象属性的设置
'
'//////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub lInitDate(ByVal Y As Long, ByVal m As Long, ByVal d As Long, Optional LeapFlag As Boolean = False)
    Dim i As Long
    Dim leap As Long
    Dim temp As Long
    Dim offset As Long

    mvarlYear = Y
    mvarlMonth = m
    mvarlDay = d
    
    offset = 0
    
    For i = 1900 To Y - 1
       temp = LunarYearDays(i - 1900) '求当年农历年天数
       offset = offset + temp
    Next
    
    leap = leapMonth(Y) '闰哪个月
    If m <> leap Then
       mvarIsLeap = False  '当前日期并非闰月
    Else
       mvarIsLeap = LeapFlag  '使用用户输入的是否闰月月份
    End If
    
    If (m < leap) Or (leap = 0) Then   '当闰月在当前日期后
        For i = 1 To m - 1
           temp = lMonthDays(Y, i) '计算非闰月天数
           offset = offset + temp
        Next
    Else   '在闰月后
        If mvarIsLeap = False Then  '用户要计算非闰月的月份
             For i = 1 To m - 1
                temp = lMonthDays(Y, i) '计算非闰月天数
                offset = offset + temp
             Next
             If m > leap Then
                temp = leapDays(Y)   '计算闰月天数
                offset = offset + temp
             End If
             
        Else  '此时只有mvarisleap=ture,
               For i = 1 To m
                 temp = lMonthDays(Y, i) '计算非闰月天数
                 offset = offset + temp
               Next
        End If
    End If
    
    offset = offset + d '加上当月的天数
    mvarDate = DateAdd("d", offset, #1/30/1900#)
    mvarsYear = Year(mvarDate)
    mvarsMonth = Month(mvarDate)
    mvarsDay = Day(mvarDate)

End Sub



