VERSION 5.00
Begin VB.Form FRMDATE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H005BB645&
   BorderStyle     =   0  'None
   Caption         =   "日历"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   828
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   11640
      Picture         =   "FRMDATE.frx":0000
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   13
      Top             =   30
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   11640
      Picture         =   "FRMDATE.frx":00E4
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   12
      Top             =   30
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   11640
      Picture         =   "FRMDATE.frx":01C8
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   11
      Top             =   30
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   5
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin VB.PictureBox PDATA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   801
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   12015
      Begin VB.PictureBox PBK 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0001A175&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   5055
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "FRMDATE.frx":02AC
         Top             =   1320
         Width           =   11535
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "今日详情"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   120
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   497
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   6
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   2
      Left            =   9360
      TabIndex        =   7
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   3
      Left            =   10800
      TabIndex        =   8
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   9
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "FRMDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim dateClass As New clsDate
Dim CanExit As Boolean
Dim oldWidth As Single
Dim oldHeight As Single
Dim curYear As Integer
Dim curMonth As Integer
Dim curDay As Integer
Dim tmpA As Single
Dim tmpB As Single
Dim tmpC As Single
Dim bi As Variant
Dim BI2 As Variant
Dim BI3 As Variant
Dim BI4 As Variant
Dim BI5 As Variant
Dim BI6 As Variant
Dim HideInfo As Boolean
Dim bool(0 To 1) As Boolean
Dim IS_MV As Boolean
Private Sub Form_Activate()
Me.BackColor = COLOR_NOR
PDATA.BackColor = COLOR_NOR
Text1.BackColor = COLOR_NOR
PBK.BackColor = COLOR_NOR
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
Dim i As Integer
For i = 0 To ICM.Count - 1
ICM(i).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next
ICM(0).SETTXT "上一年"
ICM(1).SETTXT "上一月"
ICM(2).SETTXT "下一月"
ICM(3).SETTXT "下一年"
ICM(4).SETTXT "今天"
End Sub

Private Sub Form_Load()
On Error Resume Next
    bi = Split(StrConv(LoadResData(315, "CUSTOM"), vbUnicode), vbCrLf)
    BI2 = Split(StrConv(LoadResData(316, "CUSTOM"), vbUnicode), vbCrLf)
    BI3 = Split(StrConv(LoadResData(317, "CUSTOM"), vbUnicode), vbCrLf)
    BI4 = Split(StrConv(LoadResData(318, "CUSTOM"), vbUnicode), vbCrLf)
    BI5 = Split(StrConv(LoadResData(319, "CUSTOM"), vbUnicode), vbCrLf)
    BI6 = Split(StrConv(LoadResData(320, "CUSTOM"), vbUnicode), vbCrLf)
    curYear = Year(Date$)
    curMonth = Month(Date$)
    curDay = Day(Date$)
    picMain.FontName = "微软雅黑"
End Sub
Sub MOVENOW()
X1.Visible = True
X2.Visible = False
X3.Visible = False
If IS_MV = True Then
IS_MV = False
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picMain.Move 0, 100, Me.ScaleWidth, Me.ScaleHeight - 100
    PDATA.Move picMain.Left, picMain.Top, picMain.Width, picMain.Height
    ReDrawCalendar
End Sub

Public Sub ReDrawCalendar()
    Dim i As Integer
    Dim j As Integer, X As Integer, Y As Integer
    Dim NumDays As Integer, CurrPos As Integer, bCurrMonth As Boolean
    Dim MonthStart As Date, Buffer As String
    Me.picMain.BackColor = vbWhite
    Me.picMain.Cls
    picMain.FOREColor = COLOR_HIGH
    picMain.Line (0, 0)-(Me.picMain.ScaleWidth, Me.picMain.TextHeight("星期") + 4), , BF
    picMain.FOREColor = COLOR_NOR
    picMain.Line (0, Me.picMain.TextHeight("星期") + 5)-(Me.picMain.ScaleWidth, Me.picMain.TextHeight("星期") + 5), , BF
    tmpA = Me.picMain.ScaleWidth / 7
    tmpB = (Me.picMain.ScaleHeight - Me.picMain.TextHeight("星期") - 4) / 6
    MonthStart = DateSerial(curYear, curMonth, 1)
    NumDays = DateDiff("d", MonthStart, DateAdd("m", 1, MonthStart))
    j = WeekDay(MonthStart) - 1
    j = j - 1
    For i = 1 To NumDays
        CurrPos = i + j
        X = 1 + (CurrPos Mod 7) * tmpA
        Y = Me.picMain.TextHeight("星期") + 5 + 1 + (CurrPos \ 7) * tmpB
        If i = curDay Then
            picMain.Font.Bold = True
        Else
            picMain.Font.Bold = False
        End If
        picMain.FOREColor = vbWhite
        picMain.Line (X, Y)-(X + tmpA, Y + tmpB), , BF
        Select Case WeekDay(DateSerial(curYear, curMonth, i), vbSunday)
            Case 1
                picMain.FOREColor = &H6826D5
            Case 7
                picMain.FOREColor = &H5BB645
            Case Else
                picMain.FOREColor = vbBlack
        End Select
        If curMonth = Month(Date) And i = Day(Date) And curYear = Year(Date) Then
            picMain.FOREColor = &HD19403
        End If
        dateClass.sInitDate curYear, curMonth, i
        If dateClass.sHolidayRecess = True Then picMain.FOREColor = &H6826D5
        picMain.CurrentX = X + 4
        picMain.CurrentY = Y + 4
        picMain.Print Format(i) & " " & dateClass.CDayStr(dateClass.lDay)
        If dateClass.sHoliday <> "" Then
            picMain.CurrentX = X + 4
            picMain.Print dateClass.sHoliday
        End If
        If dateClass.lHoliday <> "" Then
            picMain.CurrentX = X + 4
            picMain.Print dateClass.lHoliday
        End If
        If dateClass.lSolarTerm <> "" Then
            picMain.CurrentX = X + 4
            picMain.Print dateClass.lSolarTerm
        End If
    Next i
    picMain.FOREColor = COLOR_HIGH
    For i = 1 To 7
        picMain.Line (i * tmpA, 0)-(i * tmpA, Me.picMain.ScaleHeight)
    Next i
    For i = 1 To 6
        picMain.Line (0, Me.picMain.TextHeight("星期") + 4 + i * tmpB)-(Me.picMain.ScaleWidth, Me.picMain.TextHeight("星期") + 4 + i * tmpB)
    Next i
    picMain.FOREColor = vbWhite
    picMain.Font.Bold = False
    For i = 1 To 7
        picMain.CurrentX = (i - 1) * tmpA + 3
        picMain.CurrentY = 3
        picMain.Print WeekdayName(i, False, vbSunday)
    Next i
    picMain.FOREColor = RGB(182, 189, 210)
    picMain.FOREColor = vbBlack
    dateClass.sInitDate curYear, curMonth, curDay
  LA(1).Caption = Format(curYear) & "年" & Format(curMonth) & "月" & Format(curDay) & "日 " & dateClass.sWeekDayStr & "  " & dateClass.GanZhi(CLng(curYear)) & "(" & dateClass.YearAttribute(CLng(curYear)) & ") " & MonthName(dateClass.lMonth, False) & dateClass.CDayStr(dateClass.lDay)
UpdateInfo
    picMain.Line (0, 0)-(picMain.ScaleWidth - 1, picMain.ScaleHeight - 1), COLOR_NOR, B
End Sub

Public Function GetMonthDayCount(Year As Integer, Month As Integer) As Integer
    If Year Mod 4 = 0 Then
        Select Case Month
            Case 1, 3, 5, 7, 8, 10, 12
                GetMonthDayCount = 31
            Case 4, 6, 9, 11
                GetMonthDayCount = 30
            Case 2
                GetMonthDayCount = 29
        End Select
    Else
        Select Case Month
            Case 1, 3, 5, 7, 8, 10, 12
                GetMonthDayCount = 31
            Case 4, 6, 9, 11
                GetMonthDayCount = 30
            Case 2
                GetMonthDayCount = 28
        End Select
    End If
End Function

Private Sub ICM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Select Case Index
        Case 0
            curDay = 1
            curYear = curYear - 1
            If curYear = 1900 Then curYear = 1901
        Case 1
            curDay = 1
            curMonth = curMonth - 1
            If curMonth = 0 Then curMonth = 12: curYear = curYear - 1
            If curYear = 1900 Then curYear = 1901
        Case 2
            curDay = 1
            curMonth = curMonth + 1
            If curMonth = 13 Then curMonth = 1: curYear = curYear + 1
            If curYear = 2050 Then curYear = 2049
        Case 3
            curDay = 1
            curYear = curYear + 1
            If curYear = 2050 Then curYear = 2049
        Case 4
            curYear = Year(Date$)
            curMonth = Month(Date$)
            curDay = Day(Date$)
    End Select
    ReDrawCalendar
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PBK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = False Then
IS_MV = True
PBK.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PBK.hdc, 0, 0)
End If
End Sub

Private Sub PBK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PDATA.Visible = False: picMain.Visible = True
End Sub

Private Sub PDATA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PDATA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub picMain_DblClick()
PDATA.Visible = True: picMain.Visible = False
End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            curDay = curDay + 1
            If curDay > GetMonthDayCount(curYear, curMonth) Then
                curDay = 1
                curMonth = curMonth + 1
                If curMonth > 12 Then
                    curMonth = 1
                    curYear = curYear + 1
                    If curYear = 2050 Then curYear = 2049
                End If
            End If
        Case vbKeyLeft
            curDay = curDay - 1
            If curDay <= 0 Then
                curMonth = curMonth - 1
                If curMonth <= 0 Then
                    curMonth = 12
                    curYear = curYear - 1
                    If curYear = 1900 Then curYear = 1901
                End If
                curDay = GetMonthDayCount(curYear, curMonth)
            End If
        Case vbKeyDown
            curDay = curDay + 7
            If curDay > GetMonthDayCount(curYear, curMonth) Then
                curDay = 1
                curMonth = curMonth + 1
                If curMonth > 12 Then
                    curMonth = 1
                    curYear = curYear + 1
                    If curYear = 2050 Then curYear = 2049
                End If
            End If
        Case vbKeyUp
            curDay = curDay - 7
            If curDay <= 0 Then
                curMonth = curMonth - 1
                If curMonth <= 0 Then
                    curMonth = 12
                    curYear = curYear - 1
                    If curYear = 1900 Then curYear = 1901
                End If
                curDay = GetMonthDayCount(curYear, curMonth)
            End If
        Case vbKeyPageUp
            curDay = 1
            curMonth = curMonth - 1
            If curMonth = 0 Then curMonth = 12: curYear = curYear - 1
            If curYear = 1900 Then curYear = 1901
        Case vbKeyPageDown
            curDay = 1
            curMonth = curMonth + 1
            If curMonth = 13 Then curMonth = 1: curYear = curYear + 1
            If curYear = 2050 Then curYear = 2049
    End Select
    ReDrawCalendar
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim i As Integer, MaxDay As Integer
    If Y < Me.picMain.TextHeight("星期") + 5 + 1 Then Exit Sub
    i = WeekDay(DateSerial(curYear, curMonth, 1)) - 1
    i = ((((X + 1) \ tmpA) + 1) + (((Y - (Me.picMain.TextHeight("星期") + 5 + 1)) \ tmpB) * 7)) - i
    MaxDay = GetMonthDayCount(curYear, curMonth)
    If i >= 1 And i <= MaxDay Then
        curDay = i
    End If
    ReDrawCalendar
End Sub
Private Sub UpdateInfo()
    Dim s As String
    Dim i As Integer
    Dim a As String
    dateClass.sInitDate curYear, curMonth, curDay
    s = ""
    s = s & "===================================" & vbCrLf
    s = s & "日程信息 (" & Format(curYear) & "年" & Format(curMonth) & "月" & Format(curDay) & "日)" & vbCrLf
    s = s & "===================================" & vbCrLf
    s = s & "年份:" & dateClass.Era(CLng(curYear)) & vbCrLf
    s = s & "公历:" & Format(curYear) & "年" & Format(curMonth) & "月" & Format(curDay) & "日 " & dateClass.sWeekDayStr & vbCrLf
    s = s & "农历:" & dateClass.GanZhi(CLng(curYear)) & "(" & dateClass.YearAttribute(CLng(curYear)) & ")" & "年" & IIf(dateClass.IsLeap, "闰", "") & MonthName(dateClass.lMonth, False) & dateClass.CDayStr(dateClass.lDay) & vbCrLf
    s = s & "===================================" & vbCrLf
    s = s & "公历节日:" & vbCrLf & dateClass.sHoliday & " " & dateClass.wHoliday & vbCrLf
    s = s & "农历节日:" & vbCrLf & dateClass.lHoliday & vbCrLf
    s = s & "===================================" & vbCrLf
    s = s & BI5(curDay - 1) & vbCrLf
    s = s & "===================================" & vbCrLf
    s = s & "生日花语 - "
    a = Format(curMonth) & "月" & Format(curDay) & "日"
    For i = 0 To UBound(bi)
        If InStr(bi(i), a) Then
            s = s & Replace(BI4(i), "\n", vbCrLf) & vbCrLf
            Exit For
        End If
    Next i
    s = s & "===================================" & vbCrLf
    s = s & "星座:" & dateClass.Constellation(CLng(curMonth), CLng(curDay)) & "座(" & dateClass.Constellation2(CLng(curMonth), CLng(curDay)) & ")" & vbCrLf
    a = Format(curMonth) & "月" & Format(curDay) & "日"
    For i = 0 To UBound(bi)
        If InStr(bi(i), a) Then
            s = s & bi(i) & vbCrLf & vbCrLf
            Exit For
        End If
    Next i
    
    a = dateClass.Constellation(CLng(curMonth), CLng(curDay)) & "座"
    For i = 0 To UBound(BI6)
        If Left(BI6(i), 3) = a Then
            s = s & Replace(BI6(i), "$", vbCrLf) & vbCrLf & vbCrLf
            Exit For
        End If
    Next i
    
    s = s & vbCrLf
    a = dateClass.Constellation(CLng(curMonth), CLng(curDay)) & "座的男人和女人"
    For i = 0 To UBound(BI2)
        If InStr(BI2(i), a) Then
            s = s & Replace(BI2(i), "\n", vbCrLf) & vbCrLf
            Exit For
        End If
    Next i
    s = s & "===================================" & vbCrLf
    s = s & "属相:"
    a = dateClass.YearAttribute(CLng(curYear))
    For i = 0 To UBound(BI3)
        If InStr(BI3(i), a) Then
            s = s & Replace(BI3(i), "\n", vbCrLf) & vbCrLf
            Exit For
        End If
    Next i
    Text1.Text = s
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub x1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then Unload Me
End Sub
