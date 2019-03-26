VERSION 5.00
Begin VB.Form FRMSETINFO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "设置信息"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   Icon            =   "FRMSETINFO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   789
   Begin ICEE.ICEE_Calender ICD 
      Height          =   4440
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   7832
      Begin VB.Label LC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   5280
         MouseIcon       =   "FRMSETINFO.frx":038A
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   120
         Width           =   180
      End
   End
   Begin VB.TextBox TXTZONE 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      IMEMode         =   1  'ON
      Left            =   6840
      TabIndex        =   41
      Top             =   4650
      Width           =   4815
   End
   Begin VB.TextBox TXTMAIL 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      IMEMode         =   1  'ON
      Left            =   6840
      TabIndex        =   40
      Top             =   4080
      Width           =   4815
   End
   Begin VB.TextBox TXTQM 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      IMEMode         =   1  'ON
      Left            =   6840
      TabIndex        =   39
      Top             =   3600
      Width           =   4815
   End
   Begin VB.TextBox TXTTEL 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   720
      TabIndex        =   37
      Text            =   "15848025978"
      Top             =   8280
      Width           =   3975
   End
   Begin VB.TextBox txtADD 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   720
      TabIndex        =   33
      Text            =   "上海市"
      Top             =   6480
      Width           =   3975
   End
   Begin VB.TextBox TXTPHONE 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   720
      TabIndex        =   32
      Text            =   "15848025978"
      Top             =   7080
      Width           =   3975
   End
   Begin VB.TextBox TXTQQ 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   720
      TabIndex        =   31
      Text            =   "1043099405"
      Top             =   7680
      Width           =   3975
   End
   Begin VB.TextBox TXTNODE 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      IMEMode         =   1  'ON
      Left            =   6840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   5040
      Width           =   4815
   End
   Begin VB.ComboBox CBJOB 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5520
      Width           =   3975
   End
   Begin VB.ComboBox CBABO 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   6000
      Width           =   3975
   End
   Begin VB.ComboBox CBSTUDY 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4560
      Width           =   3975
   End
   Begin VB.ComboBox CBCUN 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   5040
      Width           =   3975
   End
   Begin VB.ComboBox CBYY 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3600
      Width           =   3975
   End
   Begin VB.ComboBox CBSX 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4080
      Width           =   3975
   End
   Begin VB.PictureBox C3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   11070
      Picture         =   "FRMSETINFO.frx":04DC
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   16
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   11070
      Picture         =   "FRMSETINFO.frx":05C0
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   11070
      Picture         =   "FRMSETINFO.frx":06A4
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   14
      Top             =   15
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   0
      Left            =   8040
      TabIndex        =   10
      Top             =   8760
      Width           =   1695
      _extentx        =   2990
      _extenty        =   873
   End
   Begin VB.ComboBox CBAGE 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   360
      TabIndex        =   8
      Text            =   "0"
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox TXTBIR 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "2013-08-09"
      Top             =   3120
      Width           =   4455
   End
   Begin VB.OptionButton SXM 
      BackColor       =   &H00404040&
      Caption         =   "男"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1530
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton SXF 
      BackColor       =   &H00404040&
      Caption         =   "女"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   1530
      Width           =   615
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   1
      Left            =   9960
      TabIndex        =   11
      Top             =   8760
      Width           =   1695
      _extentx        =   2990
      _extenty        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   2
      Left            =   8040
      TabIndex        =   12
      Top             =   2760
      Width           =   2220
      _extentx        =   2990
      _extenty        =   873
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "个性签名"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   9
      Left            =   6000
      TabIndex        =   44
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   10
      Left            =   6000
      TabIndex        =   43
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "个人主页"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   11
      Left            =   6000
      TabIndex        =   42
      Top             =   4680
      Width           =   720
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座机"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   15
      Left            =   240
      TabIndex        =   38
      Top             =   8280
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住址"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   36
      Top             =   6480
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "手机"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   13
      Left            =   240
      TabIndex        =   35
      Top             =   7080
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QQ"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   14
      Left            =   240
      TabIndex        =   34
      Top             =   7680
      Width           =   180
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   12
      Left            =   6000
      TabIndex        =   30
      Top             =   5280
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "职业"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   28
      Top             =   5520
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "血型"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   17
      Left            =   240
      TabIndex        =   27
      Top             =   6000
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "学历"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   6
      Left            =   240
      TabIndex        =   24
      Top             =   4560
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "国籍"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   5040
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "语言"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   16
      Left            =   240
      TabIndex        =   20
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "生肖"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   16
      X2              =   200
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "编辑个人信息"
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
      Index           =   18
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "生日"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   270
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年龄"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   360
   End
   Begin VB.Image STLOGO 
      Appearance      =   0  'Flat
      Height          =   2220
      Left            =   8040
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2220
   End
End
Attribute VB_Name = "FRMSETINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Sub 加载个人信息()
On Error Resume Next
If PathFileExists(App.Path & "\USER\" & frmma.Text1.Text & ".dat") = 1 Then
LA(3).Caption = frmma.Text1.Text
Open App.Path & "\USER\" & frmma.Text1.Text & ".dat" For Random As gFileNum Len = Len(MyPersonalInfo)
Get #gFileNum, 1, MyPersonalInfo
Dim Sex As String
Dim temp As String
If MyPersonalInfo.Sex = "Male" Then '男女
SXM.Value = True
SXF.Value = False
ElseIf MyPersonalInfo.Sex = "Female" Then
SXM.Value = False
SXF.Value = True
End If
TXTQM.Text = RTrim(MyPersonalInfo.Country)
CBAGE.Text = Replace((MyPersonalInfo.Age), " ", "")
TXTMAIL.Text = Replace((MyPersonalInfo.BIRTHDAY), " ", "")
TXTZONE.Text = Replace((MyPersonalInfo.Webpage), " ", "")
temp = Replace(Trim(MyPersonalInfo.About), "//crlf\\", vbCrLf)
TXTNODE.Text = Replace(temp, " ", "")
TXTPHONE.Text = MyPersonalInfo.PHONE
txtADD.Text = MyPersonalInfo.Address
TXTQQ.Text = MyPersonalInfo.QQ
CBYY.Text = MyPersonalInfo.language
CBJOB.Text = MyPersonalInfo.JOB
CBSTUDY.Text = MyPersonalInfo.STUDY
CBSX.Text = MyPersonalInfo.SX
TXTTEL.Text = MyPersonalInfo.TEL
CBABO.Text = MyPersonalInfo.OAB
CBCUN.Text = MyPersonalInfo.COU
MyPersonalInfo.Country = TXTQM.Text
If MyPersonalInfo.BIRTH = "" Then MyPersonalInfo.BIRTH = Date
TXTBIR.Text = MyPersonalInfo.BIRTH
Close #gFileNum
End If
ICM(2).SETTXT "开启生日提示"
End Sub
Sub SAVEINFO()
On Error Resume Next
If frmma.Winsock1.State <> 7 Then Exit Sub
Dim Sex As String
If SXM.Value = True Then Sex = "Male" Else Sex = "Female"
If TXTQM.Text = "" Then TXTQM.Text = "这个人很懒，什么都没留下"
If CBAGE.Text = "" Then CBAGE.Text = "未设置"
If TXTMAIL.Text = "" Then TXTMAIL.Text = "未设置"
If TXTZONE.Text = "" Then TXTZONE.Text = "未设置"
If TXTZONE.Text = "" Then TXTZONE.Text = "未设置"
MyPersonalInfo.PHONE = TXTPHONE.Text
MyPersonalInfo.Address = txtADD.Text
MyPersonalInfo.QQ = TXTQQ.Text
MyPersonalInfo.language = CBYY.Text
MyPersonalInfo.JOB = CBJOB.Text
MyPersonalInfo.STUDY = CBSTUDY.Text
MyPersonalInfo.SX = CBSX.Text
MyPersonalInfo.TEL = TXTTEL.Text
MyPersonalInfo.OAB = CBABO.Text
MyPersonalInfo.COU = CBCUN.Text
MyPersonalInfo.BIRTH = TXTBIR.Text
    Dim Temp4 As String
    Temp4 = Replace(TXTNODE.Text, vbCrLf, "//crlf\\")
    MyPersonalInfo.Sex = Sex '性别
    MyPersonalInfo.Country = RTrim(TXTQM.Text) '个性签名(原本是国家，在这里当个性签名)
    MyPersonalInfo.BIRTHDAY = Replace(TXTMAIL.Text, " ", "") '电子邮箱(原本是生日，这里作为email
    MyPersonalInfo.Age = Replace(CBAGE.Text, " ", "") '年龄
    MyPersonalInfo.Webpage = Replace(TXTZONE.Text, " ", "") '个人网站
    MyPersonalInfo.About = Replace(Temp4, " ", "") '个人说明
    Open App.Path & "\USER\" & frmma.Text1.Text & ".dat" For Random As gFileNum Len = Len(MyPersonalInfo)
    Put #gFileNum, 1, MyPersonalInfo
    Close #gFileNum
    frmma.Winsock1.SendData ".SaveInfo " & _
    frmma.Text1.Text & _
    " " & Sex & _
    " " & TXTQM.Text & _
    " " & CBAGE.Text & _
    " " & TXTMAIL.Text & _
    " " & TXTZONE.Text & _
    " " & Temp4 & _
    " " & CBJOB.Text & _
    " " & CBSTUDY.Text & _
    " " & txtADD.Text & _
    " " & CBCUN.Text & _
    " " & TXTPHONE.Text & _
    " " & TXTTEL.Text & _
    " " & TXTQQ.Text & _
    " " & CBYY.Text & _
    " " & CBABO.Text & _
    " " & CBSX.Text & _
    " " & TXTBIR.Text
    Call frmma.LoadInfo
End Sub

Private Sub CBABO_Change()
MyPersonalInfo.OAB = CBABO.Text
End Sub

Private Sub CBAGE_KeyPress(KeyAscii As Integer)
 KeyAscii = VailText(KeyAscii, "0123456789", True)
End Sub

Private Sub CBCUN_Click()
MyPersonalInfo.COU = CBCUN.Text
End Sub

Private Sub CBJOB_CLICK()
MyPersonalInfo.JOB = CBJOB.Text
End Sub

Private Sub CBSTUDY_CLICK()
MyPersonalInfo.STUDY = CBSTUDY.Text
End Sub

Private Sub CBSX_Click()
MyPersonalInfo.SX = CBSX.Text
End Sub

Private Sub CBYY_Click()
MyPersonalInfo.language = CBYY.Text
End Sub

Private Sub Form_Activate()
LOGO = GetSetting("ICEE", "Main", "logo", App.Path + "\Skin\DefaultHead.Bmp")
STLOGO.PICTURE = LoadPicture(LOGO)    '设置框中的头像
End Sub

Private Sub Form_Load()
If frmma.Left > Me.Width Then
Me.Move frmma.Left - Me.Width, frmma.Top
Else
Me.Move frmma.Left + frmma.Width, frmma.Top
End If

ICM(0).SETTXT "保存"
ICM(1).SETTXT "取消"
ICM(2).SETTXT "上传图片"
ICM(0).SETCOLOR Me.BackColor, COLOR_NOR, vbWhite
ICM(1).SETCOLOR Me.BackColor, COLOR_NOR, vbWhite
ICM(2).SETCOLOR Me.BackColor, COLOR_NOR, vbWhite

Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B

CBABO.AddItem "保密"
CBABO.AddItem "A型血"
CBABO.AddItem "B型血"
CBABO.AddItem "AB型血"
CBABO.AddItem "O型血"
CBABO.AddItem "其他血型"

Dim i As Integer
For i = 0 To 120
CBAGE.AddItem i
Next

With CBJOB
.AddItem "在校学生"
.AddItem "固定工作者"
.AddItem "自由自业者"
.AddItem "待业/无业/失业"
.AddItem "退休"
.AddItem "其他"
End With

With CBYY
.AddItem "中文"
.AddItem "英文"
.AddItem "法语"
.AddItem "西班牙语"
.AddItem "其他"
End With

With CBSX
.AddItem "鼠"
.AddItem "牛"
.AddItem "虎"
.AddItem "兔"
.AddItem "龙"
.AddItem "蛇"
.AddItem "马"
.AddItem "羊"
.AddItem "猴"
.AddItem "鸡"
.AddItem "狗"
.AddItem "猪"
End With

With CBSTUDY
.AddItem "小学及以下"
.AddItem "初中"
.AddItem "高中"
.AddItem "中专"
.AddItem "大专"
.AddItem "本科"
.AddItem "研究生"
.AddItem "博士及以上"
End With

With CBCUN
.AddItem "中国"
.AddItem "中国香港"
.AddItem "中国台湾"
.AddItem "日本"
.AddItem "美国"
.AddItem "俄罗斯"
.AddItem "蒙古"
.AddItem "泰国"
.AddItem "印度尼西亚"
.AddItem "波兰"
.AddItem "西班牙"
.AddItem "法国"
.AddItem "挪威"
.AddItem "其他"
End With

Call 加载个人信息

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub c1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = False
C2.Visible = True
End Sub
Private Sub c2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
C2.Visible = False
C3.Visible = True
End If
End Sub
Private Sub c3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C3.Visible = False
C1.Visible = True
If C3.Visible = False Then
Unload Me
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = True
C2.Visible = False
C3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnHook
End Sub

Private Sub ICD_Click()
TXTBIR.Text = ICD.mYear & "-" & ICD.mMonth & "-" & ICD.mDay
End Sub

Private Sub ICM_Click(Index As Integer)
Select Case Index
Case 0
Call SAVEINFO
Call 加载个人信息
Unload Me
Case 1
Unload Me
Case 2
FRMHEAD.Show
End Select
End Sub

Private Sub LB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Select Case Index
Case 1
On Error GoTo ERR:
If InStr(1, Clipboard.GetText, "http://") <> 1 Then Exit Sub
TXTZONE.Text = Clipboard.GetText
ERR:
Exit Sub
End Select
End Sub

Private Sub LB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub LC_Click()
ICD.Visible = False
End Sub


Private Sub STLOGO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FRMHEAD.Show
End Sub

Private Sub STLOGO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub

Private Sub TXTBIR_Change()
On Error Resume Next
Call 获得生肖(Left(TXTBIR.Text, 4), CBSX)
End Sub
Private Sub 获得生肖(Year As Integer, SHOWINFO As ComboBox)
On Error Resume Next
  Dim name As Integer
  name = Year Mod 12
  Select Case name
    Case 4
      SHOWINFO.Text = "鼠"
    Case 5
      SHOWINFO.Text = "牛"
    Case 6
      SHOWINFO.Text = "虎"
    Case 7
      SHOWINFO.Text = "兔"
    Case 8
      SHOWINFO.Text = "龙"
    Case 9
      SHOWINFO.Text = "蛇"
    Case 10
     SHOWINFO.Text = "马"
    Case 11
      SHOWINFO.Text = "羊"
    Case 0
     SHOWINFO.Text = "猴"
    Case 1
      SHOWINFO.Text = "鸡"
    Case 2
     SHOWINFO.Text = "狗"
    Case 3
     SHOWINFO.Text = "猪"
   End Select
End Sub

Private Sub TXTBIR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If ICD.Visible = True Then ICD.Visible = False Else ICD.Visible = True
End Sub

Private Sub TXTBIR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetHand
End Sub
Private Sub TXTMAIL_GotFocus()
TXTMAIL.SelStart = 0
TXTMAIL.SelLength = Len(TXTMAIL.Text)
End Sub

Private Sub TXTNODE_GotFocus()
TXTNODE.SelStart = 0
TXTNODE.SelLength = Len(TXTNODE.Text)
End Sub

Private Sub TXTPHONE_GotFocus()
TXTPHONE.SelStart = 0
TXTPHONE.SelLength = Len(TXTPHONE)
End Sub

Private Sub TXTPHONE_KeyPress(KeyAscii As Integer)
 KeyAscii = VailText(KeyAscii, "0123456789-+", True)
End Sub

Private Sub TXTQM_GotFocus()
TXTQM.SelStart = 0
TXTQM.SelLength = Len(TXTQM.Text)
End Sub

Private Sub TXTQQ_GotFocus()
TXTQQ.SelStart = 0
TXTQQ.SelLength = Len(TXTPHONE)
End Sub

Private Sub TXTQQ_KeyPress(KeyAscii As Integer)
 KeyAscii = VailText(KeyAscii, "0123456789", True)
End Sub

Private Sub TXTTEL_GotFocus()
TXTTEL.SelStart = 0
TXTTEL.SelLength = Len(TXTPHONE)
End Sub

Private Sub TXTTEL_KeyPress(KeyAscii As Integer)
 KeyAscii = VailText(KeyAscii, "0123456789-+", True)
End Sub

Private Sub TXTZONE_GotFocus()
TXTZONE.SelStart = 0
TXTZONE.SelLength = Len(TXTZONE.Text)
End Sub
