VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FRMDOWN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "下载"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   Icon            =   "FRMDOWN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PICSET 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   360
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   481
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox TXTPATH 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "APP.PATH"
         Top             =   4800
         Width           =   3735
      End
      Begin ICEE.ICEE_KEY ICL 
         Height          =   495
         Index           =   0
         Left            =   5040
         TabIndex        =   24
         Top             =   4680
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
      End
      Begin ICEE.ICHECK ICK 
         Height          =   975
         Index           =   0
         Left            =   720
         TabIndex        =   21
         Top             =   360
         Width           =   5535
         _extentx        =   9763
         _extenty        =   1720
      End
      Begin ICEE.ICHECK ICK 
         Height          =   975
         Index           =   1
         Left            =   720
         TabIndex        =   22
         Top             =   1440
         Width           =   5535
         _extentx        =   9763
         _extenty        =   1720
      End
      Begin ICEE.ICHECK ICK 
         Height          =   975
         Index           =   2
         Left            =   720
         TabIndex        =   23
         Top             =   2520
         Width           =   5535
         _extentx        =   9763
         _extenty        =   1720
      End
      Begin ICEE.ICEE_KEY ICL 
         Height          =   495
         Index           =   1
         Left            =   720
         TabIndex        =   26
         Top             =   5400
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
      End
      Begin ICEE.ICEE_KEY ICL 
         Height          =   495
         Index           =   2
         Left            =   2160
         TabIndex        =   27
         Top             =   5400
         Width           =   1335
         _extentx        =   2355
         _extenty        =   873
      End
      Begin ICEE.ICHECK ICK 
         Height          =   975
         Index           =   3
         Left            =   720
         TabIndex        =   28
         Top             =   3600
         Width           =   5535
         _extentx        =   9763
         _extenty        =   1720
      End
      Begin VB.Shape SB 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         Top             =   4680
         Width           =   4215
      End
   End
   Begin VB.PictureBox PKB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   240
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer TMP 
      Interval        =   1000
      Left            =   6960
      Top             =   960
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   1800
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
      Begin ICEE.ICEE_KEY ICS 
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   4440
         Width           =   1695
         _extentx        =   2990
         _extenty        =   873
      End
      Begin ICEE.ICEE_KEY ICS 
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   4440
         Width           =   1695
         _extentx        =   2990
         _extenty        =   873
      End
      Begin VB.Image IMEND 
         Height          =   255
         Left            =   4200
         ToolTipText     =   "关闭"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7125
      Picture         =   "FRMDOWN.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7125
      Picture         =   "FRMDOWN.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   15
      Width           =   750
   End
   Begin ICEE.ICEE_DOWNLOAD Downloader1 
      Index           =   0
      Left            =   9840
      Top             =   600
      _extentx        =   1085
      _extenty        =   1085
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMDOWN.frx":0552
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMDOWN.frx":222C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMDOWN.frx":3F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMDOWN.frx":5BE0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVIEW 
      Height          =   5985
      Left            =   330
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   10557
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LDONE 
      Height          =   5985
      Left            =   330
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   10557
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      NumItems        =   0
   End
   Begin ICEE.IList LstLog 
      Height          =   5985
      Left            =   345
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   7350
      _extentx        =   12965
      _extenty        =   10689
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7125
      Picture         =   "FRMDOWN.frx":78BA
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox PMAIN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   1320
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   9
      Top             =   2280
      Width           =   5655
      Begin VB.PictureBox PNEW 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001B27C9&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1440
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   15
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
         Begin VB.Label LBNEW 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   210
            TabIndex        =   16
            Top             =   120
            Width           =   90
         End
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1575
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   3255
         _extentx        =   8705
         _extenty        =   2778
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1575
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
         _extentx        =   9763
         _extenty        =   2778
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1575
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Top             =   2160
         Width           =   1575
         _extentx        =   3625
         _extenty        =   2778
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1575
         Index           =   3
         Left            =   3720
         TabIndex        =   13
         Top             =   2160
         Width           =   1575
         _extentx        =   2778
         _extenty        =   2778
      End
      Begin ICEE.ICEE_WIN8 IW 
         Height          =   1575
         Index           =   4
         Left            =   3720
         TabIndex        =   19
         Top             =   480
         Width           =   1575
         _extentx        =   2778
         _extenty        =   2778
      End
   End
   Begin VB.Label LC 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "17个记录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   2760
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下载任务"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1320
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "FRMDOWN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cf As String, ie8url, cf2 As String, a As String, ISMV As Boolean
Private Const MAX_PATH = 260
Public WithEvents m_NotificationWindow As FRMTIP
Attribute m_NotificationWindow.VB_VarHelpID = -1
Private Type BROWSEINFO
    hwndOwner      As Long
    pidlRoot       As Long
    pszDisplayName As Long
    lpszTitle      As String
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
Private Const BIF_NEWDIALOGSTYLE As Long = &H40 '有 "新建文件夹" 按钮
Private Const BIF_EDITBOX As Long = &H10 '含路径文本框

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Sub Downloader1_DownloadComplete(Index As Integer, Maxbytes As Long, SaveFile As String)
On Error Resume Next
If Len(Dir(SaveFile)) = 0 Then
    LVIEW.ListItems(Index + 1).SubItems(2) = "失败"
    LVIEW.ListItems(Index + 1).SmallIcon = 3
Else
    LVIEW.ListItems(Index + 1).SubItems(2) = "完成"
    LVIEW.ListItems(Index + 1).SmallIcon = 1
    LVIEW.ListItems(Index + 1).SubItems(2) = ""
    LVIEW.ListItems(Index + 1).SubItems(6) = "00:00:00"
    LDONE.ListItems.Add Index + 1, , LVIEW.ListItems(Index + 1).Text, 1, 1
     LDONE.ListItems(Index + 1).SubItems(1) = Now
     If Sound = 1 Then Call sndPlaySound(App.Path + "\Sound\DOWNLOAD_CO.WAV", 1)
     If PMAIN.Visible = True Then LBNEW.Caption = LBNEW.Caption + 1: PNEW.Visible = True
    If AUTO_OPEN_IT = True Then Call SYSTEMOPEN(Dpath & LVIEW.ListItems(Index + 1).Text)
    If AUTO_OPEN_FOLDER = True Then Shell "explorer.exe /select," & (Dpath & LVIEW.ListItems(Index + 1).Text), vbNormalFocus
    If AUTO_TIP = False Then Exit Sub
    If SaveFile = "" Then Exit Sub
    Call RequestUserNotification("NOTIFY:", "下载完成", Dpath & SaveFile, True)
    Dim lNotificationRequest      As cNotificationRequest
    Set lNotificationRequest = g_NotificationRequests.Item(1)
    Set m_NotificationWindow = New FRMTIP
    Call m_NotificationWindow.ShowNotification(lNotificationRequest)
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set g_NotificationRequests = Nothing
End Sub

Private Sub ICK_Click(Index As Integer)
lRet = SetInitEntry("System", "Weather", ICK(0).Value)
lRet = SetInitEntry("DOWNLOAD", "AUTO_OPEN", ICK(1).Value)
lRet = SetInitEntry("DOWNLOAD", "AUTO_FOLDER", ICK(2).Value)
lRet = SetInitEntry("DOWNLOAD", "AUTO_TIP", ICK(3).Value)
GETWEATHER = ICK(0).Value
If ICK(2).Value = 1 Then AUTO_OPEN_FOLDER = True Else AUTO_OPEN_FOLDER = False
If ICK(1).Value = 1 Then AUTO_OPEN_IT = True Else AUTO_OPEN_IT = False
If ICK(3).Value = 1 Then AUTO_TIP = True Else AUTO_TIP = False
End Sub

Private Sub ICL_Click(Index As Integer)
Select Case Index
Case 0
Dim Path As String
Path = BrowseForFolder
If Len(Path) <> 0 Then txtPath.Text = Path
lRet = SetInitEntry("DOWNLOAD", "PATH", Path)
Case 1
LDONE.ListItems.Clear
Case 2
LstLog.Clear
End Select
End Sub

Private Sub ICS_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Call frmma.保存一下(PO)
Case 1
Clipboard.Clear
Clipboard.SetData PO.PICTURE
End Select
End Sub
Private Sub Downloader1_DownloadProgress(Index As Integer, Curbytes As Long, Maxbytes As Long, Total As String)
If Maxbytes = 0 Then Exit Sub
LVIEW.ListItems(Index + 1).SubItems(2) = Int(Curbytes / Maxbytes * 100) & " %"
LVIEW.ListItems(Index + 1).SubItems(1) = Total
End Sub

Private Sub Downloader1_Speed(Index As Integer, Spe As String, Elapsed As String, Left As String)
On Error Resume Next
LVIEW.ListItems(Index + 1).SubItems(3) = Spe & "/s"
LVIEW.ListItems(Index + 1).SubItems(5) = Elapsed
LVIEW.ListItems(Index + 1).SubItems(6) = Left
End Sub

Private Sub Downloader1_State(Index As Integer, DString As String, SaveName As String)
LstLog.AddItem ("[" & Now & "]" & "[" & SaveName & "]:  " & DString), 0
End Sub

Private Sub Form_Activate()
ICK(0).Value = GetInitEntry("System", "Weather", 0)

Me.Cls
Me.BackColor = COLOR_NOR
Dim I As Integer
For I = 0 To ICS.Count - 1
ICS(I).SETCOLOR vbWhite, &H44DFE3, vbBlack
Next
For I = 0 To IW.Count - 1
IW(I).HASLINE = False
IW(I).SETCOLOR COLOR_HIGH, COLOR_HIGH
IW(I).SETTXTCOLOR vbWhite, vbWhite
Next
ICL(0).HASLINE = True

For I = 0 To ICL.Count - 1
ICL(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Next

For I = 0 To ICK.Count - 1
ICK(I).M_STYLE = 3
ICK(I).SETCOLOR COLOR_NOR, vbWhite
Next

SB.BackColor = COLOR_NOR
txtPath.BackColor = COLOR_NOR
PMAIN.BackColor = COLOR_NOR
PKB.BackColor = COLOR_NOR
PICSET.BackColor = COLOR_NOR
Me.LDONE.BackColor = COLOR_NOR
Me.LVIEW.BackColor = COLOR_NOR
IW(0).SETPNG App.Path & "\SKIN\AD.PNG", 20, (IW(0).Height - 64) / 2
IW(1).SETPNG App.Path & "\SKIN\DR.PNG", (IW(1).Width - 64) / 2, (IW(1).Height - 64) / 2
IW(2).SETPNG App.Path & "\SKIN\DA.PNG", (IW(2).Width - 64) / 2, (IW(2).Height - 64) / 2
IW(3).SETPNG App.Path & "\SKIN\DF.PNG", (IW(3).Width - 64) / 2, (IW(3).Height - 64) / 2
IW(4).SETPNG App.Path & "\SKIN\D_SET.PNG", (IW(4).Width - 64) / 2, (IW(4).Height - 64) / 2
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\EMP.PNG", Me.hdc, (Me.ScaleWidth - 200) / 2, (Me.ScaleHeight - 100) / 2)
Call PaintPng(App.Path & "\SKIN\D_T.PNG", Me.hdc, 8, 8)

PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)

LstLog.SETCOLOR COLOR_NOR, COLOR_HIGH
End Sub


Private Sub Form_Load()
'On Error Resume Next
Set g_NotificationRequests = New Collection

oldproc = GetWindowLong(txtPath.hwnd, GWL_WNDPROC)
SetWindowLong txtPath.hwnd, GWL_WNDPROC, AddressOf TextWndProc

Call SeekMe(Me)
ISMV = True
LBNEW.Caption = "0"
Call MOVENOW
ICS(0).SETTXT "保存"
ICS(1).SETTXT "复制到剪切板"

IW(2).SETFONT "微软雅黑", 12, False, 16, True
IW(1).SETFONT "微软雅黑", 12, False, 16, True

IW(3).SETTIP "下载日志"
IW(0).SETTIP "新建任务"
IW(4).SETTIP "下载设置"

ICL(0).SETTXT "浏览"
ICL(1).SETTXT "清空下载历史"
ICL(2).SETTXT "清空下载日志"

ICK(0).SETTXT "剪切板有可下载文件时自动下载"
ICK(1).SETTXT "下载完成自动打开文件"
ICK(2).SETTXT "下载完成后打开文件夹"
ICK(3).SETTXT "下载完成后显示提示框"

ICK(1).Value = GetInitEntry("DOWNLOAD", "AUTO_OPEN", 0)
If ICK(1).Value = 1 Then AUTO_OPEN_IT = True Else AUTO_OPEN_IT = False
ICK(2).Value = GetInitEntry("DOWNLOAD", "AUTO_FOLDER", 0)
If ICK(2).Value = 1 Then AUTO_OPEN_FOLDER = True Else AUTO_OPEN_FOLDER = False
ICK(3).Value = GetInitEntry("DOWNLOAD", "AUTO_TIP", 0)
If ICK(3).Value = 1 Then AUTO_TIP = True Else AUTO_TIP = False

LVIEW.ColumnHeaders.Add , , "文件名称", 280
LVIEW.ColumnHeaders.Add , , "文件大小", 55
LVIEW.ColumnHeaders.Add , , "进度", 55
LVIEW.ColumnHeaders.Add , , "速度", 55
LVIEW.ColumnHeaders.Add , , "保存路径", 0
LVIEW.ColumnHeaders.Add , , "已用时间", 0
LVIEW.ColumnHeaders.Add , , "剩余时间", 0
LVIEW.ColumnHeaders.Add , , "下载地址", 0
LDONE.ColumnHeaders.Add , , "文件名称", 270
LDONE.ColumnHeaders.Add , , "完成时间", 110
Call LoadList

txtPath.Text = GetInitEntry("DOWNLOAD", "PATH", App.Path & "\DOWNLOAD\")
Dpath = txtPath.Text


End Sub

Sub LoadList()
On Error Resume Next
Dim filem As String, tpStr As String, I As Integer
filem = App.Path & "\COFING\DLLIST.ini"
LVIEW.ListItems.Clear
If PathFileExists(filem) = 1 Then
Open filem For Input As #1
Do While Not EOF(1)
With LVIEW.ListItems.Add()
For I = 0 To 7
Line Input #1, tpStr
If I = 0 Then .Text = tpStr Else .SubItems(I) = tpStr
.SmallIcon = 4
.Icon = 4
Next
End With
Loop
Close #1
End If
filem = App.Path & "\COFING\NELOG.ini"
If PathFileExists(filem) = 1 Then
Open filem For Input As #1
Do Until EOF(1)
Input #1, tpStr
LstLog.AddItem Trim(tpStr) ' 将文件路径读入 隐藏的播放列表
Loop
Close
End If
filem = App.Path & "\COFING\DONELIST.ini"
If PathFileExists(filem) = 1 Then
LDONE.ListItems.Clear
Open filem For Input As #1
Do While Not EOF(1)
With LDONE.ListItems.Add()
For I = 0 To 1
Line Input #1, tpStr
If I = 0 Then .Text = tpStr Else: .SubItems(I) = tpStr
.SmallIcon = 1
.Icon = 1
Next
End With
Loop
Close #1
End If
LC.Caption = LVIEW.ListItems.Count & "个任务"
End Sub
Sub SAVELIST()
On Error Resume Next
Dim filem As String, I As Integer, tpList As ListItem
filem = (App.Path & "\COFING\DLLIST.ini")
Open filem For Output As #1
For Each tpList In LVIEW.ListItems
Print #1, tpList.Text
For I = 0 To 7
Print #1, tpList.SubItems(I)
Next
Next
Close #1
filem = (App.Path & "\COFING\DONELIST.ini")
Open filem For Output As #1
For Each tpList In LDONE.ListItems
Print #1, tpList.Text
For I = 0 To 1
Print #1, tpList.SubItems(I)
Next
Next
Close #1
    Dim IW As Long, temp As String
    filem = App.Path & "\COFING\NELOG.ini"
    For IW = 0 To LstLog.ListCount - 1
        If temp = "" Then
            temp = LstLog.List(IW)
        Else
            temp = temp & vbNewLine & LstLog.List(IW)
        End If
    Next IW
    Open filem For Output As #1
        Print #1, vbNewLine & temp
    Close #1
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
lRet = SetInitEntry("DOWNLOAD", "LEFT", Me.Left)
lRet = SetInitEntry("DOWNLOAD", "TOP", Me.Top)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub IW_Click(Index As Integer)
Select Case Index
Case 0
Frmadd.Show
Case 1
If LVIEW.ListItems.Count > 0 Then LVIEW.Visible = True Else LVIEW.Visible = False
PMAIN.Visible = False
PKB.Visible = True
PNEW.Visible = False
LBNEW.Caption = "0"
LA.Visible = True
LA.Caption = "下载任务"
LC.Visible = True
LC.Caption = LVIEW.ListItems.Count & "个任务"
Case 2
If LDONE.ListItems.Count > 0 Then LDONE.Visible = True Else LDONE.Visible = False
PKB.Visible = True
PMAIN.Visible = False
LA.Visible = True
LA.Caption = "下载完成"
LC.Visible = True
LC.Caption = LDONE.ListItems.Count & "个任务"
Case 3
PKB.Visible = True
If LstLog.ListCount > 0 Then LstLog.Visible = True Else LstLog.Visible = False
PMAIN.Visible = False
LA.Visible = True
LA.Caption = "下载日志"
LC.Visible = True
LC.Caption = ""
Case 4
PKB.Visible = True
LA.Caption = "设置"
LA.Visible = True
LC.Visible = True
LC.Caption = ""
PICSET.Visible = True
End Select

End Sub

Private Sub LA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LDONE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub LstLog_MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub LVIEW_DblClick()
On Error Resume Next
If Right(LVIEW.SelectedItem.SubItems(2), 1) = "%" Then Exit Sub
Call OpenFile(LVIEW.SelectedItem.SubItems(4))
End Sub

Private Sub LVIEW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub LVIEW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Me.PopupMenu Frmm.下载任务
End Sub

Private Sub IMEND_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IMEND.PICTURE <> Frmm.X3.PICTURE Then IMEND.PICTURE = Frmm.X3.PICTURE

End Sub

Private Sub IMEND_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMEND.PICTURE <> Frmm.X2.PICTURE Then IMEND.PICTURE = Frmm.X2.PICTURE
End Sub

Private Sub IMEND_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IMEND.PICTURE <> Frmm.X1.PICTURE Then IMEND.PICTURE = Frmm.X1.PICTURE
PO.Visible = False
End Sub

Private Sub PKB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ISMV = False Then
ISMV = True
PKB.Cls

Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PKB.hdc, 0, 0)
End If
End Sub

Private Sub PKB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PMAIN.Visible = True
LVIEW.Visible = False
LDONE.Visible = False
LstLog.Visible = False
PICSET.Visible = False
LC.Visible = False
LA.Visible = False
PO.Visible = False
PKB.Visible = False
End Sub
Private Sub PO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVENOW
End Sub

Private Sub TMP_Timer()
IW(1).SETTIP LVIEW.ListItems.Count
IW(2).SETTIP LDONE.ListItems.Count
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
If X3.Visible = False Then Call SAVELIST: Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub
Sub 筛选()
Dim fullPath$, jj%
fullPath = Clipboard.GetText
jj = InStrRev(fullPath, ".")
a = Mid$(fullPath, jj + 1, Len(fullPath) - jj)
If cf2 = Clipboard.GetText Then: Exit Sub
If a = "rar" Then: Call Meshow
If a = "exe" Then: Call Meshow
If a = "mp3" Then: Call Meshow
If a = "zip" Then: Call Meshow
If a = "swf" Then: Call Meshow
If a = "mp4" Then: Call Meshow
If a = "dll" Then: Call Meshow
If a = "mdb" Then: Call Meshow
If a = "wmv" Then: Call Meshow
If a = "rmvb" Then: Call Meshow
If a = "jpg" Then: Call Meshow
If a = "png" Then: Call Meshow
If a = "bmp" Then: Call Meshow
If a = "jpge" Then: Call Meshow
If a = "wav" Then: Call Meshow
If a = "psd" Then: Call Meshow
If a = "gif" Then: Call Meshow
End Sub
Private Sub Meshow()
cf2 = Clipboard.GetText
 Frmadd.Text1 = Clipboard.GetText
 'Frmadd.Show
 Call Frmadd.DOWNLOADIT
End Sub
Private Sub Meshow2()
 Frmadd.Text1 = GetIEAddressBarURL
 Frmadd.Show
 ie8url = ""
 cf = ie8url
End Sub
Private Function GetIEAddressBarURL() As String
Dim hwndIE As Long
Dim hwndWorker As Long
Dim hwndRebar As Long
Dim hwndAddrBand As Long
Dim hwndEdit As Long
Dim lpString As String * 256
hwndIE = FindWindow("IEFrame", vbNullString)
If hwndIE = 0 Then Exit Function
hwndWorker = FindWindowEx(hwndIE, 0, "WorkerW", vbNullString)
If hwndWorker = 0 Then Exit Function
hwndRebar = FindWindowEx(hwndWorker, 0, "ReBarWindow32", vbNullString)
If hwndRebar = 0 Then Exit Function
hwndAddrBand = FindWindowEx(hwndRebar, 0, "Address Band Root", vbNullString)
hwndEdit = FindWindowEx(hwndAddrBand, 0, "Edit", vbNullString)
SendMessage hwndEdit, WM_GETTEXT, 256, ByVal lpString
GetIEAddressBarURL = Replace(lpString, Chr$(0), "")
End Function
Private Sub geturl2()
Dim U, fullPath$, jj%
If cf = ie8url Then Exit Sub
fullPath = ie8url
jj = InStrRev(fullPath, ".")
U = Mid$(fullPath, jj + 1, Len(fullPath) - jj)
Select Case U
Case "apk", "flv", "rar", "zip", "exe", "mp3", "mp4", "swf", "dll", "wmv", "jpg", "png", "bmp", "rmvb", "wav", "mid"
Call Meshow2
End Select
End Sub
Sub 拦截IE()
ie8url = GetIEAddressBarURL
Call geturl2
End Sub

Sub 停止下载()
On Error Resume Next
If LVIEW.ListItems.Count = 0 Then Exit Sub
If Right((LVIEW.ListItems(LVIEW.SelectedItem.Index).SubItems(2)), 1) <> "%" Then Exit Sub
If Downloader1(LVIEW.SelectedItem.Index - 1).CloseDownload(LVIEW.ListItems(LVIEW.SelectedItem.Index).SubItems(4) & LVIEW.ListItems(LVIEW.SelectedItem.Index)) = True Then
LVIEW.ListItems(LVIEW.SelectedItem.Index).SubItems(2) = "停止"         '使用方法
LVIEW.ListItems(LVIEW.SelectedItem.Index).SmallIcon = 3
End If
End Sub
Sub 二维码()
On Error Resume Next
If LVIEW.SelectedItem.SubItems(7) = "" Then Call SHOWWRONG("链接是空白的,无法生成二维码!", 0): Exit Sub
frmGraphic.CreatQCode (LVIEW.SelectedItem.SubItems(7))
PO.Visible = True
PO.PICTURE = LoadPicture(App.Path & "\MEDIA\Paint\QCODE.JPG")
End Sub
Sub MOVENOW()
If ISMV = True Then
ISMV = False
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
End If
'If IMEND.PICTURE <> Frmm.X1.PICTURE Then IMEND.PICTURE = Frmm.X1.PICTURE
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub
Private Function BrowseForFolder() As String
Dim lpIDList As Long, sBuffer As String
Dim tBrowseInfo As BROWSEINFO
With tBrowseInfo
    .hwndOwner = Me.hwnd
    .lpszTitle = "请选择文件夹!" & vbNullChar
    .ulFlags = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
Else
    BrowseForFolder = ""
End If
End Function
Private Sub m_NotificationWindow_Finished()
    Set m_NotificationWindow = Nothing
End Sub

