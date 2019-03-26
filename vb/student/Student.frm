VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "学生信息管理系统"
   ClientHeight    =   5850
   ClientLeft      =   -1440
   ClientTop       =   2430
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   9030
   Begin VB.Frame fraBrowse 
      Caption         =   "浏览"
      Height          =   735
      Left            =   6840
      TabIndex        =   30
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   645
         TabIndex        =   34
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1050
         TabIndex        =   33
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1455
         TabIndex        =   32
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   400
      End
   End
   Begin VB.Frame fraManage 
      Caption         =   "管理"
      Height          =   3255
      Left            =   7560
      TabIndex        =   8
      Top             =   2280
      Width           =   1185
      Begin VB.CommandButton cmdEdit 
         Caption         =   "编辑(&E)"
         Height          =   435
         Left            =   120
         TabIndex        =   13
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "更新(&U)"
         Enabled         =   0   'False
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   2070
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   930
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加(&A)"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "报表(&R)"
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   975
      End
   End
   Begin VB.Frame fraInfo 
      Enabled         =   0   'False
      Height          =   4935
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   5895
      Begin VB.CommandButton cmdSelectPhoto 
         Caption         =   "选择图片(&S)"
         Height          =   375
         Left            =   4560
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1080
         TabIndex        =   20
         Top             =   767
         Width           =   1215
      End
      Begin VB.TextBox txtBirthday 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   1791
         Width           =   1200
      End
      Begin VB.TextBox txtAddress 
         Height          =   630
         Left            =   1080
         TabIndex        =   18
         Top             =   2770
         Width           =   3240
      End
      Begin VB.TextBox txtResume 
         Height          =   645
         Left            =   1080
         TabIndex        =   17
         Top             =   4080
         Width           =   3240
      End
      Begin VB.TextBox txtTelephone 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   3597
         Width           =   1800
      End
      Begin VB.ComboBox cboSex 
         Height          =   300
         ItemData        =   "Student.frx":0000
         Left            =   1080
         List            =   "Student.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2273
         Width           =   735
      End
      Begin VB.TextBox txtSerial 
         Height          =   330
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcbClass 
         Bindings        =   "Student.frx":0016
         Height          =   330
         Left            =   1080
         TabIndex        =   21
         Top             =   1264
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
         ListField       =   "Name"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   "Class"
      End
      Begin MSComDlg.CommonDialog dlgSelect 
         Left            =   4680
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgPhoto 
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Photo"
         DataMember      =   "Student"
         DataSource      =   "DataEnv"
         Height          =   2055
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名:"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   29
         Top             =   827
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "班级:"
         Height          =   180
         Index           =   11
         Left            =   480
         TabIndex        =   28
         Top             =   1339
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "学号:"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   27
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "出生日期:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   1843
         Width           =   810
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "性别:"
         Height          =   180
         Index           =   5
         Left            =   480
         TabIndex        =   25
         Top             =   2333
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "地址:"
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   24
         Top             =   2995
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "简历:"
         Height          =   300
         Index           =   8
         Left            =   480
         TabIndex        =   23
         Top             =   4252
         Width           =   450
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "电话:"
         Height          =   180
         Index           =   7
         Left            =   480
         TabIndex        =   22
         Top             =   3649
         Width           =   450
      End
   End
   Begin VB.Frame fraSeek 
      Caption         =   "查询"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmdSeek 
         Caption         =   "查询(&F)"
         Height          =   375
         Left            =   5400
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "列出>>"
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "根据所在的班级列出学籍信息"
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboClass 
         Height          =   300
         ItemData        =   "Student.frx":002D
         Left            =   3000
         List            =   "Student.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
      Begin VB.ComboBox cboDep 
         Height          =   300
         ItemData        =   "Student.frx":003E
         Left            =   960
         List            =   "Student.frx":0045
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "所在班："
         Height          =   180
         Left            =   2160
         TabIndex        =   5
         Top             =   330
         Width           =   840
      End
      Begin VB.Label lblDep 
         AutoSize        =   -1  'True
         Caption         =   "所在系："
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   330
         Width           =   720
      End
   End
   Begin MSDataGridLib.DataGrid grdScan 
      Bindings        =   "Student.frx":004F
      Height          =   4935
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   8705
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "sqlSeek"
      Caption         =   "导航条"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Serial"
         Caption         =   "学号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "姓名"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'标识是否能关闭
Dim mbClose As Boolean

'标识当前要显示的照片的文件
Dim mstrFileName As String

'当DataEnv.rsStudent的当前记录发生变化时，刷新所绑定的控件(用户改变了当前记录)
Sub RefreshBinding()
    On Error Resume Next
   With DataEnv.rsStudent
      If DataEnv.rssqlSeek.BOF And DataEnv.rssqlSeek.EOF Then
         '如果不存在任何记录，则清空所有的绑定的内容
         txtSerial = ""
         txtName = ""
         txtBirthday = ""
         txtTelephone = ""
         txtAddress = ""
         txtResume = ""
         imgPhoto.Picture = LoadPicture(Null)
      Else  '否则和相应的字段进行绑定
         txtSerial = .Fields("serial")
         txtName = .Fields("name")
         txtBirthday = .Fields("birthday")
         txtTelephone = .Fields("tel")
         txtAddress = .Fields("address")
         txtResume = .Fields("resume")
         cboSex.Text = .Fields("sex")
         dcbClass.Text = .Fields("class")
         imgPhoto.Picture = LoadPicture(ReadImage(.Fields("photo")))
      End If
   End With
End Sub

''在DataEnv.rsStudent中查询serial为sSerial的学籍信息
Sub SeekStudent(sSerial As String)
   If Not (DataEnv.rsStudent.EOF And DataEnv.rsStudent.BOF) Then
      Dim Temp As String
      Temp = "serial = " & "'" & sSerial & "'"
      
      DataEnv.rsStudent.MoveFirst
      DataEnv.rsStudent.Find Temp
      
      '刷新所绑定的控件
      Call RefreshBinding
  End If
End Sub

''当改变记录集时，需要刷新用户导航的网格控件
Sub RefreshGrid()
    grdScan.DataMember = ""
    grdScan.Refresh
    DataEnv.rssqlSeek.Requery
    grdScan.DataMember = "sqlSeek"
    grdScan.Refresh
    
    '刷新各个绑定控件
    Call grdScan_Change
End Sub

''用以在浏览时，根据当前记录所出的位置不同，来改变个浏览按钮的状态
Sub ChangeBrowseState()
   With DataEnv.rssqlSeek
      If .State = adStateClosed Then .Open
      '如果没有任何记录，使某些按钮无效；否则则使这些按钮有效
      If .BOF And .EOF Then
         cmdAdd.Enabled = True
         cmdEdit.Enabled = False
         cmdDelete.Enabled = False
         cmdUpdate.Enabled = False
         cmdReport.Enabled = False

         fraBrowse.Enabled = False
      Else
         cmdAdd.Enabled = True
         cmdEdit.Enabled = True
         cmdDelete.Enabled = True
         cmdUpdate.Enabled = False
         cmdReport.Enabled = True
         
         fraBrowse.Enabled = True
      End If
      
      ''假如处于记录的头部
      If .BOF Then
          If Not .EOF Then DataEnv.rsStudent.MoveFirst
          cmdPrevious.Enabled = False
          cmdFirst.Enabled = False
      Else
          cmdPrevious.Enabled = True
          cmdFirst.Enabled = True
      End If
      ''假如处于记录的尾部
      If .EOF Then
          If Not .BOF Then DataEnv.rsStudent.MoveLast
          cmdNext.Enabled = False
          cmdLast.Enabled = False
      Else
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      End If
    End With
    
    mstrFileName = ""
End Sub

Private Sub cboDep_Click()
    Dim rsClass As New ADODB.Recordset
    Dim strSQL
    '根据所选的系的不同，采用不同的SQL语句
    If cboDep.ItemData(cboDep.ListIndex) = 0 Then
        strSQL = "select * from class"
    Else
        strSQL = "select * from class where dept_id=" & cboDep.ItemData(cboDep.ListIndex)
    End If
    
    rsClass.Open strSQL, DataEnv.Con
    
    '将所查到的rsClass中的内容来填充cboClass
    cboClass.Clear
    cboClass.AddItem "全部"
    While Not rsClass.EOF
        cboClass.AddItem rsClass("Name")
        rsClass.MoveNext
    Wend
    cboClass.ListIndex = 0
    
    rsClass.Close
    Set rsClass = Nothing
End Sub

Private Sub cmdAdd_Click()
   '添加记录
   fraSeek.Enabled = False
   fraBrowse.Enabled = False
   grdScan.Enabled = False
    
   DataEnv.rsStudent.AddNew
   txtBirthday.Text = "1980-01-01"

   fraInfo.Enabled = True
   fraBrowse.Enabled = False
   
   cmdAdd.Enabled = False
   cmdEdit.Enabled = False
   cmdDelete.Enabled = False
   cmdUpdate.Enabled = True
   cmdReport.Caption = "取消"
   cmdReport.Enabled = True
   
   mbClose = False                   '不能关闭窗口
End Sub

Private Sub cmdDelete_Click()
    '如果出错，则显示错误代码
  On Error GoTo errHandler
  
  If MsgBox("要删除记录?", vbYesNo + vbQuestion + vbDefaultButton2, "确认") = vbYes Then
        '通过在DataEnv.Con中执行SQL命令，来删除记录
      DataEnv.Con.Execute "delete from student where serial ='" & txtSerial & "'"
      
      DataEnv.rsStudent.MoveNext
      If DataEnv.rsStudent.EOF Then DataEnv.rsStudent.MoveLast
      '刷新用户导航的网格控件
      Call RefreshGrid
  End If
  
  Exit Sub
  
errHandler:
  MsgBox Err.Description, vbCritical, "错误"
End Sub

Private Sub cmdEdit_Click()
    '编辑记录之前，需要设置其他控件的Enabled属性
    fraSeek.Enabled = False
    fraBrowse.Enabled = False
    grdScan.Enabled = False
    
    fraInfo.Enabled = True
      
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = True
      
    cmdReport.Caption = "取消"    ''更改cmdReport标题
    cmdReport.Enabled = True
      
    mbClose = False              '出于编辑状态，则用户不能关闭窗口
End Sub

Private Sub cmdFirst_Click()
    '移动到记录的头部，并改变各个浏览按钮的状态
    DataEnv.rssqlSeek.MoveFirst
    DataEnv.rssqlSeek.MovePrevious
    Call ChangeBrowseState
End Sub

Private Sub cmdLast_Click()
    '移动到记录的尾部，并改变各个浏览按钮的状态
    DataEnv.rssqlSeek.MoveLast
    DataEnv.rssqlSeek.MoveNext
    Call ChangeBrowseState
End Sub

Private Sub cmdList_Click()
    '针对所选的班级，列出班级中所有的学籍信息
    
    Dim strSQL
    If cboClass.Text = "全部" Then
        strSQL = " from student order by serial"
    Else
        strSQL = " from student where class='" & cboClass & "' order by serial"
    End If
    
    DataEnv.rsStudent.Close
    DataEnv.rsStudent.Open "select * " & strSQL
    
    DataEnv.rssqlSeek.Close
    DataEnv.rssqlSeek.Open "select serial, name " & strSQL
    
    
    '刷新用户导航的网格控件，并且根据记录集中记录的数目，来改变各个浏览按钮的状态。
    Call RefreshGrid
    Call ChangeBrowseState
    
    Call grdScan_Change
End Sub

Private Sub cmdNext_Click()     '移动到记录的下一条
    DataEnv.rssqlSeek.MoveNext
    Call ChangeBrowseState
End Sub

Private Sub cmdPrevious_Click() '移动到记录的上一条
    DataEnv.rssqlSeek.MovePrevious
    Call ChangeBrowseState
End Sub

Private Sub cmdReport_Click()
   On Error Resume Next
   If cmdReport.Caption = "取消" Then
      '取消所使用的更新更新
      DataEnv.rsStudent.CancelUpdate
      
      '重新显示原来数据集中的内容
      If DataEnv.rsStudent.BOF Then
         DataEnv.rsStudent.MoveFirst
      Else
         DataEnv.rsStudent.MovePrevious
         DataEnv.rsStudent.MoveNext
      End If
      Call RefreshBinding
      Call ChangeBrowseState
      
      fraSeek.Enabled = True
      fraBrowse.Enabled = True
      fraInfo.Enabled = False
      grdScan.Enabled = True
      cmdReport.Caption = "报表(R)"

      mbClose = True
   Else
    '生成报表
      Dim strSQL As String
      DataEnv.rsrptStudent.Close
      strSQL = "select * from student where serial = '" & txtSerial.Text & "'"
      DataEnv.rsrptStudent.Open strSQL
      
      rptStudent.Show
   End If
End Sub

Private Sub cmdSeek_Click()
   With frmFind
      Dim i As Integer
      '显示查找窗口
      Load frmFind
      
      '填充查找窗体的字段列表框
      .lstFields.Clear
      For i = 0 To DataEnv.rsStudent.Fields.Count - 1
        .lstFields.AddItem (DataEnv.rsStudent(i).Name)
      Next i
      .lstFields.ListIndex = 0
      .Show 1
      
      If .mbFindFailed Then Exit Sub
      
      Dim sTemp As String
      If LCase(.msFindOp) = "like" Then
          sTemp = .msFindField & " " & .msFindOp & " '%" & .msFindExpr & "%'"
      Else
          sTemp = .msFindField & " " & .msFindOp & " '" & .msFindExpr & "'"
      End If
      sTemp = "select * from student where " & sTemp & " order by serial"
      
      Unload frmFind
   End With
    
   '查找数据，并刷新用以导航的网格控件
    DataEnv.rssqlSeek.Close
    DataEnv.rssqlSeek.Open sTemp
    Call RefreshGrid
            
    Exit Sub
    
errHandler:
    MsgBox "没有符合条件的纪录！", vbExclamation, "确认"
End Sub

Private Sub cmdSelectPhoto_Click()
    On Error GoTo errHandler:
    
    dlgSelect.DialogTitle = "选择该学生的照片"
    dlgSelect.Filter = "所有图形文件|*.bmp;*.dib;*.gif;*.jpg;*.ico|位图文件(*.bmp;*.dib)|*.bmp;*.dib|GIF文件(*.gif)|*.gif|JPEG文件(*.jpg)|*.jpg|图标文件(*.ico)|*.ico"
    
    dlgSelect.ShowOpen
    
    If dlgSelect.FileName = "" Then Exit Sub

    imgPhoto.Picture = LoadPicture(dlgSelect.FileName)
    mstrFileName = dlgSelect.FileName
    
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "错误"
End Sub

Private Sub cmdUpdate_Click()
    '更新所添加或者修改的记录
   On Error GoTo errHandler:
   
   Dim str As String
   str = txtSerial.Text
   
   With DataEnv.rsStudent
      .Fields("Serial") = txtSerial.Text
      .Fields("name") = txtName.Text
      .Fields("sex") = cboSex.Text
      .Fields("class") = dcbClass.Text
      .Fields("birthday") = txtBirthday.Text
      .Fields("tel") = txtTelephone.Text
      .Fields("address") = txtAddress.Text
      .Fields("resume") = txtResume.Text
      
      Call WriteImage(.Fields("photo"), mstrFileName)
      .Update
   End With
   
   cmdReport.Caption = "报表(&R)"
   cmdUpdate.Enabled = False
   fraInfo.Enabled = False
   mbClose = True
   
   If DataEnv.rssqlSeek.State = adStateClosed Then DataEnv.rssqlSeek.Open
   '刷新右端用以导航的网格控件
   Call RefreshGrid
   '根据记录集中记录的个数，改变各个按钮的状态
   Call ChangeBrowseState
   
   '定位到刚刚添加或者修改过的记录
   DataEnv.rssqlSeek.MoveFirst
   DataEnv.rssqlSeek.Find "serial='" & str & "'"
   
   fraSeek.Enabled = True
   fraBrowse.Enabled = True
   grdScan.Enabled = True
   Exit Sub
  
errHandler:
  MsgBox Err.Description, vbCritical, " 错误"
End Sub

Private Sub dcbClass_Click(Area As Integer)
  If txtSerial = "" Then
     txtSerial = dcbClass.Text
  End If
End Sub

Private Sub Form_Load()
   On Error Resume Next
   
   Dim rsDep As New ADODB.Recordset, rsClass As New ADODB.Recordset
   Set rsDep = DataEnv.rsDepartment
   Set rsClass = DataEnv.rsClass
   
   '从Department表中读取数据，填充cboDep复合框到中
   rsDep.Open
   cboDep.Clear
   cboDep.AddItem "全部"
   '将各个系的id号作为ItemData附加到复合框中
   cboDep.ItemData(0) = 0
   While Not rsDep.EOF
       cboDep.AddItem rsDep("Name")
       cboDep.ItemData(cboDep.ListCount - 1) = rsDep("id")
       rsDep.MoveNext
   Wend
   cboDep.ListIndex = 0
   
   ''从class表中读取数据，填充到cboClass复合框中
   cboClass.Clear
   cboClass.AddItem "全部"
   While Not rsClass.EOF
       cboClass.AddItem rsClass("Name")
       rsClass.MoveNext
   Wend
   cboClass.ListIndex = 0
   
   cmdList.Value = True
      
   fraManage.Enabled = True
   fraBrowse.Enabled = True
   fraSeek.Enabled = True
   grdScan.Enabled = True
   
   mbClose = True
   
   Call grdScan_Change
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not mbClose Then
    MsgBox "数据正被修改，窗口不能关闭", vbCritical, "错误"
    Cancel = True
  End If
End Sub

Private Sub grdScan_Change()
   If grdScan.ApproxCount > 0 Then
      Call SeekStudent(grdScan.Columns(0).CellText(grdScan.Bookmark))
   End If
End Sub

Private Sub grdScan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   '当前行改变，则动态改变所要显示的记录
   If LastRow <> grdScan.Bookmark Then
      If grdScan.ApproxCount > 0 Then
         Call SeekStudent(grdScan.Columns(0).CellText(grdScan.Bookmark))
      End If
   End If
End Sub

Private Sub WriteImage(ByRef Fld As ADODB.Field, DiskFile As String)
    Dim byteData() As Byte '定义数据块数组
    Dim NumBlocks As Long '定义数据块个数
    Dim FileLength As Long '标识文件长度
    Dim LeftOver As Long '定义剩余字节长度
    Dim SourceFile As Long '定义自由文件号
    Dim i As Long '定义循环变量
    
    Const BLOCKSIZE = 4096 '每次读写块的大小
    
    SourceFile = FreeFile '提供一个尚未使用的文件号
    Open DiskFile For Binary Access Read As SourceFile '打开文件
    FileLength = LOF(SourceFile) '得到文件长度
    If FileLength = 0 Then '判断文件是否存在
        Close SourceFile
        MsgBox DiskFile & "无 内 容 或 不 存 在 !"
    Else
        NumBlocks = FileLength \ BLOCKSIZE '得到数据块的个数
        LeftOver = FileLength Mod BLOCKSIZE '得到剩余字节数
        Fld.Value = Null
        ReDim byteData(BLOCKSIZE) '重新定义数据块的大小
        For i = 1 To NumBlocks
            Get SourceFile, , byteData() ' 读到内存块中
            Fld.AppendChunk byteData() '写入FLD
        Next i
        
        ReDim byteData(LeftOver) '重新定义数据块的大小
        Get SourceFile, , byteData() '读到内存块中
        Fld.AppendChunk byteData() '写入FLD
        Close SourceFile '关闭源文件
    End If
End Sub

Private Function ReadImage(blobColumn As ADODB.Field) As String
    '取得一个临时性文件
    Dim strFileName As String
    strFileName = "ImageTmp"

    Dim FileNumber      As Integer      '文件号
    Dim DataLen             As Long         '文件长度
    Dim Chunks              As Long         '数据块数
    Dim ChunkAry()      As Byte         '数据块数组
    Dim ChunkSize       As Long         '数据块大小
    Dim Fragment        As Long         '零碎数据大小
    Dim lngI                As Long '计数器
    
    On Error GoTo errHander
    
    ChunkSize = 2048                    '定义块大小为 2K
    If IsNull(blobColumn) Then Exit Function

    DataLen = blobColumn.ActualSize         '获得图像大小
    If DataLen < 8 Then Exit Function   '图像大小小于8字节时认为不是图像信息
        FileNumber = FreeFile               '产生随机的文件号
    Open strFileName For Binary Access Write As FileNumber     '打开存放图像数据文件
    Chunks = DataLen \ ChunkSize        '数据块数
    Fragment = DataLen Mod ChunkSize    '零碎数据
    If Fragment > 0 Then            '有零碎数据，则先读该数据
            ReDim ChunkAry(Fragment - 1)
            ChunkAry = blobColumn.GetChunk(Fragment)
            Put FileNumber, , ChunkAry      '写入文件
    End If

    ReDim ChunkAry(ChunkSize - 1)             '为数据块重新开辟空间
    For lngI = 1 To Chunks                              '循环读出所有块
            ChunkAry = blobColumn.GetChunk(ChunkSize)   '在数据库中连续读数据块
            Put FileNumber, , ChunkAry()    '将数据块写入文件中
    Next lngI
    Close FileNumber            '关闭文件
    
    ReadImage = strFileName
    
    Exit Function
    
errHander:
    ReadImage = ""
End Function

