VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѧ����Ϣ����ϵͳ"
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
      Caption         =   "���"
      Height          =   735
      Left            =   6840
      TabIndex        =   30
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
      Caption         =   "����"
      Height          =   3255
      Left            =   7560
      TabIndex        =   8
      Top             =   2280
      Width           =   1185
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�༭(&E)"
         Height          =   435
         Left            =   120
         TabIndex        =   13
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "����(&U)"
         Enabled         =   0   'False
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   2070
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Height          =   435
         Left            =   120
         TabIndex        =   11
         Top             =   930
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���(&A)"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "����(&R)"
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
         Caption         =   "ѡ��ͼƬ(&S)"
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
         Caption         =   "����:"
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
         Caption         =   "�༶:"
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
         Caption         =   "ѧ��:"
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
         Caption         =   "��������:"
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
         Caption         =   "�Ա�:"
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
         Caption         =   "��ַ:"
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
         Caption         =   "����:"
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
         Caption         =   "�绰:"
         Height          =   180
         Index           =   7
         Left            =   480
         TabIndex        =   22
         Top             =   3649
         Width           =   450
      End
   End
   Begin VB.Frame fraSeek 
      Caption         =   "��ѯ"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton cmdSeek 
         Caption         =   "��ѯ(&F)"
         Height          =   375
         Left            =   5400
         TabIndex        =   35
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "�г�>>"
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "�������ڵİ༶�г�ѧ����Ϣ"
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
         Caption         =   "���ڰࣺ"
         Height          =   180
         Left            =   2160
         TabIndex        =   5
         Top             =   330
         Width           =   840
      End
      Begin VB.Label lblDep 
         AutoSize        =   -1  'True
         Caption         =   "����ϵ��"
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "sqlSeek"
      Caption         =   "������"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Serial"
         Caption         =   "ѧ��"
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
         Caption         =   "����"
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

'��ʶ�Ƿ��ܹر�
Dim mbClose As Boolean

'��ʶ��ǰҪ��ʾ����Ƭ���ļ�
Dim mstrFileName As String

'��DataEnv.rsStudent�ĵ�ǰ��¼�����仯ʱ��ˢ�����󶨵Ŀؼ�(�û��ı��˵�ǰ��¼)
Sub RefreshBinding()
    On Error Resume Next
   With DataEnv.rsStudent
      If DataEnv.rssqlSeek.BOF And DataEnv.rssqlSeek.EOF Then
         '����������κμ�¼����������еİ󶨵�����
         txtSerial = ""
         txtName = ""
         txtBirthday = ""
         txtTelephone = ""
         txtAddress = ""
         txtResume = ""
         imgPhoto.Picture = LoadPicture(Null)
      Else  '�������Ӧ���ֶν��а�
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

''��DataEnv.rsStudent�в�ѯserialΪsSerial��ѧ����Ϣ
Sub SeekStudent(sSerial As String)
   If Not (DataEnv.rsStudent.EOF And DataEnv.rsStudent.BOF) Then
      Dim Temp As String
      Temp = "serial = " & "'" & sSerial & "'"
      
      DataEnv.rsStudent.MoveFirst
      DataEnv.rsStudent.Find Temp
      
      'ˢ�����󶨵Ŀؼ�
      Call RefreshBinding
  End If
End Sub

''���ı��¼��ʱ����Ҫˢ���û�����������ؼ�
Sub RefreshGrid()
    grdScan.DataMember = ""
    grdScan.Refresh
    DataEnv.rssqlSeek.Requery
    grdScan.DataMember = "sqlSeek"
    grdScan.Refresh
    
    'ˢ�¸����󶨿ؼ�
    Call grdScan_Change
End Sub

''���������ʱ�����ݵ�ǰ��¼������λ�ò�ͬ�����ı�������ť��״̬
Sub ChangeBrowseState()
   With DataEnv.rssqlSeek
      If .State = adStateClosed Then .Open
      '���û���κμ�¼��ʹĳЩ��ť��Ч��������ʹ��Щ��ť��Ч
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
      
      ''���紦�ڼ�¼��ͷ��
      If .BOF Then
          If Not .EOF Then DataEnv.rsStudent.MoveFirst
          cmdPrevious.Enabled = False
          cmdFirst.Enabled = False
      Else
          cmdPrevious.Enabled = True
          cmdFirst.Enabled = True
      End If
      ''���紦�ڼ�¼��β��
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
    '������ѡ��ϵ�Ĳ�ͬ�����ò�ͬ��SQL���
    If cboDep.ItemData(cboDep.ListIndex) = 0 Then
        strSQL = "select * from class"
    Else
        strSQL = "select * from class where dept_id=" & cboDep.ItemData(cboDep.ListIndex)
    End If
    
    rsClass.Open strSQL, DataEnv.Con
    
    '�����鵽��rsClass�е����������cboClass
    cboClass.Clear
    cboClass.AddItem "ȫ��"
    While Not rsClass.EOF
        cboClass.AddItem rsClass("Name")
        rsClass.MoveNext
    Wend
    cboClass.ListIndex = 0
    
    rsClass.Close
    Set rsClass = Nothing
End Sub

Private Sub cmdAdd_Click()
   '��Ӽ�¼
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
   cmdReport.Caption = "ȡ��"
   cmdReport.Enabled = True
   
   mbClose = False                   '���ܹرմ���
End Sub

Private Sub cmdDelete_Click()
    '�����������ʾ�������
  On Error GoTo errHandler
  
  If MsgBox("Ҫɾ����¼?", vbYesNo + vbQuestion + vbDefaultButton2, "ȷ��") = vbYes Then
        'ͨ����DataEnv.Con��ִ��SQL�����ɾ����¼
      DataEnv.Con.Execute "delete from student where serial ='" & txtSerial & "'"
      
      DataEnv.rsStudent.MoveNext
      If DataEnv.rsStudent.EOF Then DataEnv.rsStudent.MoveLast
      'ˢ���û�����������ؼ�
      Call RefreshGrid
  End If
  
  Exit Sub
  
errHandler:
  MsgBox Err.Description, vbCritical, "����"
End Sub

Private Sub cmdEdit_Click()
    '�༭��¼֮ǰ����Ҫ���������ؼ���Enabled����
    fraSeek.Enabled = False
    fraBrowse.Enabled = False
    grdScan.Enabled = False
    
    fraInfo.Enabled = True
      
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = True
      
    cmdReport.Caption = "ȡ��"    ''����cmdReport����
    cmdReport.Enabled = True
      
    mbClose = False              '���ڱ༭״̬�����û����ܹرմ���
End Sub

Private Sub cmdFirst_Click()
    '�ƶ�����¼��ͷ�������ı���������ť��״̬
    DataEnv.rssqlSeek.MoveFirst
    DataEnv.rssqlSeek.MovePrevious
    Call ChangeBrowseState
End Sub

Private Sub cmdLast_Click()
    '�ƶ�����¼��β�������ı���������ť��״̬
    DataEnv.rssqlSeek.MoveLast
    DataEnv.rssqlSeek.MoveNext
    Call ChangeBrowseState
End Sub

Private Sub cmdList_Click()
    '�����ѡ�İ༶���г��༶�����е�ѧ����Ϣ
    
    Dim strSQL
    If cboClass.Text = "ȫ��" Then
        strSQL = " from student order by serial"
    Else
        strSQL = " from student where class='" & cboClass & "' order by serial"
    End If
    
    DataEnv.rsStudent.Close
    DataEnv.rsStudent.Open "select * " & strSQL
    
    DataEnv.rssqlSeek.Close
    DataEnv.rssqlSeek.Open "select serial, name " & strSQL
    
    
    'ˢ���û�����������ؼ������Ҹ��ݼ�¼���м�¼����Ŀ�����ı���������ť��״̬��
    Call RefreshGrid
    Call ChangeBrowseState
    
    Call grdScan_Change
End Sub

Private Sub cmdNext_Click()     '�ƶ�����¼����һ��
    DataEnv.rssqlSeek.MoveNext
    Call ChangeBrowseState
End Sub

Private Sub cmdPrevious_Click() '�ƶ�����¼����һ��
    DataEnv.rssqlSeek.MovePrevious
    Call ChangeBrowseState
End Sub

Private Sub cmdReport_Click()
   On Error Resume Next
   If cmdReport.Caption = "ȡ��" Then
      'ȡ����ʹ�õĸ��¸���
      DataEnv.rsStudent.CancelUpdate
      
      '������ʾԭ�����ݼ��е�����
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
      cmdReport.Caption = "����(R)"

      mbClose = True
   Else
    '���ɱ���
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
      '��ʾ���Ҵ���
      Load frmFind
      
      '�����Ҵ�����ֶ��б��
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
    
   '�������ݣ���ˢ�����Ե���������ؼ�
    DataEnv.rssqlSeek.Close
    DataEnv.rssqlSeek.Open sTemp
    Call RefreshGrid
            
    Exit Sub
    
errHandler:
    MsgBox "û�з��������ļ�¼��", vbExclamation, "ȷ��"
End Sub

Private Sub cmdSelectPhoto_Click()
    On Error GoTo errHandler:
    
    dlgSelect.DialogTitle = "ѡ���ѧ������Ƭ"
    dlgSelect.Filter = "����ͼ���ļ�|*.bmp;*.dib;*.gif;*.jpg;*.ico|λͼ�ļ�(*.bmp;*.dib)|*.bmp;*.dib|GIF�ļ�(*.gif)|*.gif|JPEG�ļ�(*.jpg)|*.jpg|ͼ���ļ�(*.ico)|*.ico"
    
    dlgSelect.ShowOpen
    
    If dlgSelect.FileName = "" Then Exit Sub

    imgPhoto.Picture = LoadPicture(dlgSelect.FileName)
    mstrFileName = dlgSelect.FileName
    
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "����"
End Sub

Private Sub cmdUpdate_Click()
    '��������ӻ����޸ĵļ�¼
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
   
   cmdReport.Caption = "����(&R)"
   cmdUpdate.Enabled = False
   fraInfo.Enabled = False
   mbClose = True
   
   If DataEnv.rssqlSeek.State = adStateClosed Then DataEnv.rssqlSeek.Open
   'ˢ���Ҷ����Ե���������ؼ�
   Call RefreshGrid
   '���ݼ�¼���м�¼�ĸ������ı������ť��״̬
   Call ChangeBrowseState
   
   '��λ���ո���ӻ����޸Ĺ��ļ�¼
   DataEnv.rssqlSeek.MoveFirst
   DataEnv.rssqlSeek.Find "serial='" & str & "'"
   
   fraSeek.Enabled = True
   fraBrowse.Enabled = True
   grdScan.Enabled = True
   Exit Sub
  
errHandler:
  MsgBox Err.Description, vbCritical, " ����"
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
   
   '��Department���ж�ȡ���ݣ����cboDep���Ͽ���
   rsDep.Open
   cboDep.Clear
   cboDep.AddItem "ȫ��"
   '������ϵ��id����ΪItemData���ӵ����Ͽ���
   cboDep.ItemData(0) = 0
   While Not rsDep.EOF
       cboDep.AddItem rsDep("Name")
       cboDep.ItemData(cboDep.ListCount - 1) = rsDep("id")
       rsDep.MoveNext
   Wend
   cboDep.ListIndex = 0
   
   ''��class���ж�ȡ���ݣ���䵽cboClass���Ͽ���
   cboClass.Clear
   cboClass.AddItem "ȫ��"
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
    MsgBox "���������޸ģ����ڲ��ܹر�", vbCritical, "����"
    Cancel = True
  End If
End Sub

Private Sub grdScan_Change()
   If grdScan.ApproxCount > 0 Then
      Call SeekStudent(grdScan.Columns(0).CellText(grdScan.Bookmark))
   End If
End Sub

Private Sub grdScan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   '��ǰ�иı䣬��̬�ı���Ҫ��ʾ�ļ�¼
   If LastRow <> grdScan.Bookmark Then
      If grdScan.ApproxCount > 0 Then
         Call SeekStudent(grdScan.Columns(0).CellText(grdScan.Bookmark))
      End If
   End If
End Sub

Private Sub WriteImage(ByRef Fld As ADODB.Field, DiskFile As String)
    Dim byteData() As Byte '�������ݿ�����
    Dim NumBlocks As Long '�������ݿ����
    Dim FileLength As Long '��ʶ�ļ�����
    Dim LeftOver As Long '����ʣ���ֽڳ���
    Dim SourceFile As Long '���������ļ���
    Dim i As Long '����ѭ������
    
    Const BLOCKSIZE = 4096 'ÿ�ζ�д��Ĵ�С
    
    SourceFile = FreeFile '�ṩһ����δʹ�õ��ļ���
    Open DiskFile For Binary Access Read As SourceFile '���ļ�
    FileLength = LOF(SourceFile) '�õ��ļ�����
    If FileLength = 0 Then '�ж��ļ��Ƿ����
        Close SourceFile
        MsgBox DiskFile & "�� �� �� �� �� �� �� !"
    Else
        NumBlocks = FileLength \ BLOCKSIZE '�õ����ݿ�ĸ���
        LeftOver = FileLength Mod BLOCKSIZE '�õ�ʣ���ֽ���
        Fld.Value = Null
        ReDim byteData(BLOCKSIZE) '���¶������ݿ�Ĵ�С
        For i = 1 To NumBlocks
            Get SourceFile, , byteData() ' �����ڴ����
            Fld.AppendChunk byteData() 'д��FLD
        Next i
        
        ReDim byteData(LeftOver) '���¶������ݿ�Ĵ�С
        Get SourceFile, , byteData() '�����ڴ����
        Fld.AppendChunk byteData() 'д��FLD
        Close SourceFile '�ر�Դ�ļ�
    End If
End Sub

Private Function ReadImage(blobColumn As ADODB.Field) As String
    'ȡ��һ����ʱ���ļ�
    Dim strFileName As String
    strFileName = "ImageTmp"

    Dim FileNumber      As Integer      '�ļ���
    Dim DataLen             As Long         '�ļ�����
    Dim Chunks              As Long         '���ݿ���
    Dim ChunkAry()      As Byte         '���ݿ�����
    Dim ChunkSize       As Long         '���ݿ��С
    Dim Fragment        As Long         '�������ݴ�С
    Dim lngI                As Long '������
    
    On Error GoTo errHander
    
    ChunkSize = 2048                    '������СΪ 2K
    If IsNull(blobColumn) Then Exit Function

    DataLen = blobColumn.ActualSize         '���ͼ���С
    If DataLen < 8 Then Exit Function   'ͼ���СС��8�ֽ�ʱ��Ϊ����ͼ����Ϣ
        FileNumber = FreeFile               '����������ļ���
    Open strFileName For Binary Access Write As FileNumber     '�򿪴��ͼ�������ļ�
    Chunks = DataLen \ ChunkSize        '���ݿ���
    Fragment = DataLen Mod ChunkSize    '��������
    If Fragment > 0 Then            '���������ݣ����ȶ�������
            ReDim ChunkAry(Fragment - 1)
            ChunkAry = blobColumn.GetChunk(Fragment)
            Put FileNumber, , ChunkAry      'д���ļ�
    End If

    ReDim ChunkAry(ChunkSize - 1)             'Ϊ���ݿ����¿��ٿռ�
    For lngI = 1 To Chunks                              'ѭ���������п�
            ChunkAry = blobColumn.GetChunk(ChunkSize)   '�����ݿ������������ݿ�
            Put FileNumber, , ChunkAry()    '�����ݿ�д���ļ���
    Next lngI
    Close FileNumber            '�ر��ļ�
    
    ReadImage = strFileName
    
    Exit Function
    
errHander:
    ReadImage = ""
End Function

