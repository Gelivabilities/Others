VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB����ܼ�_������ By_5mao"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16350
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   10350
   ScaleWidth      =   16350
   StartUpPosition =   2  '��Ļ����
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2730
      ItemData        =   "Form1.frx":030A
      Left            =   3960
      List            =   "Form1.frx":030C
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton Ƥ�� 
      Caption         =   "Ƥ��"
      Height          =   270
      Left            =   14400
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ���� 
      Caption         =   " ����"
      Height          =   270
      Left            =   13440
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox �������� 
      Height          =   270
      Left            =   960
      TabIndex        =   8
      ToolTipText     =   "�����뺺��ƴ������ĸ���в��Ҳ���"
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   9420
      ItemData        =   "Form1.frx":030E
      Left            =   120
      List            =   "Form1.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   3495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9975
      Width           =   16350
      _ExtentX        =   28840
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   15954
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton ������ 
      Caption         =   "������"
      Height          =   270
      Left            =   15360
      TabIndex        =   5
      Top             =   115
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10680
      Top             =   0
   End
   Begin VB.CommandButton ɾ�� 
      Caption         =   "ɾ��"
      Height          =   270
      Left            =   5640
      TabIndex        =   4
      Top             =   115
      Width           =   855
   End
   Begin VB.CommandButton ���� 
      Caption         =   "����"
      Height          =   270
      Left            =   3720
      TabIndex        =   3
      Top             =   115
      Width           =   855
   End
   Begin VB.CommandButton ��� 
      Caption         =   "���"
      Height          =   270
      Left            =   4680
      TabIndex        =   2
      Top             =   115
      Width           =   855
   End
   Begin VB.TextBox �������� 
      Height          =   9060
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      ToolTipText     =   "�˴�Ϊ���������"
      Top             =   840
      Width           =   12495
   End
   Begin VB.TextBox ������� 
      Height          =   270
      Left            =   3720
      TabIndex        =   0
      ToolTipText     =   "�˴�Ϊ����ı���"
      Top             =   480
      Width           =   12495
   End
   Begin VB.Label Label1 
      Caption         =   "�Զ�����"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   170
      Width           =   735
   End
   Begin VB.Menu pf 
      Caption         =   "Ƥ������"
      Visible         =   0   'False
      Begin VB.Menu pf_xia 
         Caption         =   "��һ��Ƥ��"
      End
      Begin VB.Menu pf_shang 
         Caption         =   "��һ��Ƥ��"
      End
      Begin VB.Menu pf_qx 
         Caption         =   "ȡ��"
      End
   End
   Begin VB.Menu fj 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu drfj 
         Caption         =   "���븽��"
      End
      Begin VB.Menu dcfj 
         Caption         =   "��������"
      End
      Begin VB.Menu sqfj 
         Caption         =   "ɾ������"
      End
      Begin VB.Menu fj_qx 
         Caption         =   "ȡ��"
      End
   End
   Begin VB.Menu gjx 
      Caption         =   "������"
      Visible         =   0   'False
      Begin VB.Menu Spy 
         Caption         =   "Spy++"
      End
      Begin VB.Menu ����ת�� 
         Caption         =   "����ת��"
      End
      Begin VB.Menu Postģ�� 
         Caption         =   "Postģ��"
      End
      Begin VB.Menu gjx_qx 
         Caption         =   "ȡ��"
      End
   End
   Begin VB.Menu Tray 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu Display 
         Caption         =   "��ʾ"
      End
      Begin VB.Menu exit 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection '�������ݿ����
Dim rs As New ADODB.Recordset  '���������
Dim ������� As String
Dim ��������() As String

Private Declare Function SkinH_SetAero Lib "SkinH.dll" (ByVal hwnd As Long) As Long
Private Declare Function SkinH_Attach Lib "SkinH.dll" () As Long
Private Declare Function SkinH_AttachEx Lib "SkinH.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
'-----------------------------she��ʽƤ��-----------------------------
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'------------------------------------����API
Private nfIconData As NOTIFYICONDATA
Const MAX_TOOLTIP As Integer = 50    '��ʾ�ַ�����Ԥ��ʾ�ĸ���
Const NIF_ICON = &H2                 'Ԥ��ӵ�ͼ��
Const NIF_MESSAGE = &H1              '�¼���Ϣ,�������̧�����
Const NIF_TIP = &H4                  'Ԥ��ʾ������
Const NIM_ADD = &H0                  '�������ͼ��
Const NIM_DELETE = &H2               'ɾ������ͼ��
Const WM_MOUSEMOVE = &H200           '����ƶ�
Const WM_LBUTTONDOWN = &H201         '�����Ҽ�
Const WM_LBUTTONUP = &H202           '���̧��
Const WM_LBUTTONDBLCLK = &H203       '���˫��
Const WM_RBUTTONDOWN = &H204         '�����Ҽ�
Const WM_RBUTTONUP = &H205           '�Ҽ�̧��
Const WM_RBUTTONDBLCLK = &H206       '�Ҽ�˫��
Const SW_RESTORE = 9                 '״̬�ָ�
Const SW_HIDE = 0                    '״̬����
'------------------------------------��������
Private Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type

Private Sub Display_Click() '��ʾ����
    Me.WindowState = 0          '��ԭ����
    Form1.Visible = True
    Form1.Show
End Sub

Private Sub dcfj_Click() '��������
    If StatusBar1.Panels(5).Text = "������Ϣ����" Then '�ж��Ƿ��и�������
        MsgBox "�������븽�������ڣ�", vbInformation, "VB����ܼ�"
    Else
        On Error GoTo ErrHandle          '�û�ȡ��ʱ��������
        CommonDialog1.CancelError = True
        CommonDialog1.FileName = ""
        CommonDialog1.DialogTitle = "��ѡ��zip��������·��"
        CommonDialog1.Flags = &H80000
        CommonDialog1.Filter = "zip�ļ�(*.zip) |*.zip"
        CommonDialog1.ShowSave
        If CommonDialog1.FileName <> "" Then
            Set mstream = New ADODB.Stream
            mstream.Type = adTypeBinary
            mstream.Open
            mstream.Position = 0
            mstream.Write rs.Fields("����").Value
            mstream.SaveToFile CommonDialog1.FileName, adSaveCreateOverWrite
            MsgBox "���������ɹ���", vbInformation, "VB����ܼ�"
        End If
        Exit Sub
ErrHandle:                       '������
        Select Case Err.Number
        Case 32755
            MsgBox "��δѡ���κ��ļ���", vbInformation, "VB����ܼ�"
        End Select
    End If
End Sub

Private Sub drfj_Click() '���븽��
    On Error GoTo ErrHandle          '�û�ȡ��ʱ��������
    CommonDialog1.CancelError = True
    CommonDialog1.FileName = ""
    CommonDialog1.DialogTitle = "��ѡ��Ҫ�����zipѹ���ļ�"
    CommonDialog1.Flags = &H80000
    CommonDialog1.Filter = "zip�ļ�(*.zip) |*.zip"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Set mstream = New ADODB.Stream
        mstream.Type = adTypeBinary
        mstream.Open
        mstream.LoadFromFile CommonDialog1.FileName
        rs.Fields("����").Value = mstream.Read
        rs.Update
        Call ����Ƿ��и�������
        MsgBox "��������ɹ���", vbInformation, "VB����ܼ�"
    End If
    Exit Sub
ErrHandle:                       '������
    Select Case Err.Number
    Case 32755
        MsgBox "��δѡ���κ��ļ���", vbInformation, "VB����ܼ�"
    End Select
End Sub

Private Sub sqfj_Click() 'ɾ������
    If StatusBar1.Panels(5).Text = "������Ϣ����" Then '�ж��Ƿ��и�������
        MsgBox "�������븽�������ڣ�", vbInformation, "VB����ܼ�"
    Else
        '����������������������������������������
        Dim v As String
        v = MsgBox("��ȷ��Ҫɾ������Ϊ:��" & List1.List(List1.ListIndex) & "���ĸ�����Ϣ��", vbOKCancel, "��ܰ��ʾ")
        If v = vbOK Then
            rs("����") = ""
            rs.Update
            Call ����Ƿ��и�������
            MsgBox "����ɾ���ɹ���", vbInformation, "VB����ܼ�"
        End If
        '������������������������������������������ʾ�Ƿ����Ҫɾ������
    End If
End Sub

Private Sub exit_Click() '�˳�
    Call �˳�
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lMsg As Single
    lMsg = x / Screen.TwipsPerPixelX
    If lMsg = WM_RBUTTONUP Then '��������Ҽ�
        Me.PopupMenu Tray            '�˵���ʾ�ڹ�괦
    End If
    
    If lMsg = WM_LBUTTONDBLCLK Then '������˫��
        Call Display_Click   '��ʾ
    End If
End Sub '���¼��еĴ���ֻ��������ϵ�ͼ��
'---------------------------------------------------��ʾ�˵�

Private Sub �������ͼ��()
    nfIconData.hwnd = Me.hwnd
    nfIconData.uID = Me.Icon
    nfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    nfIconData.uCallbackMessage = WM_MOUSEMOVE
    nfIconData.hIcon = Me.Icon.Handle
    nfIconData.szTip = "VB����ܼ�" & vbNullChar  'vbNullChar��ʾɾ���ұ߶��ڵĿո�
    nfIconData.cbSize = Len(nfIconData)
    
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)    '���ͼ�굽����
End Sub

Private Sub Form_Load() '�����ݿ�
    '��������������������������������������������������������������������
    If App.PrevInstance = True Then
        End
    End If
    '����������������������������������������������������������������������ֹ�ظ�����
    '��������������������������������������������������������������������
    If Dir(App.Path & "\SkinH.dll") = "" Then
        MsgBox "Ƥ��Dll�ļ������ڣ������Զ��˳���", vbInformation, "VB����ܼ�"
        End
    End If
    '���������������������������������������������������������������������ж�Ƥ��Dll
    '��������������������������������������������������������������������
    If Dir(App.Path & "\Ƥ��", vbDirectory) = "" Then
        MsgBox "Ƥ���ļ��в����ڣ������Զ��˳���", vbInformation, "VB����ܼ�"
        End
    End If
    '���������������������������������������������������������������������ж�Ƥ���ļ���
    '��������������������������������������������������������������������
    If Dir(App.Path & "\VB�������ݿ�.mdb") = "" Then
        MsgBox "���ݿ��ļ������ڣ������Զ��˳���", vbInformation, "VB����ܼ�"
        End
    End If
    '���������������������������������������������������������������������ж����ݿ�
    ��������.BackColor = RGB(238, 238, 238)
    �������.BackColor = RGB(238, 238, 238)
    ��������.BackColor = RGB(238, 238, 238)
    Call �������ͼ��
    
    On Error GoTo ErrHandle                    '������
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\VB�������ݿ�.mdb;Jet OLEDB:database password=admin" 'adminΪ���ݿ�����
    rs.Open "Select * From ����� ", db, 1, 3
ErrHandle:                                 '������
    
    Select Case Err.Number
    Case -2147217843
        MsgBox "���ݿ���������뽫���ݿ������޸�Ϊadmin�����ԣ�", vbExclamation, "������ʾ"
    End Select
    
    If rs.State = adStateOpen Then '������ӳɹ�
        Call ��������б��¼�
    End If
    
    '������������������������������������������������������������������������������������
    List2.Clear
    Dim abc As String
    Dim genmulu As String
    genmulu = App.Path & "\Ƥ��\"                      '·��,�ǵ�·������һ��Ҫ��"\"
    abc = Dir(genmulu, vbNormal)
    Do While abc <> ""
        If abc <> "." And abc <> ".." And Right(abc, 4) = ".she" Then
            List2.AddItem genmulu & abc
        End If
        abc = Dir                                      '�ٴε���dir����,��ʱ���Բ�������
    Loop
    
    If List2.ListCount > 0 Then '���Ƥ����Ϊ��
        List2.ListIndex = 0 'ѡ�е�һ��Ƥ��
        '��������������������������������������������������������
        If GetSetting(App.Title, "Settings", "pifu", "") <> "" Then '�������������������������������������һ�����ù�Ƥ��
            If Dir(GetSetting(App.Title, "Settings", "pifu", "")) <> "" Then '�����һ�����õ�Ƥ����Ȼ����
                Dim n As Integer
                For n = 0 To List2.ListCount - 1
                    '����������������������������������
                    If List2.List(n) = GetSetting(App.Title, "Settings", "pifu", "") Then
                        SkinH_AttachEx GetSetting(App.Title, "Settings", "pifu", ""), "" '������һ�����õ�Ƥ��
                        List2.ListIndex = n 'ѡ�е�n��Ƥ��
                    End If
                    '���������������������������������������ϴ�Ƥ����·��
                Next n
            End If
        Else
            List2.ListIndex = 0 'ѡ�е�һ��Ƥ��
        End If
        '��������������������������������������������������������Ƥ�����ط�ʽ
    Else
        MsgBox "Ƥ���ļ�ȱʧ�������Զ��˳���", vbInformation, "VB����ܼ�"
        End
    End If
    '����������������������������������������������������������������������������������������Ƥ���б�
End Sub

Private Sub pf_shang_Click() '�л�����һ��Ƥ��
    If List2.ListIndex = 0 Then
        MsgBox "��һ��Ƥ��Ϊ�գ����л�����һ�飡", vbInformation, "VB����ܼ�"
    Else
        List2.ListIndex = List2.ListIndex - 1
        SaveSetting App.Title, "Settings", "pifu", List2.List(List2.ListIndex) '����ѡ��Ƥ��
    End If
End Sub

Private Sub pf_xia_Click()   '�л�����һ��Ƥ��
    If List2.ListIndex = List2.ListCount - 1 Then
        MsgBox "��һ��Ƥ��Ϊ�գ����л�����һ�飡", vbInformation, "VB����ܼ�"
    Else
        List2.ListIndex = List2.ListIndex + 1
        SaveSetting App.Title, "Settings", "pifu", List2.List(List2.ListIndex) '����ѡ��Ƥ��
    End If
End Sub

Private Sub List2_Click()    '����Ƥ��
    SkinH_AttachEx List2.List(List2.ListIndex), ""
End Sub

Private Sub ��������б��¼�()
    Dim ��ǰ�� As Integer
    
    ������� = "�ڴ�����������"
    �������� = "�ڴ������������"
    List1.Clear '����б�
    
    If rs.RecordCount <> 0 Then '************************************1
        ReDim ��������(1 To rs.RecordCount) '���¶������鷶Χ
        rs.MoveFirst '----------------------ָ���һ��
        
        For ��ǰ�� = 1 To rs.RecordCount
            ��������(��ǰ��) = rs.Fields("����")
            rs.MoveNext '-----------------------ָ����һ��
        Next ��ǰ��
        
        If ��������.Text = "" Then '**********************2
            For ��ǰ�� = 1 To rs.RecordCount
                List1.AddItem ��������(��ǰ��)
            Next ��ǰ��
        Else '----------------------**********************2
            For ��ǰ�� = 1 To rs.RecordCount
                If InStr(UCase(test(��������(��ǰ��))), UCase(test(��������.Text))) > 0 Then
                    List1.AddItem ��������(��ǰ��)
                End If
            Next ��ǰ��
        End If '--------------------**********************2
        
    Else                        '************************************1
        MsgBox "�����ݿ�û���κμ�¼��", vbInformation, "��ܰ��ʾ"
    End If                      '************************************1
    
    '------------------------------------------------
    If List1.ListCount <> 0 Then
        List1.ListIndex = 0
        For ��ǰ�� = 0 To List1.ListCount - 1
            If ������� = List1.List(��ǰ��) Then
                List1.ListIndex = ��ǰ��
                If List1.ListCount - 1 - ��ǰ�� > 30 Then
                    List1.ListIndex = ��ǰ�� + 30: List1.ListIndex = ��ǰ��
                Else
                    List1.ListIndex = List1.ListCount - 1: List1.ListIndex = ��ǰ��
                End If
                Exit For
            End If
        Next ��ǰ��
    End If
    '------------------------------------------------ѡ��ĳ��
End Sub

Private Sub List1_Click()  '��ʾ��ǰ����
    If List1.ListCount > 0 Then '**************
        
        rs.MoveFirst                   'ָ���һ��
        rs.Find "����='" & List1.List(List1.ListIndex) & "'"
        If rs.EOF = False Then
            ������� = rs.Fields("����")
            �������� = rs.Fields("����")
            StatusBar1.Panels(4).Text = "�޸����ڣ�" & rs.Fields("�޸�����")
            Call ����Ƿ��и�������
        Else
            MsgBox "�Բ���ɶҲû�ҵ���", vbInformation, "����"
            Call ��������б��¼�
        End If
        
    End If '--------------------------------***************
    
    '--------------------------------------------
    'If Clipboard.GetFormat(1) = True Then Clipboard.SetText List1.List(List1.ListIndex)
    '--------------------------------------------���ƴ������
End Sub

Private Sub ����Ƿ��и�������()
    '��������������������������������������������������������������������
    On Error GoTo ErrHandle          '�û�ȡ��ʱ��������
    If CStr(rs.Fields("����")) <> "" Then
        StatusBar1.Panels(5).Text = "������Ϣ����"
    End If
    Exit Sub
ErrHandle:                                   '������
    StatusBar1.Panels(5).Text = "������Ϣ����"
    '������������������������������������������������������������������������Ƿ��и�������
End Sub

Private Sub Spy_Click()
    Form2.Show
End Sub

Private Sub ����ת��_Click()
    Form3.Show
End Sub

Private Sub Postģ��_Click()
    Form4.Show
End Sub

Private Sub Timer1_Timer()
    If List1.ListCount = 0 Then '****************1
        
        ɾ��.Enabled = False
        ����.Enabled = False
        ����.Enabled = False
        
        If ������� = "" Then '***3
            ���.Enabled = False
        Else                  '***3
            ���.Enabled = True
        End If                '***3
        
        StatusBar1.Panels(1).Text = "��ǰ��������0" '����һ�������������
        StatusBar1.Panels(2).Text = "��ǰѡ���У�0" '���ڶ��������������
    Else '-----------------------****************1
        ɾ��.Enabled = True
        
        If ������� = "" Or �������� = "" Then '*********2
            ����.Enabled = False
            ���.Enabled = False
            ����.Enabled = False
        Else                                   '*********2
            ����.Enabled = True
            ���.Enabled = True
            ����.Enabled = True
        End If                                 '*********2
        
        StatusBar1.Panels(1).Text = "��ǰ��������" & List1.ListCount                '����һ�������������
        StatusBar1.Panels(2).Text = "��ǰѡ���У�" & List1.ListIndex + 1            '���ڶ��������������
    End If '---------------------****************1
    
    If Me.WindowState = 1 Then '�����С������
        Call �������ͼ��          '���ͼ�굽����
        Form1.Visible = False
    End If
    
    StatusBar1.Panels(3).Text = "ϵͳʱ�䣺" & Now
    StatusBar1.Panels(6).Text = "��ϵ���ߣ�QQ:1668066802"
    
End Sub

Private Sub ����_Click()
    Me.PopupMenu fj, , ����.Left, ����.Top + ����.Height '��ʾ�����˵�
End Sub

Private Sub ������_Click()
    Me.PopupMenu gjx, , ������.Left, ������.Top + ������.Height '��ʾ������˵�
End Sub

Private Sub Ƥ��_Click()
    Me.PopupMenu pf, , Ƥ��.Left, Ƥ��.Top + Ƥ��.Height '��ʾƤ���б�˵�
End Sub

Private Sub �������_KeyPress(KeyAscii As Integer) '�����ַ�
    Dim a As String
    a = "`~!@#$%^&*_+-=[];'\./{}:|<>?" '������Ų����ܵ��ַ�
    If InStr(1, a, Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ��������_KeyPress(KeyAscii As Integer) '�����ַ�
    Dim a As String
    a = "`~!@#$%^&*_+-=[];'\./{}:|<>?" '������Ų����ܵ��ַ�
    If InStr(1, a, Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ��������_Change() '���˷�����������
    Call ��������б��¼�
End Sub

Private Sub ����_Click() '���¼�¼(���µ��ǵ�ǰ��)
    '_________________________________________________________________________________________________________1
    Dim v As String
    v = MsgBox("���º�ԭ�е����ݽ��ᱻ���ǣ��Ƿ������", vbOKCancel, "VB����ܼ�")
    If v = vbOK Then '________________________________________________________________________________________1
        Dim ��̬��ֵ As Integer
        Dim ��ǰ�� As Integer
        ��̬��ֵ = 1
        For ��ǰ�� = 1 To rs.RecordCount
            If ��������(��ǰ��) = �������.Text And ��������(��ǰ��) <> List1.List(List1.ListIndex) Then
                ��̬��ֵ = ��̬��ֵ + 1
            End If
        Next ��ǰ��
        
        If ��̬��ֵ = 1 Then '_____________2
            rs("����") = �������           '��Ӧ������
            rs("����") = ��������           '��Ӧ������
            rs("�޸�����") = Now            '��Ӧ�޸�ʱ����
            rs.Update
            ������� = �������
            �������� = ""
            Call ��������б��¼�
            MsgBox "������³ɹ���", vbOKOnly, "��ʾ"
        Else '_____________________________2
            MsgBox "�ñ����Ѵ��ڣ��������޸ı��������ӣ�", vbExclamation, "����"
            ������� = �������
            �������� = ""
            Call ��������б��¼�
        End If '___________________________2
    End If '__________________________________________________________________________________________________1
End Sub

Private Sub ɾ��_Click() 'ɾ����ǰ��
    Dim v As String
    v = MsgBox("��ȷ��Ҫɾ������Ϊ:��" & List1.List(List1.ListIndex) & "����������", vbOKCancel, "��ܰ��ʾ")
    If v = vbOK Then          'vbCancelҲ�ɻ���vbOK���ʾȷ����
        If List1.ListIndex > 1 Then
            ������� = List1.List(List1.ListIndex - 1)
        End If
        
        rs.Delete
        
        �������� = ""
        Call ��������б��¼�
    End If
End Sub

Private Sub ���_Click() '��Ӽ�¼
    Dim �ж� As Boolean
    Dim ��ǰ�� As Integer
    �ж� = True
    
    For ��ǰ�� = 1 To rs.RecordCount
        If ��������(��ǰ��) = �������.Text Then
            �ж� = False
            Exit For
        End If
    Next ��ǰ��
    
    If �ж� = True Then '***********1
        rs.AddNew
        rs("����") = �������           '��Ӧ������
        rs("����") = ��������           '��Ӧ������
        rs("�޸�����") = Now            '��Ӧ�޸�ʱ����
        rs.Update
        ������� = �������
        �������� = ""
        Call ��������б��¼�
        MsgBox "������ӳɹ���", vbOKOnly, "��ʾ"
    Else '--------------***********1
        MsgBox "�ñ����Ѵ��ڣ��������޸ı��������ӣ�", vbExclamation, "����"
        ������� = �������
        �������� = ""
        Call ��������б��¼�
    End If '------------***********1
End Sub

Private Sub Form_Unload(Cancel As Integer) '�ر����ݿ�
    Cancel = True  'ȡ���ر�
    Call �˳�
End Sub

Private Sub �˳�()
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData) '������ɾ��
    If rs.State = adStateOpen Then
        rs.Close '�رձ�
        db.Close '�ر����ݿ�
        
        Name App.Path & "\VB�������ݿ�.mdb" As App.Path & "\VB�������ݿ�2.mdb"
        Dim miJRO As JRO.JetEngine
        Set miJRO = New JRO.JetEngine
        miJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0; " & "Data Source=" & App.Path & "\VB�������ݿ�2.mdb;Jet OLEDB:Database Password=admin", _
        "Provider=Microsoft.Jet.OLEDB.4.0; " & "Data Source=" & App.Path & "\VB�������ݿ�.mdb;Jet OLEDB:Database Password=admin"
        Kill App.Path & "\VB�������ݿ�2.mdb"
    End If
End            '�ر�
End Sub
