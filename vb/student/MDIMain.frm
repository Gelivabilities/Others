VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "ѧ����Ϣ����ϵͳ"
   ClientHeight    =   6225
   ClientLeft      =   1800
   ClientTop       =   1815
   ClientWidth     =   6630
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuGeneral 
      Caption         =   "ͨ��(&G)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuStudent 
         Caption         =   "ѧ����Ϣ����(&S)..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFind 
         Caption         =   "ѧ����Ϣ��ѯ(&F)..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuTemp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "���µ�¼(&L)..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "����(&A)..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'����Դ������:http://www.newxing.com/

Option Explicit

'��ʾ��ǰ���û�����
'0---����Ա���͵��û�; 1---ѧ�����͵��û�
Public mnUserType As Integer
'��ʾ��ǰ��¼���û���
Public msUserName As String

Private Sub MDIForm_Activate()
'���ݲ�ͬ���û����ͣ�ʹ��Ӧ�Ĳ˵���ɼ�
  Select Case mnUserType
    Case 0:                       '�Թ���Ա��ݵ�¼
        mnuFind.Visible = True
    Case 1:                       '��ѧ����ݵ�¼�� ֻ�ܲ�ѯ�Լ�����Ϣ
        mnuFind.Visible = False
  End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("���Ҫ�Գ���ϵͳ��", vbQuestion + vbYesNo + vbDefaultButton2, "�˳�") = vbNo Then
    Cancel = 1
  End If
End Sub

Private Sub mnuAbout_Click()
 '��ʾ������...������
  Load frmSplash
  frmSplash.mbAbout = True
  frmSplash.Show vbModal
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuFind_Click()
   frmStudent.Show
   frmStudent.cmdSeek.Value = True
End Sub

Private Sub mnuLogin_Click()
  If MsgBox("�����µ�¼�����д��嶼���رգ��Ƿ����µ�¼��", _
    vbQuestion + vbYesNo + vbDefaultButton2, "���µ�¼") = vbYes Then
     Unload MDIMain
     frmLogin.Show
  End If
End Sub

Private Sub mnuStudent_Click()
   If mnUserType = 0 Then   '��Ϊ����Ա�û�
      frmStudent.Show
   Else                     '��Ϊѧ�����û�
      frmView.Show
   End If
End Sub
