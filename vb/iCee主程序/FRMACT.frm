VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FRMACT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00221C13&
   BorderStyle     =   0  'None
   Caption         =   "�"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   6210
      Picture         =   "FRMACT.frx":0000
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   6210
      Picture         =   "FRMACT.frx":00E4
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   6210
      Picture         =   "FRMACT.frx":01C8
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   4
      Top             =   15
      Width           =   750
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6735
      _extentx        =   11880
      _extenty        =   4895
      Begin MSWinsockLib.Winsock Winsock1 
         Index           =   2
         Left            =   4560
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Index           =   1
         Left            =   5280
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Index           =   0
         Left            =   5880
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   2775
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   6735
      _extentx        =   11880
      _extenty        =   4895
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   2775
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   6480
      Width           =   6735
      _extentx        =   11880
      _extenty        =   4895
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ�¶������λ"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1440
   End
End
Attribute VB_Name = "FRMACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private FName As String, FNAME2 As String, FNAME3 As String
Private Sub Form_Load()
Call PaintPng(App.Path & "\SKIN\A_T.PNG", Me.hdc, 8, 8)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H404000, B
FName = App.Path & "\USER\ACT\ACTIVE1.Bmp"      'ָ�������ļ�����·��
FNAME2 = App.Path & "\USER\ACT\ACTIVE2.Bmp"      'ָ�������ļ�����·��
FNAME3 = App.Path & "\USER\ACT\ACTIVE3.Bmp"      'ָ�������ļ�����·��
Call GETACT(0, frmma.Text3.Text)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub LA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
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
Sub GETACT(Index As Integer, IP As String)
    If Winsock1(Index).State <> sckClosed Then Winsock1(Index).Close                    '�ر�����
    Winsock1(Index).RemoteHost = IP     '��������ַ
    Winsock1(Index).RemotePort = 4567            '�������˿�
    Winsock1(Index).Connect                      '����
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim I As Integer
For I = 0 To Winsock1.Count - 1
    Winsock1(I).Close   '�ر�����
Next
End Sub
Private Sub Winsock1_Close(Index As Integer)
    Winsock1(Index).Close   '�ر�����
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim TheFile() As Byte                    '�������ݵ�����
    ReDim TheFile(bytesTotal)                '�ض��������½�
    Static YNLen As Boolean                  '�Ƿ�������ļ�����
    Dim I As Integer
    Dim Strs As String   '�����ļ������ַ���
    Select Case Index
    Case 0
    Winsock1(0).GetData TheFile                 '�����յ����ݱ��浽����
    If bytesTotal = 2 And Chr(TheFile(0)) = "C" And Chr(TheFile(1)) = "S" Then    '����յ����ǳɹ�������Ϣ
        'Me.Caption = "�ͻ���-----�ɹ�����"    '��ʾ��Ϣ
        Winsock1(0).SendData "GetFileLen"        '����Ҫ���ļ�������Ϣ
        Exit Sub          '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "N" And Chr(TheFile(1)) = "o" And Chr(TheFile(2)) = "F" Then '����յ������޴��ļ�����Ϣ
        Call SHOWWRONG("���������޴��ļ�", 2)
        Exit Sub     '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "T" And Chr(TheFile(1)) = "h" And Chr(TheFile(2)) = "E" Then  '����յ��ļ����ͽ�����Ϣ
        Close #1      '�ر��ļ�
        YNLen = False 'δ�����ļ�����������Ϣ
       Debug.Print "�ļ��ѳɹ�����"   '��ʾ��Ϣ
        Winsock1(0).SendData "ConClose"    '�ر�����
        Exit Sub
    End If
    If YNLen = True Then   '����Ѿ��������ļ�������Ϣ
        Put #1, , TheFile                '�����յ����ݰ�д����ļ�
        Winsock1(0).SendData "NextB"        '����Ҫ����һ���ݰ�����Ϣ
        Debug.Print bytesTotal  '�����ļ�����
    Else
        Debug.Print "���ڽ�������"      '��ʾ��Ϣ
        For I = 0 To bytesTotal - 1
            Strs = Strs & Chr(TheFile(I))   '����ļ����������ַ���
        Next I
        YNLen = True                   '�Ѿ��������ļ�����������Ϣ
        Winsock1(0).SendData "FLA"        '�����Ѿ��յ��ļ�����������Ϣ����Ϣ"FLA"
        Open FName For Binary As #1
        IW(0).IS_PIC = True
        IW(0).SETPIC FName
        Call GETACT(1, "127.0.0.1")
    End If
Case 1
    Winsock1(1).GetData TheFile                  '�����յ����ݱ��浽����
    If bytesTotal = 2 And Chr(TheFile(0)) = "C" And Chr(TheFile(1)) = "S" Then    '����յ����ǳɹ�������Ϣ
        'Me.Caption = "�ͻ���-----�ɹ�����"    '��ʾ��Ϣ
        Winsock1(1).SendData "GETFILE2"         '����Ҫ���ļ�������Ϣ
        Exit Sub          '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "N" And Chr(TheFile(1)) = "o" And Chr(TheFile(2)) = "F" Then '����յ������޴��ļ�����Ϣ
        Call SHOWWRONG("���������޴��ļ�", 2)
        Exit Sub     '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "T" And Chr(TheFile(1)) = "h" And Chr(TheFile(2)) = "E" Then  '����յ��ļ����ͽ�����Ϣ
        Close #1      '�ر��ļ�
        YNLen = False 'δ�����ļ�����������Ϣ
       Debug.Print "�ļ��ѳɹ�����"   '��ʾ��Ϣ
        Winsock1(1).SendData "ConClose"     '�ر�����
        Exit Sub
    End If
    If YNLen = True Then   '����Ѿ��������ļ�������Ϣ
        Put #1, , TheFile                '�����յ����ݰ�д����ļ�
        Winsock1(1).SendData "NextB"         '����Ҫ����һ���ݰ�����Ϣ
        Debug.Print bytesTotal  '�����ļ�����
    Else
        Debug.Print "���ڽ�������"      '��ʾ��Ϣ
        For I = 0 To bytesTotal - 1
            Strs = Strs & Chr(TheFile(I))   '����ļ����������ַ���
        Next I
        YNLen = True                   '�Ѿ��������ļ�����������Ϣ
        Winsock1(1).SendData "FLA"         '�����Ѿ��յ��ļ�����������Ϣ����Ϣ"FLA"
        Open FNAME2 For Binary As #1
        IW(1).IS_PIC = True
        IW(1).SETPIC FNAME2
        Call GETACT(2, "127.0.0.1")
    End If
Case 2
    Winsock1(2).GetData TheFile                  '�����յ����ݱ��浽����
    If bytesTotal = 2 And Chr(TheFile(0)) = "C" And Chr(TheFile(1)) = "S" Then    '����յ����ǳɹ�������Ϣ
        'Me.Caption = "�ͻ���-----�ɹ�����"    '��ʾ��Ϣ
        Winsock1(2).SendData "GETFILE3"         '����Ҫ���ļ�������Ϣ
        Exit Sub          '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "N" And Chr(TheFile(1)) = "o" And Chr(TheFile(2)) = "F" Then '����յ������޴��ļ�����Ϣ
        Call SHOWWRONG("���������޴��ļ�", 2)
        Exit Sub     '��������
    End If
    If bytesTotal = 3 And Chr(TheFile(0)) = "T" And Chr(TheFile(1)) = "h" And Chr(TheFile(2)) = "E" Then  '����յ��ļ����ͽ�����Ϣ
        Close #1      '�ر��ļ�
        YNLen = False 'δ�����ļ�����������Ϣ
       Debug.Print "�ļ��ѳɹ�����"   '��ʾ��Ϣ
        Winsock1(2).SendData "ConClose"     '�ر�����
        Exit Sub
    End If
    If YNLen = True Then   '����Ѿ��������ļ�������Ϣ
        Put #1, , TheFile                '�����յ����ݰ�д����ļ�
        Winsock1(2).SendData "NextB"         '����Ҫ����һ���ݰ�����Ϣ
        Debug.Print bytesTotal  '�����ļ�����
    Else
        Debug.Print "���ڽ�������"      '��ʾ��Ϣ
        For I = 0 To bytesTotal - 1
            Strs = Strs & Chr(TheFile(I))   '����ļ����������ַ���
        Next I
        YNLen = True                   '�Ѿ��������ļ�����������Ϣ
        Winsock1(2).SendData "FLA"         '�����Ѿ��յ��ļ�����������Ϣ����Ϣ"FLA"
        Open FNAME3 For Binary As #1
        IW(2).IS_PIC = True
        IW(2).SETPIC FNAME3
    End If
End Select
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock1(Index).Close   '�ر�����
    Call SHOWWRONG("��������æ,���Ժ�����!", 2): Unload Me
End Sub

