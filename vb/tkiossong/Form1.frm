VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[̫�Ĥ��_��]iOS������Ϣ��ѯ��"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ListBox List2 
      Height          =   5100
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ֻ��ʾ����"
      Height          =   300
      Left            =   2880
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "ħ��"
      Height          =   1335
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   2895
      Begin VB.Label Label5 
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "���ѣ��ɣ�"
      Height          =   1335
      Left            =   4200
      TabIndex        =   4
      Top             =   1560
      Width           =   2895
      Begin VB.Label Label4 
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ͨ����"
      Height          =   1335
      Left            =   4200
      TabIndex        =   3
      Top             =   3000
      Width           =   2895
      Begin VB.Label Label3 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�򵥣�÷����"
      Height          =   1335
      Left            =   4200
      TabIndex        =   2
      Top             =   4440
      Width           =   2895
      Begin VB.Label Label2 
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Sub Check1_Click()

If Check1.Value = 1 Then '��ѡ�У�ֻ��ʾ����
List2.Visible = True '��ʾ���׵�list��ɼ�
Check1.Left = 1680  '����λ�ã���ֹ������ѡ��combobox���ص�
Combo2.Visible = True '����ѡ��combobox
List2.ListIndex = 1 'ͬһ�б��ͬһ�����ѡ�е�2������ʱ���ǲ����κβ����ģ������л��б��ʱ��ѡ�е���һ��list�ĸ��������Ի�����һ��list
List2.ListIndex = 0 'ѡ�е�һ�ֱ����ʾ��������
List1.Visible = False '������ʾ���и�����list�򣬷�ֹͨ������tab���Ӵ�������Ķ���
Else
If Check1.Value = 0 Then 'ȡ��ѡ�У���ʾ��������
List1.Visible = True '��ʾ���и�����list��ɼ�
List1.ListIndex = 1 'ͬ��
List1.ListIndex = 0 'ͬ��
List2.Visible = False '������ʾ���׵�list�򣬷�ֹͨ������tab���Ӵ�������Ķ���
Check1.Left = 2880 '�ƻ�ԭλ��������û��ô��Ť
Combo2.ListIndex = 0 'ѡ�ر����棬û�������Ļ�������ղ��������棬ѡ����û���׵ĸ裬ħ�����ﻹ�����ף��Ͳ���ȷ��
Combo2.Visible = False '��Ȼ���и����г�����ÿҳ��һ�����Ƿ����׵ĸ裬�ɴ��������
Else
End If
End If

End Sub


Private Sub Combo1_Click()

  Combo2.ListIndex = 0 '����ף�ԭ���check1��ͬ
 Combo2.Visible = False 'ԭ���check1��ͬ
 x = Combo1.ListIndex 'x=ϵ�ж�Ӧ���
 listthesongs (x) '����ϵ�ж�Ӧ��ţ������Զ��и�������
 List1.ListIndex = 0 'ѡ�е�һ�׸裬��Ȼ�л���ԭ�����׸����Ի��ڻ�ܱ�Ť


End Sub

Private Sub Combo2_Click()
If List2.Visible = False Then '��ʵ���ǵȼ���û�й�ѡֻ��ʾ���ף�list2�����ڣ��������ı�list1��ֵ
    If Combo2.ListIndex = 1 Then 'ѡ��������
        Frame4.Caption = "ħ�����" 'ħ�����ӡ������һ���ȡ������
        Label5.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List1.Text, "lind", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List1.Text, "lilj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List1.Text, "litj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "licx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "ligc", App.Path & "\songdata.ini")
    Else 'ѡ�˱�����
         Frame4.Caption = "ħ��" 'ħ��ȥ���������һ���ȡħ����Ĭ��Ϊ��
        Label5.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List1.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List1.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List1.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "mwgc", App.Path & "\songdata.ini")
    End If
Else
        If Combo2.ListIndex = 1 Then '�ȼ��ڹ�ѡ��ֻ��ʾ���ף��������ı�list2��ֵ������ͬ��
        Frame4.Caption = "ħ�����"
        Label5.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List2.Text, "lind", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List2.Text, "lilj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List2.Text, "litj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "licx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "ligc", App.Path & "\songdata.ini")
    Else
         Frame4.Caption = "ħ��"
        Label5.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List2.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List2.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List2.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "mwgc", App.Path & "\songdata.ini")
    End If
End If
End Sub

Private Sub Form_Load()
'����9�����������������������ѡ��
Combo1.AddItem "ȫ��", 0
Combo1.AddItem "Namcoԭ��", 1
Combo1.AddItem "JPOP", 2
Combo1.AddItem "�ŵ�", 3
Combo1.AddItem "��Ϸ", 4
Combo1.AddItem "����", 5
Combo1.AddItem "��ͯ", 6
Combo1.AddItem "V��", 7
Combo1.AddItem "����", 8

'������������ʹcombo2�б�������ѡ��
Combo2.AddItem "������", 0
Combo2.AddItem "������", 1

'��combobox����ʼ��Ϊѡ���һ��
Combo1.ListIndex = 0
Combo2.ListIndex = 0

'�г�ȫ������������������Combo1.ListIndex = 0���Ӧ
listthesongs (0)

'ѡ��һ�׸裬ԭ���check1
List1.ListIndex = 0

'����ʱ���г�ȫ��������list1�к����׵ĸ������뵽list2
        i = 1
        Do While i <= 259 '����һ������ô����
        li = ini.mfncGetFromIni(List1.List(i), "li", App.Path & "\songdata.ini") '��ȡini�ļ��е�������Ϣ��li=1��ʾ�ø躬����
        If li = "1" Then '�жϳ�����������
        List2.AddItem List1.List(i) '������׸�����list2
        Else
        End If
        i = i + 1
        Loop
        
If List2.ListCount <= 1 Or List1.ListCount <= 1 Then '����check1Ҫѡ���ڶ��׸裨ԭ���check1���������������ѡ��ᵼ�³���Ϊ��������������ִ���ʽ
MsgBox "��ȡ���ִ��󡣿���ԭ�����£�" & vbCrLf & "1�������б�������ļ�������" & vbCrLf & "2��������Ϣ�ļ������쳣" '��ʾ����
Unload Me '��������
End If

End Sub
Public Function listthesongs(x) As Integer '�г�����
List1.Clear '��գ���Ȼֻ������ظ���Ӹ���
Select Case x 'x��ֵ��combo1��ѡ�еķ�����Ŷ�Ӧ

    Case 0 'һ���������г����и���
                i = 1
        Do While i <= 68
            List1.AddItem ini.mfncGetFromIni("Namcoԭ��", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
         i = 1
        Do While i <= 82
            List1.AddItem ini.mfncGetFromIni("JPOP", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
        i = 1
        Do While i <= 11
            List1.AddItem ini.mfncGetFromIni("�ŵ�", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
      
        i = 1
        Do While i <= 24
            List1.AddItem ini.mfncGetFromIni("��Ϸ", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 59
            List1.AddItem ini.mfncGetFromIni("����", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 2
            List1.AddItem ini.mfncGetFromIni("��ͯ", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 4
            List1.AddItem ini.mfncGetFromIni("V��", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop

        i = 1
        Do While i <= 9
            List1.AddItem ini.mfncGetFromIni("����", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 1 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 68
            List1.AddItem ini.mfncGetFromIni("Namcoԭ��", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 2 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 82
            List1.AddItem ini.mfncGetFromIni("JPOP", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 3 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 11
            List1.AddItem ini.mfncGetFromIni("�ŵ�", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 4 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 24
            List1.AddItem ini.mfncGetFromIni("��Ϸ", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 5 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 59
            List1.AddItem ini.mfncGetFromIni("����", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 6 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 2
            List1.AddItem ini.mfncGetFromIni("��ͯ", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 7 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 4
            List1.AddItem ini.mfncGetFromIni("V��", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
    Case 8 'ֻ�г��÷���ĸ���
        i = 1
        Do While i <= 9
            List1.AddItem ini.mfncGetFromIni("����", "song" & i, App.Path & "\songlist.ini")
            i = i + 1
        Loop
        
 End Select
End Function

Private Sub List1_Click()
    Combo2.ListIndex = 0 '����ף�ԭ��ͬcheck1
    
    If List1.Text = "" Then 'ȱ�ٸ����ļ�ʱ��ȡ�����ĸ���������ֿհף�����հ��������Ϣ��Ϊ��
    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Else '��ȡ�������ɹ�������������һ�ļ���ȡ��������
    Label2.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List1.Text, "jdnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List1.Text, "jdlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List1.Text, "jdtj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "jdcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "jdgc", App.Path & "\songdata.ini")
    Label3.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List1.Text, "ptnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List1.Text, "ptlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List1.Text, "pttj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "ptcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "ptgc", App.Path & "\songdata.ini")
    Label4.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List1.Text, "knnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List1.Text, "knlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List1.Text, "kntj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "kncx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "kngc", App.Path & "\songdata.ini")
    Label5.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List1.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List1.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List1.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List1.Text, "mwgc", App.Path & "\songdata.ini")

    sc = ini.mfncGetFromIni(List1.Text, "sc", App.Path & "\songdata.ini") '��ȡ�׳�
    End If
    If sc = "" Then 'û�׳ƣ����������ʾ���ԣ�������ʾ�׳�
        Label1.Caption = "�������ƣ�" & List1.Text & vbCrLf & "BPM��" & ini.mfncGetFromIni(List1.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf
    Else '���׳ƣ����������ʾ�׳ƺ���������
        Label1.Caption = "�������ƣ�" & List1.Text & vbCrLf & "BPM��" & ini.mfncGetFromIni(List1.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf & "�׳ƣ�" & ini.mfncGetFromIni(List1.Text, "sc", App.Path & "\songdata.ini")
    End If
    
'ԭ����form1��check1����
    li = ini.mfncGetFromIni(List1.Text, "li", App.Path & "\songdata.ini")
    If li = "1" Then
        Combo2.Visible = True
        Check1.Left = 1680
    Else
    Combo2.Visible = False
    Check1.Left = 2880
    End If
End Sub

Private Sub List2_Click() 'ԭ��ͬlist1
    Combo2.ListIndex = 0
    Label2.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List2.Text, "jdnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List2.Text, "jdlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List2.Text, "jdtj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "jdcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "jdgc", App.Path & "\songdata.ini")
    Label3.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List2.Text, "ptnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List2.Text, "ptlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List2.Text, "pttj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "ptcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "ptgc", App.Path & "\songdata.ini")
    Label4.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List2.Text, "knnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List2.Text, "knlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List2.Text, "kntj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "kncx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "kngc", App.Path & "\songdata.ini")
    Label5.Caption = "�Ѷȣ����" & ini.mfncGetFromIni(List2.Text, "mwnd", App.Path & "\songdata.ini") & vbCrLf & "�����������" & ini.mfncGetFromIni(List2.Text, "mwlj", App.Path & "\songdata.ini") & vbCrLf & "�쾮��" & ini.mfncGetFromIni(List2.Text, "mwtj", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "mwcx", App.Path & "\songdata.ini") & vbCrLf & "���" & ini.mfncGetFromIni(List2.Text, "mwgc", App.Path & "\songdata.ini")

    sc = ini.mfncGetFromIni(List2.Text, "sc", App.Path & "\songdata.ini")
    If sc = "" Then
        Label1.Caption = "�������ƣ�" & List2.Text & vbCrLf & "BPM��" & ini.mfncGetFromIni(List2.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf
    Else
        Label1.Caption = "�������ƣ�" & List2.Text & vbCrLf & "BPM��" & ini.mfncGetFromIni(List2.Text, "bpm", App.Path & "\songdata.ini") & vbCrLf & "�׳ƣ�" & ini.mfncGetFromIni(List2.Text, "sc", App.Path & "\songdata.ini")
    End If
End Sub


