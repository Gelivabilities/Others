VERSION 5.00
Begin VB.Form FrmHelp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "������֧��-����ICEE��Ȩ���Ľ���"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9240
      Picture         =   "FrmHelp.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   16
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9240
      Picture         =   "FrmHelp.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   15
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   9240
      Picture         =   "FrmHelp.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   14
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   120
      ScaleHeight     =   561
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox PKB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00241D0A&
         BorderStyle     =   0  'None
         Height          =   900
         Left            =   360
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   13
         Top             =   120
         Width           =   900
      End
      Begin ICEE.ICEE_TEXT HOT 
         Height          =   6495
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11456
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӱ��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   7680
         TabIndex        =   26
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   6960
         TabIndex        =   25
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ؿ�ζ��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   6000
         TabIndex        =   24
         Top             =   8040
         Width           =   720
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ں���"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   5280
         TabIndex        =   23
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ϵĽ�Ѿ��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   4080
         TabIndex        =   22
         Top             =   8040
         Width           =   1080
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ֲ���"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   3360
         TabIndex        =   21
         Top             =   8040
         Width           =   540
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PS��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   2760
         TabIndex        =   20
         Top             =   8040
         Width           =   360
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   19
         Top             =   8040
         Width           =   360
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "С���ֻ�����"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   18
         Top             =   8040
         Width           =   1080
      End
      Begin VB.Label LF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ٷ�����"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   8040
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   136
         X2              =   342
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Question and Answers"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   0
         Left            =   2040
         TabIndex        =   12
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label LA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   10
         Top             =   120
         Width           =   840
      End
      Begin VB.Shape SB 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   0
         Top             =   7800
         Width           =   9735
      End
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   975
      Index           =   6
      Left            =   3480
      TabIndex        =   6
      Top             =   6600
      Width           =   2895
      _ExtentX        =   5953
      _ExtentY        =   1085
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1095
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   7680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1931
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3413
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   3615
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6376
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1335
      Index           =   2
      Left            =   6480
      TabIndex        =   2
      Top             =   5160
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1935
      Index           =   3
      Left            =   6480
      TabIndex        =   3
      Top             =   3120
      Width           =   3375
      _ExtentX        =   4471
      _ExtentY        =   1296
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   2175
      Index           =   5
      Left            =   6480
      TabIndex        =   5
      Top             =   6600
      Width           =   3375
      _ExtentX        =   6800
      _ExtentY        =   1296
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1335
      Index           =   7
      Left            =   3480
      TabIndex        =   7
      Top             =   5160
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2355
   End
   Begin ICEE.ICEE_WIN8 IHELP 
      Height          =   1935
      Index           =   8
      Left            =   3480
      TabIndex        =   8
      Top             =   3120
      Width           =   2895
      _ExtentX        =   11245
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IS_MV As Boolean

Private Sub Form_Activate()
Me.BackColor = COLOR_NOR

Dim PBOX As Control
For Each PBOX In Me.Controls
If TypeOf PBOX Is PictureBox Then
PBOX.Cls
PBOX.BackColor = Me.BackColor
End If
Next
Call PaintPng(App.Path & "\SKIN\H_T.PNG", Me.hdc, 8, 8)
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
HOT.SETBACKCOLOR Me.BackColor
HOT.SETFORECOLOR vbWhite
End Sub

Private Sub Form_Load()
For i = 0 To IHELP.Count - 1
IHELP(i).HASLINE = False
IHELP(i).HASTIP = False
Next
IS_MV = False
IHELP(0).SETTXT "����������"
IHELP(1).SETTXT "Ϳѻ����"
IHELP(2).SETTXT "�ļ���������"
IHELP(3).SETTXT "�����๦������"
IHELP(4).SETTXT ""
IHELP(5).SETTXT "����������"
IHELP(6).SETTXT "�ļ�����������"
IHELP(7).SETTXT "UI��������"
IHELP(8).SETTXT "�������"
IHELP(1).SETCOLOR RGB(100, 28, 40), RGB(146, 19, 41)
IHELP(2).SETCOLOR RGB(170, 48, 63), RGB(203, 75, 75)
IHELP(3).SETCOLOR RGB(9, 43, 84), RGB(14, 83, 146)
IHELP(4).SETCOLOR &H25614B, &H2EBC7C
IHELP(5).SETCOLOR &H5B2989, &H563AB6
IHELP(6).SETCOLOR RGB(8, 70, 112), RGB(26, 109, 161)
IHELP(7).SETCOLOR RGB(67, 135, 148), RGB(77, 172, 190)
IHELP(0).SETCOLOR RGB(50, 28, 40), RGB(96, 19, 41)
IHELP(8).SETCOLOR vbBlack, COLOR_HIGH
If frmma.Left > Me.Width Then
Me.Move frmma.Left - Me.Width, frmma.Top
Else
Me.Move frmma.Left + frmma.Width, frmma.Top
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = True Then
IS_MV = False
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
End If

X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub
Private Sub IHELP_CLICK(Index As Integer)
Select Case Index
Case 0
HOT.SETTXT "Q:��α��沥���б�." & vbCrLf & _
"A:�����˳�����Զ������б�,��Ҳ�����ڲ˵�ѡ�񵼳��б�" & vbCrLf & _
"Q:��β��Ҹ�������." & vbCrLf & _
"A:�������ṩ�����߲��ҷ��湦��,�ɿṷ�ṩ,�������ڷ��浥�����Ͻǵ�������ť,�������ִ���������ʱ�ر���������,��Ҳ���Խ�������ק�����������ֶ�����" & vbCrLf & _
"Q:��β�������." & vbCrLf & _
"A:�����Ե����򿪰�ť�������б�˵���ѡ������ļ����ļ���,���ߴ����ִ��ڴ�ϲ���ĸ���,��ק�ļ�������ɼ��Ĳ���Ҳ���Բ���" & vbCrLf & _
"Q:���ִ����ּ�����ʱʧЧ." & vbCrLf & _
"A:����,��ǰICEEֻ��һ��ý��,���ֿ����԰ٶ���������,���ָ�������ʧЧ�����޷�����,��ʱ�˿ڱ�ռ�ó�����޷������б�,����������������" & vbCrLf & _
"Q:����˳�򾭳�����. " & vbCrLf & _
"A:��ǰ�汾����˳����ܻ���ڴ���,���߻����պ��Ż�����" & vbCrLf & _
"Q:��η�������." & vbCrLf & _
"A:�����õ�����ʱ,�����Ȳ���������������ѷ���,��ֻҪ�������б�˵�ѡ�� �������� ����" & vbCrLf & _
"Q:����Ӣ�ĸ���ʱ����㾭������." & vbCrLf & _
"A:����,ICEE�ĸ�ʿ��������ĸ��վ,Ӣ�ĸ�ʿ϶����Ǻ�ȫ��,��Ȼ,Ҳ��һЩ������������,������ѡ�� ɾ����� �����ֶ���������" & vbCrLf & _
"Q:�����б��ڵ��ļ����б���ʲô��." & vbCrLf & _
"A:���ݸ��˰���,�ļ����б���Է����û������ļ��������и��� " & vbCrLf & _
"Q:��θ��ĸ���." & vbCrLf & _
"A:���������ļ����ִ���ʱ,�����Ե�����������,����һ����������޸�" & vbCrLf & "Q:��������С����������������㲥����" & vbCrLf & _
"A:Ϊ�˷����û�,��������ļ����ڲ��ų�������С��������������������㲥����,�����û������и�" & vbCrLf & _
"Q:�����ļ��Ĳ���" & vbCrLf & _
"A:����,Ŀǰ�汾�����б��ڲ��Ҹ�������û��ʵ�ֵ�" & vbCrLf & _
"Q:����ɾ���ļ�." & vbCrLf & _
"A:��ν����ɾ�����ǴӴ�����ɾ��,ɾ�����ļ��ᱻ�Ƴ��б��������վ,�������ڲ����б�ѡ�� ����ɾ��,���߲��������Ͻǵ�����Ͱ����ɾ��" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Case 1
HOT.SETTXT "Q:Ϳѻ��ʲô." & vbCrLf & "A:�ܶ಻�˽�Ϳѻ�����ǻ���ΪͿѻ������Ϳ�һ�����ʵ��Ȼ. GRAFFITI ��һ���Ӿ��������������Ϳѻ���ݰ����ܶ�:��Ҫ�Ա���Ӣ������Ϊ���������3Dдʵ.����дʵ.���ֳ���дʵ .��ͨ����ȵ�.������������ɫ���˲���ǿ�ҵ��Ӿ�Ч���ĺ�����Ч��" & vbCrLf & _
"Q:��ô��Ϳѻ." & vbCrLf & "A:��Ϳѻ�ķ�ʽ���������˵��� Ϳѻ���� /���������ʹ�� Ϳѻ��ͼƬ/ͼ�������ļ�Ԥ��ѡ��˵� Ϳѻ��ͼƬ,���ɽ���Ϳѻ����" & vbCrLf & "Q:Ϳѻ����Ĺ���" & vbCrLf & _
"A:Ϳѻ�����Ϊ�������빤��������,����������Ȼ�ǻ滭��,ICEE�Ļ������𻯵�,������Ӳ��,��Ҳ������Ϳѻ�ı���" & vbCrLf & "Q:����Ϳѻ" & vbCrLf & "A:�������ṩ����Ϳѻ,��ͼƬ,��ӡͼƬ,����ͼƬ���ֹ���,��Ҳ����ͨ���˵�����ͿѻΪBMP����JPG��ʽ,JPGʧ���ʱȽϸ�,�Ƽ�ʹ��BMP" & vbCrLf & _
"Q:Ϳѻ�Ļ��ʵ���" & vbCrLf & "A:Ϳѻ�Ļ��ʿ��Ըı��ϸ,��1��20֮��ѡ��,Ӳ�ȿ�����10��100��ѡ��,�ʵ�����Ӳ�����ϸ��ʹ��Ʒ��ϸ��" & vbCrLf & _
"Q:�������Ϳѻ��ʲô����." & vbCrLf & "A:�õ���Ʒ��Ҫ�ḻ����ɫ,������Ԥ������6��ɫ�����·��Ĺ������Է������,Ҳ����ͨ��ɫ�������ɫ��ѡ��,�ḻ��ɫ������,��Ȼ��ҪͿĨ,��Ҫ���鷳,һ��һ����ɫ���ۻ���ʹ�������" & vbCrLf & _
"Q:��η���Ϳѻ." & vbCrLf & "����һ���û���ȻҪ�԰��԰�,����Ҫ��¼ICEE���ɷ���,��½�ɹ���,����������ѡ��Ҫ����ĺ���" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "<���ν���>"
Case 2
HOT.SETTXT "Q:��δ��ļ�����." & vbCrLf & "A:�ļ�������ָ���ļ����������������������Ӳ��,��Ҫ���ӻ�����,������ͨ�����������½ǵ����������ֲ������Ϸ������ؽ���" & vbCrLf & _
"Q:��������������." & vbCrLf & "A:���������������ᷢ�ֽ���ǳ���,��ֻ��Ҫ����������񼴿ɵ�������������,�����ַ,ȷ������,Ŀǰ֧�ִ󲿷ֵ������ļ�(Ѹ��,�쳵,����ĵ�ַ����)" & vbCrLf & _
"Q:�����ٶȹ�����ô��." & vbCrLf & "A:����,�����ٶ��������ٶ��й�,��������ռ��������Դ�Ľ���������,�������������������,Ҳ�п������ļ����������ػ�ά��." & vbCrLf & _
"Q:�����ļ���ַ�Ĳ���" & vbCrLf & "�Ҽ���������Ӳ˵����Զ�������и���,ɾ��,ֹͣ,�򿪱���λ�õȲ���" & vbCrLf & _
"Q:����ʧ��." & vbCrLf & "����,����������������ļ�ʧЧ����,Ҳ��������δ���ӵ�������"
Case 3
HOT.SETTXT "Q:ICEE�������ܶ�����Щ." & vbCrLf & "A:ICEE�����˶��ָ�������,������Ļ�����/��Ļ�Ŵ�/����/�����ǩ/ϵͳ��Դ����/������/���а����/�ļ������ȹ���" & vbCrLf & vbCrLf & _
 vbCrLf & "Q:��δ򿪷Ŵ�." & vbCrLf & "A:����,��Ļ�Ŵ󾵿���ͨ�� ���˵�-��Ļ�Ŵ� ����,������������ȫ����ʾΪ��Ļ����,�����ԶԷŴ�ı������е���" & vbCrLf & "Q:���ڽ�����һЩ����." & vbCrLf & _
"A:1.24�汾���������м�����2����ݰ�ť�������н�����ť����ǩ�İ�ť,�������Է����û��ҵ�,�����Ե����ɫ�İ�ť����,����������˿�ݼ�,F8Ҳ�ǿ��Խ���,�ڸ������ײ��Ĺ�����Ҳ�н����İ�ť." & vbCrLf & "Q:�����������ǩ." & vbCrLf & "A:��������-�����ǩ �����˵�=�½���ǩ ���������ɫ��ť" & vbCrLf & _
"Q:�����ǩ����������." & vbCrLf & "A:Ϊ�˲����ڴ�ռ��������,ICEE�������ǩ��������,��ֻ���Խ���10����ǩ,����,�����û�����������Ӧ���Ǻ��������" & vbCrLf & "Q:�رձ�ǩ�����ݻᱻ������." & vbCrLf & "A:����,��ǩ������������رյĻ������ǲ��ᱻ�����,��ǩ�Ľ��������ֹرշ�ʽ,[���]���޺ۼ��ر�,[X]���Ǳ������ݲ��ر�" & vbCrLf & _
"Q:ICEEΪʲôҪ��ϵͳ��Դ���м���." & vbCrLf & "A:����,ICEE����Դ���ӿ��Է����û��۲�ϵͳ�ı仯,1.24�Ľ����ֱ��,�����Բ鿴CPU/�ڴ�/USB/�����ڴ�ı仯" & vbCrLf & "Q:����������." & vbCrLf & "A:������-����������,ֱ�����뷽��ʽ�������㼴��" & vbCrLf & "Q:���а�ļ���." & vbCrLf & _
"A:���а���Ի��ͼ���ļ����ı��ļ�,ͼ���ļ����Ա���,Ϳѻ,������ӡ,�ı������Զ�����Ϊ�ı��ļ�,�û�����ͨ���ı����·��İ�ť�鿴�м�¼�����������ı�." & vbCrLf & "Q:���а����ʧЧ����ô��." & vbCrLf & "A:ͨ�� ʧЧ�˵����� ���ɻָ�ICEE�Լ��а�ļ���" & vbCrLf & _
"Q:��������ļ�." & vbCrLf & "A:����,����Ʒ�ݲ�֧���ļ�������,Ŀǰֻ������׺�ļ�.�û�����ͨ�����ֲ�����-�����б�-�������� ��������ģʽ,�����׺��������,����������ᱣ����Media�ļ�����,�´����к���Զ������������." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
Case 4
HOT.SETTXT "��û���дʲô"
Case 5
HOT.SETTXT "Q:���ע��ICEE���˺�." & vbCrLf & "A:ICEE�ķ�����Ŀǰ���Ǻ��ȶ�,�������ĵ�ַ��Ҫ�û��ֶ�����,Ҳ����ͨ�� ���� ��������,����������û�,����������ID������󽫵�½�����[�����û���¼]��ѡ��¼����,������������û������乴ѡ��,����ʧ��" & vbCrLf & "Q:�����Ӻ���." & vbCrLf & "A:��½�����������ѿ�����ID���س�����" & vbCrLf & "Q:ɾ�����Ѽ����κ��ѵĳ�������" & vbCrLf & _
"A:�Ҽ����������б�,�����˵�ѡ�����ѡ���." & vbCrLf & "Q:����ѿ��ٵ�����" & vbCrLf & "A:˫������ID���������ı���,�������ݼ��ɷ���,�Է������յ���������" & vbCrLf & "��ʱ��������������������ʲô." & vbCrLf & "��������,����˼��,���Դ��ı�����ʽ����,��ʱ����ʱ�������и���Ĺ���ѡ��,���緢���ļ�,���ͱ���,�ٱ�,Զ��Э���ȹ���" & vbCrLf & _
"Q:�����ļ���һЩ����" & "A:����,��ν��Ƚϳ�,���߽�����[�ļ�����]�������ϸ���" & vbCrLf & "Q:��θ�������." & vbCrLf & "�ں����б�˵�ѡ���޸����뼴�ɽ����ӽ���,���볤�Ȳ�����16λ,�޸ĳɹ���������Ϣ��֪ͨ." & vbCrLf & _
""
Case 6
HOT.SETTXT "����׫����"
Case 7
HOT.SETTXT "����׫����"
Case 8
HOT.SETTXT "����׫����"
End Select
PO.Visible = True
LA(2).Caption = IHELP(Index).MYTIT
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub LF_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Dim WILL_URL As String
Select Case Index
Case 0
WILL_URL = "http://tieba.baidu.com/f.kw=icee&fr=index"
Case 1
WILL_URL = "http://tieba.baidu.com/f.kw=%E5%B0%8F%E7%B1%B3&fr=index&fp=0&ie=utf-8"
Case 2
WILL_URL = "http://tieba.baidu.com/f.kw=vb&fr=index"
Case 3
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=PS"
Case 4
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E6%81%90%E6%80%96"
Case 5
WILL_URL = "http://tieba.baidu.com/f.kw=%BC%E7%C9%CF%B5%C4%BD%C5%D1%BE&fr=index"
Case 6
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E5%86%85%E6%B6%B5"
Case 7
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E9%87%8D%E5%8F%A3%E5%91%B3"
Case 8
WILL_URL = "http://tieba.baidu.com/f.ie=utf-8&kw=%E7%BE%8E%E5%89%A7"
Case 9
WILL_URL = "http://tieba.baidu.com/f.kw=%B5%E7%D3%B0&fr=ala0"
End Select
ShellExecute 0&, vbNullString, WILL_URL, vbNullString, vbNullString, 0 '����ie
End Sub

Private Sub LF_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If LF(Index).FOREColor <> &H30F1F1 Then LF(Index).FOREColor = &H30F1F1
End Sub
Private Sub PKB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = False Then
IS_MV = True
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_H.PNG", PKB.hdc, 0, 0)
End If
End Sub

Private Sub PKB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
PO.Visible = False
End Sub

Private Sub PO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IS_MV = True Then
IS_MV = False
PKB.Cls
Call PaintPng(App.Path & "\SKIN\BK_N.PNG", PKB.hdc, 0, 0)
End If
X1.Visible = True
X2.Visible = False
X3.Visible = False
Dim i As Integer
For i = 0 To LF.Count - 1
If LF(i).FOREColor <> vbWhite Then LF(i).FOREColor = vbWhite
Next
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
