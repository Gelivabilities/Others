VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CPUռ���ʼ��"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2745
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4560
      Top             =   840
   End
   Begin VB.Label lblCPU 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'������****************************������**********************������
'��ŵVB԰-�ղ�����
'��վ��רע��VB��VBNET��̵�Դ������վ
'�������ڣ�2008-4-2 16:47:08
'��    վ��http://www.vbget.com/          (��ŵVB԰)
'��    վ��http://www.vbget.com/daohan/   (VB�����ַ����)
'E-Mail  ��vbget@yahoo.cn
'QQ      ��158676144
'Դ�����ߣ��������VB��ҵԴ����Ҫ������棬��վ����VIP�շ�����Ƶ���ɹ��㷢��!
'         ����Ȩ����;�ļ�;ɾ��;����ʱ�鿴������(������)����������ȫ��������
'         ��վ����˫��Э�̵�һ����������ڴ������ʻ��У���ֻ�踺������ã�
'         ��վֻ��Ϊһ��ƽ̨�ṩ����VBԴ����Ѷ��Դ�����أ�
'��ע����<վ������֮����ע��>����Զ���ӣ���ŵVB԰�д�������أ�
'������****************************������**********************������

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private CPU As clsCPUUsage

Private Sub Form_Load()

    Set CPU = New clsCPUUsage

End Sub



Private Sub Timer1_Timer()

    lblCPU.Caption = "CPUռ����: " & CPU.Usage & "%"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set CPU = Nothing

End Sub
