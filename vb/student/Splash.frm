VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3420
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraEdge 
      Height          =   3330
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "����������Visual Basic 6.0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2640
         TabIndex        =   6
         Top             =   1560
         Width           =   3450
      End
      Begin VB.Image imgLogo 
         Height          =   825
         Left            =   360
         Picture         =   "Splash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   735
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "��Ȩ���У�Υ�߱ؾ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   480
         TabIndex        =   2
         Top             =   2640
         Width           =   2100
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "1.0.0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4680
         TabIndex        =   3
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "���ݻ�����Access"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2640
         TabIndex        =   4
         Top             =   2040
         Width           =   2100
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ȩ���� �κθ���ϵͳ��������������"
         Height          =   180
         Index           =   5
         Left            =   480
         TabIndex        =   1
         Top             =   3000
         Width           =   3330
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "ѧ����Ϣ����ϵͳ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   4320
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ô������������ã�һΪϵͳ����ʱ�Ĵ��壬��Ϊϵͳ����ʱ�ġ�����...�����壬��mbAbout��Ϊ��ʶ
'��mbAboutΪtrue, ���ʾΪϵͳ����ʱ�Ĵ���
'��mbAboutΪfalse�����ʾΪϵͳ����ʱ�ġ�����...������
Public mbAbout As Boolean

Sub UnloadForm()
    Unload Me
    ''�����ǰΪϵͳ����ʱ����ʾ���壬�����˳�������֮����Ҫ���ص�¼����
    If Not mbAbout Then frmLogin.Show
End Sub
'���¸����룬��ʾ�������������ϵ��κβ��֣����߰�����һ�������������UnloadForm�ӳ���
Private Sub Form_Click()
    UnloadForm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    UnloadForm
End Sub

Private Sub fraEdge_Click()
   UnloadForm
End Sub

Private Sub imgLogo_Click()
    UnloadForm
End Sub

Private Sub lblInfo_Click(Index As Integer)
    UnloadForm
End Sub
