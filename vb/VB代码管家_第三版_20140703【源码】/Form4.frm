VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Postģ���ύ����"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   10305
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton ��ʼ 
      Caption         =   "��ʼ"
      Height          =   280
      Left            =   9240
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton ���Cookie 
      Caption         =   "���Cookie"
      Height          =   280
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton ʱ��� 
      Caption         =   "ʱ���"
      Height          =   280
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "���ݰ�"
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   10095
      Begin VB.TextBox ���ݰ� 
         Height          =   5895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Э��ͷ"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   10095
      Begin VB.TextBox Э��ͷ 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame Frame5 
         Caption         =   "�ύ��ַ"
         Height          =   650
         Left            =   120
         TabIndex        =   8
         Top             =   200
         Width           =   9855
         Begin VB.ComboBox �ύ��ʽ 
            Height          =   300
            ItemData        =   "Form4.frx":030A
            Left            =   8760
            List            =   "Form4.frx":0314
            TabIndex        =   11
            Text            =   "Get"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox ���뷽ʽ 
            Height          =   300
            ItemData        =   "Form4.frx":0323
            Left            =   7680
            List            =   "Form4.frx":0330
            TabIndex        =   10
            Text            =   "UTF-8"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox POST_GET_URL 
            Height          =   300
            Left            =   120
            TabIndex        =   9
            Text            =   "http://www.meilishuo.com/"
            Top             =   240
            Width           =   7455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "POST����"
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   9855
         Begin VB.TextBox POST_GET_DATE 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   9615
         End
      End
   End
   Begin VB.Menu sjc 
      Caption         =   "ʱ���"
      Visible         =   0   'False
      Begin VB.Menu ����13λʱ��� 
         Caption         =   "����13λʱ���"
      End
      Begin VB.Menu ����10λʱ��� 
         Caption         =   "����10λʱ���"
      End
      Begin VB.Menu ȡ�� 
         Caption         =   "ȡ��"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ��ʼ_Click()
    POST_GET_URL.Text = Replace(POST_GET_URL.Text, "��ʱ���10λ��", ʱ���B())
    POST_GET_URL.Text = Replace(POST_GET_URL.Text, "��ʱ���13λ��", ʱ���A())
    POST_GET_DATE.Text = Replace(POST_GET_DATE.Text, "��ʱ���10λ��", ʱ���B())
    POST_GET_DATE.Text = Replace(POST_GET_DATE.Text, "��ʱ���13λ��", ʱ���A())
    If �ύ��ʽ.Text = "GET" Then
        ���ݰ�.Text = GetData(POST_GET_URL.Text, ���뷽ʽ.Text)
    Else
        ���ݰ�.Text = PostData(POST_GET_URL.Text, POST_GET_DATE.Text, ���뷽ʽ.Text)
    End If
End Sub

Private Sub ���Cookie_Click()
    Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351", vbMaximizedFocus
End Sub

Private Sub ����10λʱ���_Click()
    POST_GET_DATE.SetFocus
    SendKeys "��ʱ���10λ��"
End Sub

Private Sub ����13λʱ���_Click()
    POST_GET_DATE.SetFocus
    SendKeys "��ʱ���13λ��"
End Sub

Private Sub ʱ���_Click()
    Me.PopupMenu sjc, , ʱ���.Left, ʱ���.Top + ʱ���.Height '��ʾʱ����˵�
End Sub
