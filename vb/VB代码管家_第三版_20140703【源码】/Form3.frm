VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ת��"
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10335
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10335
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame2 
      Caption         =   "ת����"
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   10095
      Begin VB.TextBox Text2 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ת��ǰ"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.TextBox Text1 
         Height          =   3015
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Menu gjx_bmzh 
      Caption         =   "�������"
      Begin VB.Menu gjx_bmzh_UTF_8���� 
         Caption         =   "UTF-8����"
      End
      Begin VB.Menu gjx_bmzh_UTF_8���� 
         Caption         =   "UTF-8����"
      End
      Begin VB.Menu gjx_bmzh_GBK���� 
         Caption         =   "GBK����"
      End
      Begin VB.Menu gjx_bmzh_GBK���� 
         Caption         =   "GBK����"
      End
      Begin VB.Menu gjx_bmzh_Unicode���� 
         Caption         =   "Unicode����"
      End
      Begin VB.Menu gjx_bmzh_Unicode���� 
         Caption         =   "Unicode����"
      End
      Begin VB.Menu gjx_bmzh_qx 
         Caption         =   "ȡ��"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub gjx_bmzh_GBK����_Click()
    Text2.Text = URLEncode(Text1.Text)
End Sub

Private Sub gjx_bmzh_GBK����_Click()
    Text2.Text = URLDecode(Text1.Text)
End Sub

Private Sub gjx_bmzh_Unicode����_Click()
    Text2.Text = ToUnicode(Text1.Text)
End Sub

Private Sub gjx_bmzh_Unicode����_Click()
    Text2.Text = UnUnicode(Text1.Text)
End Sub

Private Sub gjx_bmzh_UTF_8����_Click()
    Text2.Text = UTF8_URLEncoding(Text1.Text)
End Sub

Private Sub gjx_bmzh_UTF_8����_Click()
    Text2.Text = UTF8_UrlDecode(Text1.Text)
End Sub
