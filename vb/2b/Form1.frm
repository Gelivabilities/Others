VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BY:Diboro"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   3885
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      MaxLength       =   1
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      MaxLength       =   1
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "QQ���һλ"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "QQ����λ"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "QQ��һλ"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "1" Then
a = "��ѧУ"
Else
If Text1.Text = "2" Then
a = "��Ů������"
Else
If Text1.Text = "3" Then
a = "���س�"
Else
If Text1.Text = "4" Then
a = "�ڶĲ���"
Else
If Text1.Text = "5" Then
a = "�ڼ���"
Else
If Text1.Text = "6" Then
a = "���찲�Ź㳡"
Else
If Text1.Text = "7" Then
a = "����ռ�"
Else
If Text1.Text = "8" Then
a = "�ڸ��ٹ�·��"
Else
If Text1.Text = "9" Then
a = "������"
End If
End If
End If
End If
End If
End If
End If
End If

If Text2.Text = "1" Then
b = "�ڵص�"
Else
If Text2.Text = "2" Then
b = "��¥"
Else
If Text2.Text = "3" Then
b = "����"
Else
If Text2.Text = "4" Then
b = "ϴ��"
Else
If Text2.Text = "5" Then
b = "����"
Else
If Text2.Text = "6" Then
b = "�����౦��"
Else
If Text2.Text = "7" Then
b = "��ָ��"
Else
If Text2.Text = "8" Then
b = "����"
Else
If Text2.Text = "9" Then
b = "�����"
Else
If Text2.Text = "0" Then
b = "�Բ�"
End If
End If
End If
End If
End If
End If
End If
End If

If Text3.Text = "1" Then
c = "����������"
Else
If Text3.Text = "2" Then
c = "���ϵ۴�����"
Else
If Text3.Text = "3" Then
c = "��Ǯ����"
Else
If Text3.Text = "4" Then
c = "�����������"
Else
If Text3.Text = "5" Then
c = "��Ű����"
Else
If Text3.Text = "6" Then
c = "���ŵ�������"
Else
If Text3.Text = "7" Then
c = "�����г�ײ����"
Else
If Text3.Text = "8" Then
c = "������������"
Else
If Text3.Text = "9" Then
c = "����ˮ������"
Else
If Text3.Text = "0" Then
c = "��������"
Else
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

Label4.Caption = a & b & c
Dim s As String
s = Text1.Text
If IsNumeric(s) Then
g = 1999
Else
Label4.Caption = "��������ȷ����"
End If
Dim u As String
u = Text3.Text
If IsNumeric(u) Then
g = 999
Else
Label4.Caption = "��������ȷ����"
End If
Dim t As String
t = Text2.Text
If IsNumeric(t) Then
g = 19
Else
Label4.Caption = "��������ȷ����"
End If
If Text1.Text = "0" Then
Label4.Caption = "��������ȷ����"
End If
End If
End Sub
