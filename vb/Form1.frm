VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Բ���"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4980
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
r = InputBox("������Բ�İ뾶��", "����Բ�뾶")
pi = 3.14
s = pi * r * r
Print "Բ�İ뾶Ϊ��"; r; "����"
Print "Բ�����Ϊ��"; s; "ƽ������"
If r > 2000 Then
    Print "������ʾ��Χ"
  Else
    Circle (2500, 2500), r
    End If
End Sub
