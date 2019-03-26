VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "计算圆面积"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   4815
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
r = InputBox("请输入圆的半径：", "输入圆半径")
pi = 3.14
s = pi * r * r
Print "圆的半径为："; r; "厘米"
Print "圆的面积为："; s; "平方厘米"
If r > 2000 Then
Print "超出显示范围"
 Else
 Circle (2500, 2500), r
End If
End Sub
