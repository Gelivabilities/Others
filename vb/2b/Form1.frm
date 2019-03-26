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
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "测试"
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
      Caption         =   "QQ最后一位"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "QQ第五位"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "QQ第一位"
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
a = "在学校"
Else
If Text1.Text = "2" Then
a = "在女厕所上"
Else
If Text1.Text = "3" Then
a = "在秦朝"
Else
If Text1.Text = "4" Then
a = "在赌博场"
Else
If Text1.Text = "5" Then
a = "在监狱"
Else
If Text1.Text = "6" Then
a = "在天安门广场"
Else
If Text1.Text = "7" Then
a = "在李刚家"
Else
If Text1.Text = "8" Then
a = "在高速公路上"
Else
If Text1.Text = "9" Then
a = "在网吧"
End If
End If
End If
End If
End If
End If
End If
End If

If Text2.Text = "1" Then
b = "挖地道"
Else
If Text2.Text = "2" Then
b = "跳楼"
Else
If Text2.Text = "3" Then
b = "梦游"
Else
If Text2.Text = "4" Then
b = "洗澡"
Else
If Text2.Text = "5" Then
b = "发呆"
Else
If Text2.Text = "6" Then
b = "看海绵宝宝"
Else
If Text2.Text = "7" Then
b = "剪指甲"
Else
If Text2.Text = "8" Then
b = "上网"
Else
If Text2.Text = "9" Then
b = "玩泥巴"
Else
If Text2.Text = "0" Then
b = "自残"
End If
End If
End If
End If
End If
End If
End If
End If

If Text3.Text = "1" Then
c = "被判死刑了"
Else
If Text3.Text = "2" Then
c = "被上帝带走了"
Else
If Text3.Text = "3" Then
c = "被钱砸了"
Else
If Text3.Text = "4" Then
c = "被警察击毙了"
Else
If Text3.Text = "5" Then
c = "被虐死了"
Else
If Text3.Text = "6" Then
c = "被吓得想尿尿"
Else
If Text3.Text = "7" Then
c = "被自行车撞飞了"
Else
If Text3.Text = "8" Then
c = "被人妖非礼了"
Else
If Text3.Text = "9" Then
c = "被洪水冲走了"
Else
If Text3.Text = "0" Then
c = "被表扬了"
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
Label4.Caption = "请输入正确数字"
End If
Dim u As String
u = Text3.Text
If IsNumeric(u) Then
g = 999
Else
Label4.Caption = "请输入正确数字"
End If
Dim t As String
t = Text2.Text
If IsNumeric(t) Then
g = 19
Else
Label4.Caption = "请输入正确数字"
End If
If Text1.Text = "0" Then
Label4.Caption = "请输入正确数字"
End If
End If
End Sub
