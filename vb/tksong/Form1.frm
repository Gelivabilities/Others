VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   4575
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   5640
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
   Begin VB.ComboBox Combo5 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "街机" Then
Combo5.Visible = True
Combo4.Visible = True
Combo2.Width = 1335
Combo2.Left = 3120
Else
If Combo1.Text = "PSP" Then
Combo4.Visible = False
Combo5.Visible = False
Combo2.Width = 3255
Combo2.Left = 1200
Else
If Combo1.Text = "iOS" Then
Combo4.Visible = False
Combo5.Visible = False
Combo2.Width = 3255
Combo2.Left = 1200
Else
If Combo1.Text = "3ds1" Then
Combo5.Visible = False
Combo4.Visible = False
Combo2.Width = 3255
Combo2.Left = 1200
Else
If Combo1.Text = "3ds2" Then
Combo4.Visible = False
Combo5.Visible = False
Combo2.Width = 3255
Combo2.Left = 1200
End If
End If
End If
End If
End If
End Sub

Private Sub Combo5_Click()
If Combo5.Text = "旧框体" Then
Combo4.Clear
n = 1
Do While n <= 14
Combo4.AddItem "AC" & n, n - 1
n = n + 1
Loop
Combo4.ListIndex = 0
a = Combo1.ListIndex * 1000
b = Combo5.ListIndex * 400
c = Combo4.ListIndex * 10
d = Combo2.ListIndex
x = a + b + c + d
listthesongs (x)
Else
If Combo5.Text = "新框体" Then
Combo4.Clear
Combo4.AddItem "モモイロ"
Combo4.AddItem "ソライロ"
Combo4.AddItem "KATSU-DON"
Combo4.AddItem "初代"
Combo4.AddItem "モモイロ（アジア版）"
Combo4.ListIndex = 0
a = Combo1.ListIndex * 1000
b = Combo5.ListIndex * 400
c = Combo4.ListIndex * 10
d = Combo2.ListIndex
x = a + b + c + d
listthesongs (x)
End If
End If

End Sub

Private Sub Form_Load()
Combo1.AddItem "街机", 0
Combo1.AddItem "PSP", 1
Combo1.AddItem "iOS", 2
Combo1.AddItem "3ds1", 3
Combo1.AddItem "3ds2", 4
Combo2.AddItem "Namco原创曲", 0
Combo2.AddItem "JPOP", 1
Combo2.AddItem "古典", 2
Combo2.AddItem "游戏", 3
Combo2.AddItem "动漫", 4
Combo2.AddItem "儿童", 5
Combo2.AddItem "V家", 6
Combo5.AddItem "旧框体", 0
Combo5.AddItem "新框体", 1
Combo4.AddItem "モモイロ", 0
Combo4.AddItem "ソライロ", 1
Combo4.AddItem "KATSU-DON", 2
Combo4.AddItem "初代", 3
Combo4.AddItem "モモイロ（アジア版）", 4
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo4.ListIndex = 0
Combo5.ListIndex = 1
a = Combo1.ListIndex * 1000
b = Combo5.ListIndex * 400
c = Combo4.ListIndex * 10
d = Combo2.ListIndex
x = a + b + c + d
listthesongs (x)
End Sub

Public Function listthesongs(x) As Integer
List1.Clear
Select Case x
    Case 400
    i = 1
    Do While i <= 133
        List1.AddItem ini.mfncGetFromIni(400, "song" & i, App.Path & "\songlist.ini")
        i = i + 1
    Loop
    Case 3001
    i = 1
    Do While i <= 66
        List1.AddItem ini.mfncGetFromIni(3001, "song" & i, App.Path & "\songlist.ini")
        i = i + 1
    Loop
 End Select
End Function
