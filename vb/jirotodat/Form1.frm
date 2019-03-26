VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "add"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   6135
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "calculate"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      MaxLength       =   16
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "gc"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "cx"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "offset"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "bpm"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "yf"
      Height          =   375
      Left            =   120
      TabIndex        =   2
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
Dim head, yfNum, ending, yfString, dygyfNum, dygyfData, cx, gc, offset As String



yfNum = Int(Len(Replace(Text1.Text, " ", "")))

If yfNum < 16 Then
yfNum = "0" & Hex(yfNum)
Else
yfNum = Hex(yfNum)
End If

yfString = Text1.Text

For i = 1 To 16
If Mid(yfString, i, 1) <> " " Then
dygyf = Mid(yfString, i, 1)
Exit For
End If
Next

If dygyf = "o" Then
dygyfNum = "02"
End If

If dygyf = "x" Then
dygyfNum = "04"
End If

If dygyf = "O" Then
dygyfNum = "07"
End If

If dygyf = "X" Then
dygyfNum = "08"
End If

'第一个
dygyfData = "00 39 00 00 00 80 3F " & dygyfNum & " 00 00 00 00 00 00 00 00 "


'开头
head = "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF F4 FC 12 00 " & yfNum

bpm = Text3.Text

offset = Text4.Text

cx = Text5.Text

gc = Text6.Text

'最后面部分
ending = "FC 12 00 88 95 43 00 " & Hex(Text5.Text - 256) & " 01 64 " & "00 00 00 00 00 00 00 39 00 00 00 80 3F 00 00 39 00 00 00 80 3F 0F 00 " & Hex(bpm - 128) & " 43 3D 2E 81 46 00 01 00 00"

Text2.Text = head & " " & dygyfData & ending

End Sub

Private Sub Command2_Click()
Text7.Text = Text7.Text & Text2.Text
End Sub
