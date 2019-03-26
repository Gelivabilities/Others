VERSION 5.00
Begin VB.Form FRMHIS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "程序日志"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   Icon            =   "FRMHIS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   1200
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   8385
      Picture         =   "FRMHIS.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   8385
      Picture         =   "FRMHIS.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   4
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   8385
      Picture         =   "FRMHIS.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   15
      Width           =   750
   End
   Begin VB.Timer TimAutoSave 
      Interval        =   5000
      Left            =   8520
      Top             =   720
   End
   Begin VB.TextBox TXTNODE 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6600
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   900
      Width           =   8550
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   1200
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   1200
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   1200
      Index           =   3
      Left            =   6360
      TabIndex        =   9
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
   End
   Begin ICEE.ICEE_WIN8 IW 
      Height          =   1200
      Index           =   4
      Left            =   7680
      TabIndex        =   10
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
   End
   Begin VB.TextBox TXTSYS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6600
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   8550
   End
   Begin VB.TextBox TXTMSG 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6600
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   8550
   End
End
Attribute VB_Name = "FRMHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

Private Sub Form_Activate()
'Call FrmTrans(Me)
Me.BackColor = COLOR_NOR
Call PaintPng(App.Path & "\SKIN\HS_T.PNG", Me.hdc, 8, 8)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), Frmm.PTCO.POINT(1, 1), B
Dim i As Integer
For i = 0 To IW.Count - 1
IW(i).HASLINE = False
IW(i).SETCOLOR COLOR_NOR, COLOR_HIGH
IW(i).SETTXTCOLOR vbWhite, vbWhite
Next

IW(0).SETPNG App.Path & "\SKIN\SMSG.PNG", (IW(0).Width - 64) / 2, (IW(0).Height - 64) / 2
IW(2).SETPNG App.Path & "\SKIN\FMSG.PNG", (IW(2).Width - 64) / 2, (IW(2).Height - 64) / 2
IW(1).SETPNG App.Path & "\SKIN\CZ.PNG", (IW(1).Width - 64) / 2, (IW(1).Height - 64) / 2

TXTSYS.BackColor = Me.BackColor
TXTMSG.BackColor = Me.BackColor
TXTNODE.BackColor = Me.BackColor
End Sub

Private Sub Form_Load()


IW(0).SETTIP "服务器推送"
IW(1).SETTIP "操作记录"
IW(2).SETTIP "好友消息"

IW(3).HASTIP = False
IW(4).HASTIP = False
IW(3).SETTXT "导出"
IW(4).SETTXT "清空"

On Error GoTo ERR
Dim filea As String, FileB As String, FILEC As String, stg As String, lne As String
filea = App.Path & "\COFING\Action_Note.txt"
FileB = App.Path & "\COFING\MSG_Note.txt"
FILEC = App.Path & "\COFING\SYSTEM_Note.txt"
If PathFileExists(filea) = 1 Then
Open filea For Input As #1
Do Until EOF(1)
Line Input #1, lne
stg = stg & lne
If Not EOF(1) Then stg = stg & vbNewLine
DoEvents
Loop
Close #1
TXTNODE.Text = stg
End If
stg = ""
Open FileB For Input As #1
If PathFileExists(FileB) = 1 Then
Do Until EOF(1)
Line Input #1, lne
stg = stg & lne
If Not EOF(1) Then stg = stg & vbNewLine
DoEvents
Loop
Close #1
TXTMSG.Text = stg
End If
stg = ""
If PathFileExists(FILEC) = 1 Then
Open FILEC For Input As #1
Do Until EOF(1)
Line Input #1, lne
stg = stg & lne
If Not EOF(1) Then stg = stg & vbNewLine
DoEvents
Loop
Close #1
TXTSYS.Text = stg
End If
GoTo ERR
ERR:
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub
Private Sub IW_Click(Index As Integer)
Dim linecount As Integer
Dim filename As String
Select Case Index
Case 0
TXTNODE.Visible = True '操作记录
TXTSYS.Visible = False '系统消息
TXTMSG.Visible = False '好友信息
linecount = SendMessage(TXTNODE.hwnd, EM_GETLINECOUNT, 0&, 0&)

Case 1
TXTNODE.Visible = False
TXTSYS.Visible = True
TXTMSG.Visible = False
linecount = SendMessage(TXTSYS.hwnd, EM_GETLINECOUNT, 0&, 0&)

Case 2
TXTNODE.Visible = False
TXTSYS.Visible = False
TXTMSG.Visible = True
linecount = SendMessage(TXTMSG.hwnd, EM_GETLINECOUNT, 0&, 0&)
Case 3
filename = ShowSave(Me.hwnd, "程序日志" & Chr(0) & "*.his", "保存记录")
If filename = "" Then Exit Sub
If TXTNODE.Visible = True Then
If TXTNODE.Text = "" Then Call SHOWWRONG("没有需要记录的事情,无法导出", 0): Exit Sub
Open filename For Output As #1
Print #1, LOF(1) + 1, TXTNODE.Text
Close #1
ElseIf TXTSYS.Visible = True Then
If TXTSYS.Text = "" Then Call SHOWWRONG("没有需要记录的事情,无法导出", 0): Exit Sub
Open filename For Output As #1
Print #1, LOF(1) + 1, TXTSYS.Text
Close #1
ElseIf TXTMSG.Visible = True Then
If TXTMSG.Text = "" Then Call SHOWWRONG("没有需要记录的事情,无法导出", 0): Exit Sub
Open filename For Output As #1
Print #1, LOF(1) + 1, TXTMSG.Text
Close #1
End If

Case 4
If TXTNODE.Visible = True Then
TXTNODE.Text = ""
ElseIf TXTSYS.Visible = True Then
TXTSYS.Text = ""
ElseIf TXTMSG.Visible = True Then
TXTMSG.Text = ""
End If
End Select
End Sub

Private Sub TimAutoSave_Timer()
On Error Resume Next
'BINARY 是追加写入,OUTPUT 是重新写入
Open App.Path & "\COFING\Msg_Note.txt" For Output As #1
Print #1, LOF(1) + 1, vbCrLf & Replace(TXTMSG.Text, "1" & vbCrLf, "") & vbCrLf
Close #1
Open App.Path & "\COFING\Action_Note.txt" For Output As #1
Print #1, LOF(1) + 1, vbCrLf & Replace(TXTNODE.Text, "1" & vbCrLf, "") & vbCrLf
Close #1
Open App.Path & "\COFING\System_Note.txt" For Output As #1
Print #1, LOF(1) + 1, vbCrLf & Replace(TXTSYS.Text, "1" & vbCrLf, "") & vbCrLf
Close #1
End Sub

Private Sub TXTMSG_Change()
TXTMSG.Text = Replace(Trim(TXTMSG.Text), "1", "")
TXTMSG.Text = Replace(Trim(TXTMSG.Text), vbCrLf & vbCrLf & vbCrLf, "")
TXTMSG.SelStart = Len(TXTMSG)
End Sub

Private Sub TXTNODE_Change()
TXTNODE.Text = Replace(Trim(TXTNODE.Text), "1", "")
TXTNODE.Text = Replace(Trim(TXTNODE.Text), vbCrLf & vbCrLf & vbCrLf, "")
TXTNODE.SelStart = Len(TXTNODE)
End Sub

Private Sub TXTSYS_Change()
TXTSYS.Text = Replace(Trim(TXTSYS.Text), "1", "")
TXTSYS.Text = Replace(Trim(TXTSYS.Text), vbCrLf & vbCrLf & vbCrLf, "")
TXTSYS.SelStart = Len(TXTSYS)
End Sub

Private Sub x1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then
Me.Hide
End If
End Sub
