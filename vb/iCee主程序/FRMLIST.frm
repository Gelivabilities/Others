VERSION 5.00
Begin VB.Form FRMLIST 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "�����б�"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin ICEE.IList ILIST 
      Height          =   4320
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ItemHeight      =   18
   End
   Begin ICEE.ICEE_KEY ICW 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
   End
End
Attribute VB_Name = "FRMLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.BackColor = COLOR_NOR
ILIST.SETCOLOR COLOR_NOR, COLOR_HIGH
ICW.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
icm.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call REZOR
End Sub
Sub REZOR()
ICW.L_M_R = 1
If LOLIPOP = 3 Then
ICW.SETTXT "˳�򲥷�"
ElseIf LOLIPOP = 1 Then
ICW.SETTXT "����ѭ��"
ElseIf LOLIPOP = 2 Then
ICW.SETTXT "�б�ѭ��"
ElseIf LOLIPOP = 0 Then
ICW.SETTXT "�������"
End If
End Sub
Sub RELIST()
If frmma.PLIST.ListCount = 0 Then Exit Sub
ILIST.Clear
Dim I As Integer
For I = 0 To frmma.PLIST.ListCount - 1
Me.ILIST.AddItem frmma.PLIST.Title(I)
Next
End Sub
Private Sub Form_Load()
Call RELIST
Call oMagneticWnd.AddWindow(Me.hwnd, FRMTASK.hwnd) '���Դ���
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
IS_MINI_LIST = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call oMagneticWnd.RemoveWindow(Me.hwnd) '�رմ��Դ���
IS_MINI_LIST = False
End Sub
Private Sub ICW_Click()
Me.PopupMenu Frmm.˳��, , ICW.Left, ICW.Top + ICW.Height
End Sub

Private Sub ILIST_DBClick()
If ILIST.ListCount = 0 Then Exit Sub
frmma.Wm.URL = frmma.PLIST.URL(ILIST.ListIndex)
frmma.Wm.Controls.Play
End Sub
