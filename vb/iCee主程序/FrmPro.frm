VERSION 5.00
Begin VB.Form FrmPro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "����������"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   ControlBox      =   0   'False
   Icon            =   "FrmPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox C2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4320
      Picture         =   "FrmPro.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4320
      Picture         =   "FrmPro.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox C1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   4320
      Picture         =   "FrmPro.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   15
      Width           =   750
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
   End
   Begin VB.Label ts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "    ��������Ҫ����������쳣!�Դ˸�����ɵĲ��������!"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      Top             =   915
      Width           =   3675
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

Private Sub c1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = False
C2.Visible = True
End Sub
Private Sub c2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
C2.Visible = False
C3.Visible = True
End If
End Sub
Private Sub c3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C3.Visible = False
C1.Visible = True
If C3.Visible = False Then
Unload Me
End If
End Sub


Private Sub Form_Load()
On Error Resume Next
MakeTransparent Me.hWnd, 250
Call SeekMe(Me)
Me.BackColor = COLOR_NOR
ts.Top = (Me.ScaleHeight - ts.Height) / 2
sndPlaySound App.Path + "\Sound\popo.wav", 1
Me.DrawWidth = 1
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H201400, B
ICM(0).SETTXT "����"
ICM(1).SETTXT "��ֹICEE"
ICM(2).SETTXT "����"
ICM(0).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICM(1).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICM(2).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite

Call PaintPng(App.Path & "\SKIN\MSG_ASK.PNG", Me.hdc, 8, 40)
Call PaintPng(App.Path & "\SKIN\W_T.PNG", Me.hdc, 8, 8)
RESL = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags) '�ö�

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
C1.Visible = True
C2.Visible = False
C3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call RemoveFromTray '�Ƴ�����ͼ��
End Sub

Private Sub ICM_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
SoftSAFE = EXCEPTION_CONTINUE_SEARCH      '����,������������ﴦ��,��ϵͳ�ұ�Ĵ��������ȥ����....�����Ȼ��.....
Unload Me
Case 1
End
SoftSAFE = EXCEPTION_CONTINUE_EXECUTION
Unload Me '����ִ��,ִ�е�ַ��pContextRecord��ָ��
Case 2
StCT.regEIP = StCT.regEIP + 1   '�ؼ�.CPU��EIP�Ĵ���������ǵ�ǰ�쳣���ĵ�ַ,+1���ǽ���ǰִ���������,����һ��ַ��ʼִ��.
SoftSAFE = EXCEPTION_CONTINUE_EXECUTION
Unload Me '����ִ��,ִ�е�ַ��pContextRecord��ָ��
End Select
End Sub

Private Sub imgInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub ts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
