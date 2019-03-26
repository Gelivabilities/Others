VERSION 5.00
Begin VB.Form FrmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00B37417&
   BorderStyle     =   0  'None
   ClientHeight    =   1725
   ClientLeft      =   495
   ClientTop       =   495
   ClientWidth     =   3495
   Icon            =   "FrmChat.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmChat.frx":038A
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00B98200&
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   0
      Picture         =   "FrmChat.frx":1458E
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   234
      TabIndex        =   0
      Top             =   0
      Width           =   3510
      Begin VB.TextBox TxtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H0008AF66&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1230
         Width           =   3015
      End
      Begin VB.TextBox TxtRes 
         BackColor       =   &H0008AF66&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Image X1 
         Height          =   240
         Left            =   1920
         Picture         =   "FrmChat.frx":28792
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image X2 
         Height          =   240
         Left            =   2520
         Picture         =   "FrmChat.frx":287ED
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image X3 
         Height          =   240
         Left            =   2160
         Picture         =   "FrmChat.frx":28850
         Top             =   120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image IU 
         Height          =   240
         Left            =   3075
         Picture         =   "FrmChat.frx":288AB
         ToolTipText     =   "退出谈话"
         Top             =   195
         Width           =   240
      End
      Begin VB.Label La 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户ID"
         ForeColor       =   &H00E0E0E0&
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   225
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Public RecieversID As String
Public SenderID
Public SenderName
Public NewChatBack As New FrmChatBk
Private WithEvents PSubClass As cSubclass
Attribute PSubClass.VB_VarHelpID = -1
Private Sub Form_Activate()
PSubClass.AddWindowMsgs Me.hwnd
Call MoveWindow(NewChatBack.hwnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, 243, 117, True)
End Sub

Private Sub Form_Load()
Dim WindowRegion As Long
IU.PICTURE = X1.PICTURE
WindowRegion = getpic(PO)
SetWindowRgn Me.hwnd, WindowRegion, True
Load NewChatBack
Set PSubClass = New cSubclass '继承无拖影
Call PSubClass.AddWindowMsgs(Me.hwnd)  '继承无拖影
Call NewChatBack.Show
LA.Caption = SenderName
RecieversID = SenderID
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
Me.ZOrder 0
Call SeekMe(Me)
oldproc = GetWindowLong(TxtRes.hwnd, GWL_WNDPROC)
SetWindowLong TxtRes.hwnd, GWL_WNDPROC, AddressOf TextWndProc
oldproc = GetWindowLong(TxtMessage.hwnd, GWL_WNDPROC)
SetWindowLong TxtMessage.hwnd, GWL_WNDPROC, AddressOf TextWndProc
If Sound = 1 Then sndPlaySound App.Path + "\Sound\MSG.wav", 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
SetWindowLong TxtRes.hwnd, GWL_WNDPROC, oldproc
SetWindowLong TxtMessage.hwnd, GWL_WNDPROC, oldproc
Unload NewChatBack
End Sub
Private Sub IU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If IU.PICTURE = X2.PICTURE Then IU.PICTURE = X3.PICTURE
End Sub
Private Sub IU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE = X1.PICTURE Then IU.PICTURE = X2.PICTURE

End Sub
Private Sub IU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE = X3.PICTURE Then IU.PICTURE = X1.PICTURE
If Button = 1 Then Me.Hide: NewChatBack.Hide
End Sub
Private Sub LA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub
Private Sub PO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IU.PICTURE <> X1.PICTURE Then IU.PICTURE = X1.PICTURE
End Sub
Private Sub PSubClass_MsgCome(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, lng_hWnd As Long, uMsg As Long, wParam As Long, lParam As Long)
If bBefore Then
If uMsg = WM_MOVE Then Call MoveWindow(NewChatBack.hwnd, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, 243, 117, True)
End If
End Sub
Private Sub TxtMessage_KeyPress(KeyAscii As Integer)
'禁止 Ctrl+V 粘贴功能 keyascii=22 是 Ctrl+V
If KeyAscii = 22 Then KeyAscii = 0
If Len(Trim(TxtMessage.Text)) > 0 And KeyAscii = 13 Then Call 回复
End Sub
Sub 回复()
On Error Resume Next
Dim TempData As String
Dim Temp2 As String
TxtMessage.Enabled = False
Temp2 = Word(RecieversID, 1)
TempData = ".msg " & Replace(TxtMessage.Text, vbCrLf, "//crlf\\") '& " " & Me.La.Caption
frmma.Winsock1.SendData TempData
TxtMessage.Enabled = True
Me.Hide
NewChatBack.Hide
End Sub

Private Sub TxtMessage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then TxtMessage.SetFocus: Me.PopupMenu Frmm.文本
End Sub

