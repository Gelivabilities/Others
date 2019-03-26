VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "学生信息管理系统"
   ClientHeight    =   6225
   ClientLeft      =   1800
   ClientTop       =   1815
   ClientWidth     =   6630
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuGeneral 
      Caption         =   "通用(&G)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuStudent 
         Caption         =   "学生信息管理(&S)..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFind 
         Caption         =   "学生信息查询(&F)..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuTemp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "重新登录(&L)..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "关于(&A)..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'新兴源码下载:http://www.newxing.com/

Option Explicit

'表示当前的用户类型
'0---管理员类型的用户; 1---学生类型的用户
Public mnUserType As Integer
'表示当前登录的用户名
Public msUserName As String

Private Sub MDIForm_Activate()
'根据不同的用户类型，使相应的菜单项可见
  Select Case mnUserType
    Case 0:                       '以管理员身份登录
        mnuFind.Visible = True
    Case 1:                       '以学生身份登录， 只能查询自己的信息
        mnuFind.Visible = False
  End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("真的要对出本系统吗？", vbQuestion + vbYesNo + vbDefaultButton2, "退出") = vbNo Then
    Cancel = 1
  End If
End Sub

Private Sub mnuAbout_Click()
 '显示“关于...”窗口
  Load frmSplash
  frmSplash.mbAbout = True
  frmSplash.Show vbModal
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuFind_Click()
   frmStudent.Show
   frmStudent.cmdSeek.Value = True
End Sub

Private Sub mnuLogin_Click()
  If MsgBox("若重新登录，所有窗体都将关闭！是否重新登录？", _
    vbQuestion + vbYesNo + vbDefaultButton2, "重新登录") = vbYes Then
     Unload MDIMain
     frmLogin.Show
  End If
End Sub

Private Sub mnuStudent_Click()
   If mnUserType = 0 Then   '若为管理员用户
      frmStudent.Show
   Else                     '若为学生类用户
      frmView.Show
   End If
End Sub
