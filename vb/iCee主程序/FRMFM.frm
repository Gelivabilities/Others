VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FRMFM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "搜索封面"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   Icon            =   "FRMFM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   750
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser WEB 
      Height          =   9135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10935
      ExtentX         =   19288
      ExtentY         =   16113
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.FileListBox filHidden 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   2760
      Pattern         =   "*.bmp;*.dib;*.rle;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur"
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Image PSP 
      Height          =   1935
      Left            =   480
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "FRMFM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVE_SINGER As String, hPic As String
Private Sub Form_Load()
On Error Resume Next
fso.CreateFolder App.Path & "\MEDIA"
fso.CreateFolder App.Path & "\MEDIA\MUSICPICTURE"
fso.DeleteFolder App.Path & "\Thumbs\Singer_Thumbs"
fso.CreateFolder App.Path & "\Thumbs\Singer_Thumbs"
fso.DeleteFolder App.Path & "\Thumbs\B_ThumbS"
fso.CreateFolder App.Path & "\Thumbs\B_ThumbS"
filHidden.Path = App.Path & "\Thumbs\Singer_Thumbs"

Me.Move frmma.Left + (frmma.Width - Me.Width) / 2, frmma.Top + (Frmm.Height / 2 - Me.Height / 2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
IS_CAPTURE = False
KBS = 0
End Sub
Sub SERCHSINGER(SINGER As String)

SAVE_SINGER = SINGER
'WEB.Navigate "http://www.kuwo.cn/mingxing/" & SINGER & "/#intro" '
WEB.Navigate "http://image.youdao.com/search?q=" & SINGER & "&keyfrom=image.left&color=all&active=all&size=400x400"
End Sub
Private Sub WEB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
WEB.Silent = True
Call DOWNLOADAD(WEB, App.Path & "\Thumbs\Singer_Thumbs\", SAVE_SINGER)
Call 过滤
End Sub
Sub 设置头像()
On Error Resume Next
Call FileCopy(hPic, App.Path & "\MEDIA\MusicPicture\" & SAVE_SINGER & ".Bmp")
Call frmma.REGETSINGER
End Sub

Private Sub 过滤()
filHidden.Path = App.Path & "\Thumbs\Singer_Thumbs"
filHidden.Refresh
For I = 0 To filHidden.ListCount - 1
PSP.PICTURE = LoadPicture(filHidden.Path & "\" & filHidden.List(I))
If PSP.Width >= 400 Then hPic = filHidden.Path & "\" & filHidden.List(I): Exit For
filHidden.Refresh
Next
Call 设置头像
End Sub
Private Sub WEB_DownloadBegin()
IS_CAPTURE = True
End Sub

Private Sub WEB_FileDownload(ByVal ActiveDocument As Boolean, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Web_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
End Sub
