VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FRMNEWS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FEE4AC&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   474
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00261700&
      BorderStyle     =   0  'None
      Height          =   6570
      Left            =   15
      ScaleHeight     =   438
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   558
      TabIndex        =   0
      Top             =   525
      Width           =   8370
      Begin SHDocVwCtl.WebBrowser WEB 
         Height          =   6600
         Left            =   -30
         TabIndex        =   1
         Top             =   -30
         Width           =   8400
         ExtentX         =   14817
         ExtentY         =   11642
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
         Location        =   ""
      End
   End
   Begin VB.Image X3 
      Height          =   300
      Left            =   7560
      Picture         =   "FRMNEWS.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image X2 
      Height          =   300
      Left            =   7560
      Picture         =   "FRMNEWS.frx":0B84
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image X1 
      Height          =   300
      Left            =   7560
      Picture         =   "FRMNEWS.frx":1708
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "FRMNEWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim WithEvents M_Dom As MSHTML.HTMLDocument
Attribute M_Dom.VB_VarHelpID = -1
Private Sub Form_Load()
On Error Resume Next
Call PaintPng(App.Path & "\SKIN\DA_T.PNG", Me.hdc, 8, 8)
WEB.Navigate "http://minisite.qq.com/others08/index.htm"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub LA_Click()
Call CMV(Me)
End Sub

Private Function M_Dom_oncontextmenu() As Boolean
On Error Resume Next
WEB.Document.oncontextmenu = False
End Function

Private Sub M_Dom_onmousemove()
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub WEB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Set M_Dom = WEB.Document
End Sub
Private Sub WEB_FileDownload(ByVal ActiveDocument As Boolean, Cancel As Boolean)
On Error Resume Next
Cancel = True
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
If X3.Visible = False Then Unload Me
End Sub
