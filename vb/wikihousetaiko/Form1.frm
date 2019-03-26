VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   1740
   ClientTop       =   795
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   10875
      TabIndex        =   2
      Top             =   0
      Width           =   10935
      Begin VB.Label Label1 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   10935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   8895
      Left            =   10920
      ScaleHeight     =   8835
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton Command2 
         Caption         =   "按类别"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "按曲包"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   8520
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   12480
      Top             =   240
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   50460
      Left            =   -4080
      TabIndex        =   0
      Top             =   -22800
      Width           =   15375
      ExtentX         =   27120
      ExtentY         =   89006
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Load()
'Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
'xmlHTTP1.Open "get", "http://www.wikihouse.com/taiko/index.php?iOS%A4%CE%BC%FD%CF%BF%B6%CA", True
'xmlHTTP1.send
'While xmlHTTP1.readyState <> 4
'DoEvents
'Wend
'Dim a, b As String, c, d As Long
'a = xmlHTTP1.responseText



Private Sub Command1_Click()
WebBrowser1.Navigate "http://www.wikihouse.com/taiko/index.php?iOS%A4%CE%BC%FD%CF%BF%B6%CA"
End Sub

Private Sub Command2_Click()
WebBrowser1.Navigate "http://www.wikihouse.com/taiko/index.php?iOS%A4%CE%BC%FD%CF%BF%B6%CA%2F%A5%B8%A5%E3%A5%F3%A5%EB%CA%CC"
End Sub

'清除前面内容

'b = a

'c = InStr(1, b, "組曲一番終曲")

'b = Left(a, c - 5)

'a = Replace(a, b, "")


'清除style与★之间的内容

'For i = 1 To 500

'c = InStr(1, a, "style_td")

'd = InStr(1, a, "★")

'b = Mid(a, c - 1, d - c + 2)

'a = Replace(a, b, "☆")

'Next

'Text1.Text = a
'Set xmlHTTP1 = Nothing

'End Sub


Private Sub Form_Load()
WebBrowser1.Top = -22800
WebBrowser1.Height = 50460
WebBrowser1.Width = 15375
WebBrowser1.Left = -4080
WebBrowser1.Navigate "http://www.wikihouse.com/taiko/index.php?iOS%A4%CE%BC%FD%CF%BF%B6%CA"
End Sub



Private Sub Timer1_Timer()
Label1.Caption = WebBrowser1.LocationURL
Form1.Caption = WebBrowser1.LocationName
If WebBrowser1.LocationURL = "http://www.wikihouse.com/taiko/index.php?iOS%A4%CE%BC%FD%CF%BF%B6%CA" Then
    WebBrowser1.Top = -22800
    WebBrowser1.Height = 50460
    Else
    If WebBrowser1.LocationURL = "http://www.wikihouse.com/taiko/index.php?iOS%A4%CE%BC%FD%CF%BF%B6%CA%2F%A5%B8%A5%E3%A5%F3%A5%EB%CA%CC" Then
    WebBrowser1.Top = -9500
    WebBrowser1.Height = 30740
    Else
    WebBrowser1.Top = -6700
    WebBrowser1.Height = 17940
    End If
End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
    WebBrowser1.Silent = True
    Label2.Caption = "Loading..."
End Sub
Private Sub WebBrowser1_DownloadComplete()
    WebBrowser1.Silent = True
    Label2.Caption = ""
End Sub
