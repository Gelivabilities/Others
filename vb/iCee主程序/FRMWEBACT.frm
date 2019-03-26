VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FRMWEBACT 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00AD7900&
   BorderStyle     =   0  'None
   Caption         =   "活动接受"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox IMD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AD7900&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   5
      Left            =   1680
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox IMD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AD7900&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   4
      Left            =   1440
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox IMD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AD7900&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   3
      Left            =   1080
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox IMD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AD7900&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   2
      Left            =   840
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox IMD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AD7900&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   600
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox IMD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AD7900&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   0
      Left            =   360
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox LSTLINK 
      Height          =   5640
      Left            =   8640
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
   Begin SHDocVwCtl.WebBrowser WEB 
      Height          =   3015
      Left            =   8760
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      ExtentX         =   4895
      ExtentY         =   5318
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
Attribute VB_Name = "FRMWEBACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TGUID
   Data1                            As Long
   data2                            As Integer
   Data3                            As Integer
   Data4(0 To 7)                    As Byte
End Type
  
'// 用来加载Internet上的图片
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
  
'// 从Internet上加载图片
Public Function LoadPicture(ByVal strFilename As String) As PICTURE
   Dim IID  As TGUID
   With IID
      .Data1 = &H7BF80980
      .data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(2) = &H0
      .Data4(3) = &HAA
      .Data4(4) = &H0
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
   End With
     
   On Error GoTo LocalErr
     
   OleLoadPicturePath StrPtr(strFilename), 0&, 0&, 0&, IID, LoadPicture
   Exit Function
LocalErr:
   Set LoadPicture = VB.LoadPicture(strFilename)
   ERR.Clear
End Function
Private Sub Form_Load()
Me.Hide
WEB.Navigate "http://hi.baidu.com/iceeorgan/item/dfb41b7f65775f6eef1e53c8"
End Sub
Private Sub WEB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Dim I As Integer, s As String, SB As String, CAST As String
LSTLINK.Clear
s = ""
SB = ""
CAST = "[活动]"
For I = 0 To WEB.Document.links.Length - 1
If WEB.Document.links.Item(I) <> s Then
SB = WEB.Document.links.Item(I).innerText 'SB是页面中所有超链接文字
s = WEB.Document.links.Item(I) 'S是页面中所有超链接
If Left(SB, Len(CAST)) = CAST Then LSTLINK.AddItem SB & "|" & s
End If
Next I
WEB.Silent = True
If LSTLINK.ListCount = 0 Then Exit Sub
IMD(0).PICTURE = LoadPicture(Split(LSTLINK.List(0), "|")(2))
IMD(1).PICTURE = LoadPicture(Split(LSTLINK.List(1), "|")(2))
IMD(2).PICTURE = LoadPicture(Split(LSTLINK.List(2), "|")(2))
IMD(3).PICTURE = LoadPicture(Split(LSTLINK.List(3), "|")(2))
IMD(4).PICTURE = LoadPicture(Split(LSTLINK.List(4), "|")(2))
IMD(5).PICTURE = LoadPicture(Split(LSTLINK.List(5), "|")(2))
For I = 0 To IMD.Count - 1
frmma.IWG(I).SETCOLOR COLOR_NOR, COLOR_HIGH
frmma.IWG(I).SETIMG IMD(I)
frmma.IWG(I).MYTIT = (Split(LSTLINK.List(I), "|")(1))
frmma.IWG(I).SETTIP (Replace(Split(LSTLINK.List(I), "|")(0), "[活动]", ""))
Next
frmma.PF(11).Visible = True
frmma.PF(11).ZOrder 0
Call frmma.RUNSAFE
IS_FIRST_LOAD_ACT = False
End Sub
