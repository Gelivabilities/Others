VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FRMFAV 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00899F1E&
   BorderStyle     =   0  'None
   Caption         =   "收藏夹"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8310
   Icon            =   "FRMFAV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMFAV.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LFAV 
      Height          =   8295
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   14631
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   4213029
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "歌名"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "地址"
         Object.Width           =   6615
      EndProperty
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7530
      Picture         =   "FRMFAV.frx":1264
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7530
      Picture         =   "FRMFAV.frx":1348
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7530
      Picture         =   "FRMFAV.frx":142C
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   15
      Width           =   750
   End
End
Attribute VB_Name = "FRMFAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SAVEFAV()
On Error Resume Next
Dim filem As String
filem = App.Path & "\MEDIA\Favourite\M_Favourite.isw"
Dim i As Integer, tpList As ListItem
Open filem For Output As #1
For Each tpList In LFAV.ListItems
Print #1, tpList.Text
For i = 0 To 1
Print #1, tpList.SubItems(i)
Next
Next
Close #1
End Sub
Sub CLEAR_FAV()
LFAV.ListItems.Clear
Call SAVEFAV
Call LOADFAV
End Sub
Sub LOADFAV()
On Error Resume Next
Dim filem As String, tpStr As String, i As Integer
filem = App.Path & "\MEDIA\Favourite\M_Favourite.isw"
Debug.Print PathFileExists(filem)
If PathFileExists(filem) <> 0 Then
LFAV.ListItems.Clear
'加载收藏夹数据
Open filem For Input As #1
Do While Not EOF(1)
With LFAV.ListItems.Add()
For i = 0 To 1
Line Input #1, tpStr

If i = 0 Then
If tpStr <> "" Then .Text = tpStr
Else
If tpStr <> "" Then .SubItems(i) = tpStr
End If

.SmallIcon = 1
.Icon = 1
Next
End With
Loop
Close #1
End If

If LFAV.ListItems.Count = 0 Then LFAV.Visible = False Else LFAV.Visible = True
End Sub
Sub ADD_ITEM(TIT As String, URL As String)
On Error Resume Next
If TIT = "" Or URL = "" Then Exit Sub
With LFAV.ListItems.Add()
.Text = TIT
.SubItems(1) = URL
.Icon = 1
.SmallIcon = 1
End With

Call SAVEFAV
Call LOADFAV

Call CHECK_ITEM(frmma.Wm.URL)
End Sub

Sub CHECK_ITEM(URL As String)
On Error Resume Next
If LFAV.ListItems.Count = 0 Then Exit Sub
Dim i As Integer
For i = 1 To LFAV.ListItems.Count
If UCase(URL) = UCase(LFAV.ListItems(i).SubItems(1)) Then
If IS_NET = True Then FrmNetMusic.IMFAV.PICTURE = Frmm.pic(52).PICTURE
FAV_IT = True
Else
FAV_IT = False
If IS_NET = True Then FrmNetMusic.IMFAV.PICTURE = Frmm.pic(54).PICTURE
End If
Next
End Sub
Sub REMOVE_ITEM(TIT As String)
On Error Resume Next
Dim i As Integer
For i = 1 To LFAV.ListItems.Count
If LFAV.ListItems.Count = 0 Then Exit Sub
If UCase(LFAV.ListItems(i).Text) = UCase(TIT) Then LFAV.ListItems.REMOVE (i)
Next

Call SAVEFAV
Call LOADFAV

Call CHECK_ITEM(frmma.Wm.URL)
End Sub

Private Sub Form_Activate()
Me.Cls
Me.BackColor = COLOR_NOR
Call PaintPng(App.Path & "\SKIN\FAV_T.PNG", Me.hdc, 8, 8)
Call PaintPng(App.Path & "\SKIN\NO_FAV.PNG", Me.hdc, (Me.ScaleWidth - 200) / 2, (Me.ScaleHeight - 60) / 2)
Me.LFAV.BackColor = Me.BackColor
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B

End Sub

Private Sub Form_Load()
Call LOADFAV
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub LFAV_DblClick()
frmma.PLIST.AddItem FRMFAV.LFAV.SelectedItem.Text, "", FRMFAV.LFAV.SelectedItem.SubItems(1)
frmma.Wm.URL = FRMFAV.LFAV.SelectedItem.SubItems(1)
End Sub

Private Sub LFAV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Me.PopupMenu Frmm.我的收藏
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

