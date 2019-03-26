VERSION 5.00
Begin VB.Form FRMPIC_MAN 
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   13785
      Picture         =   "FRMPIC_MAN.frx":0000
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   13785
      Picture         =   "FRMPIC_MAN.frx":00E4
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   13785
      Picture         =   "FRMPIC_MAN.frx":01C8
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   4
      Top             =   15
      Width           =   750
   End
   Begin VB.FileListBox filHidden 
      Height          =   2250
      Left            =   9840
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox PICFRAME 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   120
      ScaleHeight     =   569
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   865
      TabIndex        =   0
      Top             =   840
      Width           =   12975
      Begin VB.Timer TMA 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox PS 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00633F0E&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   12240
         Picture         =   "FRMPIC_MAN.frx":02AC
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         ToolTipText     =   "·µ»Ø¶¥²¿"
         Top             =   7800
         Visible         =   0   'False
         Width           =   480
      End
      Begin ICEE.ucScrollbar vsbSlide 
         Height          =   1815
         Left            =   12240
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   3201
      End
      Begin VB.PictureBox picSlide 
         BackColor       =   &H00231C09&
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   0
         ScaleHeight     =   289
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   313
         TabIndex        =   1
         Top             =   120
         Width           =   4695
         Begin ICEE.ICEE_PIC optThumb 
            Height          =   3600
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   6350
         End
      End
   End
End
Attribute VB_Name = "FRMPIC_MAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call UnHook
H_DOS = 6
gHW = Me.hWnd 'Êó±ê¿Ø¼þ
Call Hook '»½ÐÑÊó±ê»¬ÂÖAPI
End Sub

Private Sub Form_Load()
On Error Resume Next
If frmma.Left >= FRMPIC_MAN.Width / 2 Then
FRMPIC_MAN.Move frmma.Left - FRMPIC_MAN.Width, frmma.Top
Else
FRMPIC_MAN.Move frmma.Left + frmma.Width, frmma.Top
End If

filHidden.Path = App.Path & "\MEDIA\MUSICPICTURE"
Call CreateThumbs
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\FM_T.PNG", Me.hdc, 8, 8)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub Form_Resize()
Dim X       As Long
Dim Y       As Long
Dim lIdx    As Long
Dim lCols   As Long
            PICFRAME.Move 5, 60, Me.ScaleWidth - 10, Me.ScaleHeight - 70
            vsbSlide.Move PICFRAME.ScaleWidth - vsbSlide.Width, 0, vsbSlide.Width, PICFRAME.ScaleHeight
            lCols = Int((PICFRAME.ScaleWidth) / OPTTHUMB(0).Width)
            For lIdx = 0 To OPTTHUMB.Count - 1
                X = (lIdx Mod lCols) * OPTTHUMB(0).Width
                Y = Int(lIdx / lCols) * OPTTHUMB(0).Height
                OPTTHUMB(lIdx).Move X, Y
            Next lIdx
            PICSLIDE.Width = lCols * OPTTHUMB(0).Width
            PICSLIDE.Height = Int(OPTTHUMB.Count / lCols) * OPTTHUMB(0).Height
            If Int(OPTTHUMB.Count / lCols) < (OPTTHUMB.Count / lCols) Then
                PICSLIDE.Height = PICSLIDE.Height + OPTTHUMB(0).Height
            End If
            vsbSlide.value = 0
            vsbSlide.Max = PICSLIDE.Height - PICFRAME.ScaleHeight
            If vsbSlide.Max < 0 Then
                vsbSlide.Max = 0
            Else
                vsbSlide.SmallChange = OPTTHUMB(0).Height
                vsbSlide.LargeChange = PICFRAME.ScaleHeight
            End If
            PS.Move PICFRAME.ScaleWidth - PS.Width - 5, PICFRAME.ScaleHeight - PS.Height - 5
End Sub
Private Sub CreateThumbs()
Dim iMaxLen As Integer
Dim X       As Long
Dim Y       As Long
Dim lIdx    As Long
Dim lPicCnt As Long
Dim lFilCnt As Long
Dim sPath   As String
Dim sText   As String
    filHidden.Refresh
    PICSLIDE.Move 0, 0, OPTTHUMB(0).Width, OPTTHUMB(0).Height
    PICSLIDE.Visible = False
    While OPTTHUMB.Count > 1
        Unload OPTTHUMB(OPTTHUMB.Count - 1)
    Wend
    DoEvents
    On Error Resume Next
    sPath = filHidden.Path
    sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
    lFilCnt = filHidden.ListCount
    If Len(sPath) > 0 Then
        For lIdx = 0 To filHidden.ListCount - 1
            ERR.Clear
                If lPicCnt > 0 Then
                    Load OPTTHUMB(lPicCnt)
                    Set OPTTHUMB(lPicCnt).Container = PICSLIDE
                End If
                DoEvents
                OPTTHUMB(lPicCnt).ToolTipText = filHidden.List(lIdx)
                OPTTHUMB(lPicCnt).SETPIC (filHidden.Path & "\" & filHidden.List(lIdx))
                OPTTHUMB(lPicCnt).Visible = True
                lPicCnt = lPicCnt + 1
        Next lIdx
        Call Form_Resize
        PICSLIDE.Visible = True
    End If
End Sub

Private Sub optThumb_MOUSEMOVE(INDEX As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub picSlide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub PS_Click()
If PICSLIDE.Top = 0 Then Exit Sub
If PICSLIDE.Top < 0 Then TMA.Enabled = True
End Sub

Private Sub TMA_Timer()
vsbSlide.value = vsbSlide.value - 50
If vsbSlide.value = 0 Then TMA.Enabled = False
End Sub

Private Sub vsbSlide_Change()
    PICSLIDE.Top = -vsbSlide.value
    PICFRAME.SetFocus
    If vsbSlide.value = 0 Then PS.Visible = False Else PS.Visible = True
End Sub
Private Sub vsbSlide_Scroll()
    vsbSlide_Change
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
