VERSION 5.00
Begin VB.Form frmRotate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "设置旋转参数"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   Icon            =   "frmRotate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   577
   ShowInTaskbar   =   0   'False
   Begin ICEE.ICEE_KEY CMDNO 
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4440
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_KEY CMDOK 
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   4440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7890
      Picture         =   "frmRotate.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   8
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7890
      Picture         =   "frmRotate.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   7890
      Picture         =   "frmRotate.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   15
      Width           =   750
   End
   Begin VB.TextBox txtAngle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00231C09&
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "0"
      Top             =   2520
      Width           =   855
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00231C09&
      BorderStyle     =   0  'None
      Height          =   3660
      Left            =   3855
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   4
      Top             =   735
      Width           =   4545
      Begin VB.Image IA 
         Enabled         =   0   'False
         Height          =   240
         Index           =   4
         Left            =   4200
         Picture         =   "frmRotate.frx":0636
         Top             =   3360
         Width           =   240
      End
   End
   Begin VB.PictureBox picBackColor 
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1680
      ScaleHeight     =   255
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   2
      Left            =   1680
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label LA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "旋转图像"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape SB 
      BorderColor     =   &H00808080&
      Height          =   3735
      Index           =   1
      Left            =   240
      Top             =   720
      Width           =   3135
   End
   Begin VB.Shape SB 
      BorderColor     =   &H00808080&
      Height          =   3735
      Index           =   0
      Left            =   3840
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "背景色:"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   630
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "角度:"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   450
   End
End
Attribute VB_Name = "frmRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

Option Explicit

Private sAngle As Single
Private clsSrc As CLSPICDIBS
Private clsDst As CLSPICDIBS

Private Sub CreateSrc()
    Dim sngScale As Single
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    
    With frmGraphic.clsDIB
         If .Width > .Height Then
            sngScale = .Width / 350
            If .Height / sngScale > 300 Then sngScale = .Height / 300
         Else
            sngScale = .Height / 300
            If .Width / sngScale > 350 Then sngScale = .Width / 350
         End If
         lngWidth = CLng(.Width / sngScale)
         lngHeight = CLng(.Height / sngScale)
         Set clsSrc = New CLSPICDIBS
         clsSrc.Create lngWidth, lngHeight
         .PaintPicture clsSrc.hdc, APIStretchBlt, 0, 0, lngWidth, lngHeight
    End With
    lngLeft = 240
    lngTop = 360
    With picPreview
         clsSrc.PaintPicture .hdc
         .Refresh
    End With
End Sub

Private Sub PaintDIB(ByVal Angle As Single, ByVal BackClr As Long)
    Dim sngScale As Single
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim clsProcess As New cDIBProcess
    
    If Angle = 0# Then
       lngWidth = clsSrc.Width
       lngHeight = clsSrc.Height
    Else
       If Not (clsDst Is Nothing) Then clsDst.ClearUp
       Set clsDst = clsSrc
       clsProcess.RotateDIB clsDst, Angle, BackClr
    
       With clsDst
            If .Width > .Height Then
               sngScale = .Width / 350
               If .Height / sngScale > 300 Then sngScale = .Height / 300
            Else
               sngScale = .Height / 300
               If .Width / sngScale > 350 Then sngScale = .Width / 350
            End If
            lngWidth = CLng(.Width / sngScale)
            lngHeight = CLng(.Height / sngScale)
       End With
    End If
    lngLeft = 240
    lngTop = 360
    With picPreview
    .Cls
         If Angle = 0# Then
            clsSrc.PaintPicture .hdc
         Else
            clsDst.PaintPicture .hdc, APIStretchBlt, 0, 0, lngWidth, lngHeight
         End If
         .Refresh
    End With
    Set clsProcess = Nothing
End Sub

Private Sub CMDNO_CLICK()
Unload Me
End Sub

Private Sub CMDOK_CLICK()
    frmGraphic.sngAngle = sAngle
    frmGraphic.lngBackColor = picBackColor.BackColor
    Unload Me
End Sub

Private Sub Form_Load()
Call CreateSrc
Me.BackColor = COLOR_NOR
CMDOK.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
CMDNO.SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
CMDOK.SETTXT "确    定"
CMDNO.SETTXT "取    消"

Me.Move frmGraphic.Left + (frmGraphic.Width - Me.Width) / 2, frmGraphic.Top + (frmGraphic.Height - Me.Height) / 2
Call SeekMe(Me)
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsSrc = Nothing
    Set clsDst = Nothing
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub picBackColor_Click()
On Error GoTo ERR
picBackColor.BackColor = frmma.ShowColor(Me)
picPreview.BackColor = picBackColor.BackColor
txtAngle_Change
ERR:
Exit Sub
End Sub

Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False

End Sub

Private Sub PO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
End Sub

Private Sub txtAngle_Change()
    On Error GoTo ErrorHandle
    If txtAngle.Text = "" Then
       sAngle = 0#
    ElseIf Mid$(txtAngle.Text, 1, 1) = "." Then
       sAngle = CSng(Val("0" & txtAngle.Text))
    ElseIf Mid$(txtAngle.Text, 1, 2) = "-." Then
       sAngle = -CSng(Val("0" & Mid$(txtAngle.Text, 2)))
    Else
       sAngle = CSng(txtAngle.Text)
    End If
    Call PaintDIB(sAngle, picBackColor.BackColor)
    Exit Sub
ErrorHandle:
End Sub

Private Sub txtAngle_KeyPress(KeyAscii As Integer)
    KeyAscii = VailText(KeyAscii, "0123456789.-", True)
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
