VERSION 5.00
Begin VB.UserControl ICEE_TEXT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5E7D0&
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ToolboxBitmap   =   "ICEE_TEXT.ctx":0000
   Begin ICEE.IVScroll SCRO 
      Height          =   4455
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7858
      MinV            =   0
      MaxV            =   20
      Value           =   0
      SmallChange     =   1
      LargeChange     =   10
   End
   Begin VB.PictureBox PTXT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00F5E7D0&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label LBTXT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAPTION"
         Height          =   180
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4500
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "ICEE_TEXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Public MyTxt As String, HASLINE As Boolean

Private Sub LBTXT_Change()
LBTXT.Move 0, 0
SCRO.MaxV = LBTXT.Height - PTXT.ScaleHeight
MyTxt = LBTXT.Caption
If LBTXT.Height > PTXT.ScaleHeight - 50 Then
SCRO.Visible = True
Else
SCRO.Visible = False
End If
End Sub
Private Sub PTXT_Resize()
LBTXT.Move 0, 0, PTXT.ScaleWidth, PTXT.ScaleHeight
End Sub

Private Sub SCRO_Change()
LBTXT.Top = -SCRO.value

End Sub
Private Sub UserControl_Initialize()
On Error Resume Next
SCRO.MaxV = 0
HASLINE = False
SCRO.LargeChange = 50
SCRO.value = 0
LBTXT.Caption = ""
LBTXT.FontName = "Î¢ÈíÑÅºÚ"
End Sub

Sub SETTXT(TXT As String)
LBTXT.Caption = TXT
End Sub

Private Sub UserControl_Resize()

SCRO.Move UserControl.ScaleWidth - SCRO.Width - 5, 5, 17, UserControl.ScaleHeight - 10
PTXT.Move 8, 5, UserControl.ScaleWidth - SCRO.Width - 10, UserControl.ScaleHeight - 10

UserControl.Cls
If HASLINE = True Then UserControl.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H808080, B
End Sub
Sub SETBACKCOLOR(Color As Long)
UserControl.BackColor = Color
PTXT.BackColor = Color
UserControl.Cls
If HASLINE = True Then UserControl.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), &H808080, B
End Sub
Sub SETFORECOLOR(Color As Long)
LBTXT.ForeColor = Color
End Sub
