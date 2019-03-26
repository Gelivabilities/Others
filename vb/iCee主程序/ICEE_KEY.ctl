VERSION 5.00
Begin VB.UserControl ICEE_KEY 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ToolboxBitmap   =   "ICEE_KEY.ctx":0000
   Begin VB.PictureBox PH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00899F1E&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3240
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2760
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer TMIN 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Image I1 
      Height          =   600
      Index           =   3
      Left            =   1680
      Picture         =   "ICEE_KEY.ctx":0312
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image I1 
      Enabled         =   0   'False
      Height          =   630
      Index           =   1
      Left            =   1680
      Picture         =   "ICEE_KEY.ctx":0676
      Top             =   1080
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image I1 
      Enabled         =   0   'False
      Height          =   630
      Index           =   0
      Left            =   1440
      Picture         =   "ICEE_KEY.ctx":08B2
      Top             =   1080
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image I1 
      Height          =   585
      Index           =   2
      Left            =   1440
      Picture         =   "ICEE_KEY.ctx":0AEE
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label LBTIT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   90
   End
   Begin VB.Image IMNOR 
      Enabled         =   0   'False
      Height          =   630
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2205
   End
   Begin VB.Image IMLIGHT 
      Enabled         =   0   'False
      Height          =   630
      Left            =   960
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   3180
   End
End
Attribute VB_Name = "ICEE_KEY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  X As Long
  Y As Long
End Type
Private Mouse As POINTAPI
Private Button As RECT
Public HASLINE As Boolean
Private MouseIn As Boolean
Public Event Click()
Public M_STYLE As Integer, MY_TIT As String, L_M_R As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public IS_SELECT As Boolean
Private Sub IMLIGHT_Click()
RaiseEvent Click
End Sub

Private Sub IMNOR_Click()
RaiseEvent Click
End Sub

Private Sub IPL_Click()
Debug.Print "XXX"
End Sub

Private Sub LBTIT_Click()
RaiseEvent Click
End Sub

Private Sub LBTIT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBTIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub LBTIT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub
Private Sub TMIN_Timer()
GetCursorPos Mouse
GetWindowRect hwnd, Button
With Button
If Mouse.X >= .Left And Mouse.X <= .Right And Mouse.Y >= .Top And Mouse.Y <= .Bottom Then
MouseIn = True
If IMNOR.Visible = False Then IMNOR.Visible = True
If IMLIGHT.Visible = True Then IMLIGHT.Visible = False
Else
MouseIn = False
If IS_SELECT = True Then IMNOR.Visible = True: Exit Sub
If IMLIGHT.Visible = False Then IMLIGHT.Visible = True
If IMNOR.Visible = True Then IMNOR.Visible = False
TMIN.Enabled = False
End If
End With
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
TMIN.Enabled = True
IMLIGHT.Visible = True
IMNOR.Visible = False
LBTIT.Caption = ""
HASLINE = False

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
    TMIN.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))

End Sub

Private Sub UserControl_Resize()
Select Case M_STYLE
Case 0
IMNOR.PICTURE = I1(0).PICTURE
LBTIT.FOREColor = vbBlack
IMLIGHT.PICTURE = I1(1).PICTURE
IMLIGHT.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
IMNOR.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
Case 1

IMNOR.PICTURE = I1(2).PICTURE
IMLIGHT.PICTURE = I1(3).PICTURE
LBTIT.FOREColor = vbWhite
IMLIGHT.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
IMNOR.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
Case 2
Set IMNOR.PICTURE = PN.image
Set IMLIGHT.PICTURE = ph.image
If HASLINE = True Then
IMLIGHT.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
IMNOR.Move 1, 1, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2
Else
IMLIGHT.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
IMNOR.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End If
End Select

Select Case L_M_R
Case 0 'ÖÐ¼ä
LBTIT.Move (UserControl.ScaleWidth - LBTIT.Width) / 2, (UserControl.ScaleHeight - LBTIT.Height) / 2
Case 1 '×ó
LBTIT.Move 10, (UserControl.ScaleHeight - LBTIT.Height) / 2
Case 2 'ÓÒ
LBTIT.Move (UserControl.ScaleWidth - LBTIT.Width) - 10, (UserControl.ScaleHeight - LBTIT.Height) / 2
End Select
TMIN.Enabled = True
End Sub
Sub SETTXT(TXT As String)
LBTIT.Caption = TXT
MY_TIT = TXT
UserControl_Resize
End Sub
Sub SETCOLOR(NOR As Long, HIGH As Long, FUCK As Long)
On Error Resume Next
M_STYLE = 2
LBTIT.FOREColor = FUCK
PN.BackColor = HIGH
ph.BackColor = NOR
Set IMNOR.PICTURE = PN.image
Set IMLIGHT.PICTURE = ph.image
UserControl_Resize
End Sub

