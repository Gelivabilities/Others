VERSION 5.00
Begin VB.Form FRMCOLOR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "颜色拾取器"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   Icon            =   "FRMCOLOR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   0
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   0
      Width           =   1920
      Begin VB.TextBox TxtColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   840
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   240
         Top             =   120
         Width           =   255
      End
      Begin VB.Label LblcolorB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "256"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   270
      End
      Begin VB.Label LblcolorG 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "256"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   270
      End
      Begin VB.Label LblcolorR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "256"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   270
      End
   End
End
Attribute VB_Name = "FRMCOLOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Dim ColorRGB As Long
Dim IsFlow As Boolean
Dim R As Integer, G As Integer, b As Integer
Dim r1 As String, G1 As String, b1 As String
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim hdc As Long
    Dim SX As Integer, SY As Integer
    Me.Move 0, 0, Screen.Width, Screen.Height
    FRMCOLOR.AutoRedraw = True '为了永久保存图像
    SX = Screen.Width \ Screen.TwipsPerPixelX
    SY = Screen.Height \ Screen.TwipsPerPixelY
    hdc = GetDC(0)
    BitBlt FRMCOLOR.hdc, 0, 0, SX, SY, hdc, 0, 0, vbSrcCopy
    FRMCOLOR.WindowState = 2
    FRMCOLOR.AutoRedraw = False  '为了防止窗体闪烁
    FRMCOLOR.Show
    ReleaseDC 0, hdc
    IsFlow = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  ColorRGB = FRMCOLOR.POINT(x, y)
  Call GetRGB(x, y)
End Sub

 Sub GetRGB(x As Single, y As Single)
   If IsFlow Then '超出显示范围后调整控件位置
     If x > FRMCOLOR.Width / 2 Then PO.Left = x - PO.Width - 100
     If y > FRMCOLOR.Height / 2 Then PO.Top = y - PO.Height - 100
     If x <= FRMCOLOR.Width / 2 Then PO.Left = x + 100
     If y <= FRMCOLOR.Height / 2 Then PO.Top = y + 100
   End If
   Shape1.FillStyle = 0 '实色填充
   Shape1.FillColor = ColorRGB
   R = ColorRGB Mod 256
   G = ColorRGB \ 256 Mod 256
   b = ColorRGB \ 256 \ 256
   LblcolorR.Caption = R
   LblcolorG.Caption = G
   LblcolorB.Caption = b
   r1 = IIf(R <> 0, Hex(R), "00") '解决一个0的问题
   G1 = IIf(G <> 0, Hex(G), "00")
   b1 = IIf(b <> 0, Hex(b), "00")
   TxtColor.Text = "#" & r1 & G1 & b1
   
 End Sub

Private Sub Form_Terminate()
Set FRMCOLOR = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
RGBColor = Shape1.FillColor
FRMCOPEN.PSD.BackColor = Shape1.FillColor
FRMBOARD.PF.BackColor = Shape1.FillColor
FRMCOPEN.t1(4).Text = R
FRMCOPEN.t1(5).Text = G
FRMCOPEN.t1(6).Text = b

End Sub

Private Sub IA_Click(Index As Integer)
End Sub
