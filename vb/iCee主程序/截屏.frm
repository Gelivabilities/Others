VERSION 5.00
Begin VB.Form Capture 
   BackColor       =   &H00221C13&
   BorderStyle     =   0  'None
   Caption         =   "iCee截屏"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   DrawWidth       =   3
   Icon            =   "截屏.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4560
      Picture         =   "截屏.frx":038A
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   27
      Top             =   1080
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   3620
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   3240
         Picture         =   "截屏.frx":0C54
         ToolTipText     =   "退出截屏"
         Top             =   75
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   2440
         Picture         =   "截屏.frx":0FDE
         ToolTipText     =   "保存"
         Top             =   75
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1700
         Picture         =   "截屏.frx":1368
         ToolTipText     =   "撤销"
         Top             =   75
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   1220
         Picture         =   "截屏.frx":16F2
         ToolTipText     =   "添加箭头"
         Top             =   75
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   5
         Left            =   860
         Picture         =   "截屏.frx":1A7C
         ToolTipText     =   "添加文字"
         Top             =   75
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   6
         Left            =   480
         Picture         =   "截屏.frx":1E06
         ToolTipText     =   "添加椎圆"
         Top             =   75
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   7
         Left            =   100
         Picture         =   "截屏.frx":2190
         ToolTipText     =   "添加矩形"
         Top             =   75
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   2070
         Picture         =   "截屏.frx":251A
         ToolTipText     =   "重做"
         Top             =   75
         Width           =   240
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   7
         Left            =   80
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   6
         Left            =   450
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   5
         Left            =   820
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   4
         Left            =   1200
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   1
         Left            =   2040
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   0
         Left            =   1670
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   2
         Left            =   2410
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H005C6105&
         Height          =   300
         Index           =   3
         Left            =   3210
         Top             =   45
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00494849&
         Height          =   390
         Left            =   0
         Top             =   0
         Width           =   3620
      End
   End
   Begin VB.TextBox TextEdit 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   560
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   1
      Top             =   390
      Visible         =   0   'False
      Width           =   3620
      Begin VB.PictureBox PicColor 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   440
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   2415
         TabIndex        =   6
         Top             =   60
         Width           =   2415
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   0
            Left            =   480
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   23
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   1
            Left            =   720
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   22
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   2
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   21
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   3
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   20
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   4
            Left            =   1440
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   19
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   5
            Left            =   1680
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   18
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00800080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   6
            Left            =   1920
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   17
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   7
            Left            =   2160
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   16
            Top             =   0
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   8
            Left            =   480
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   15
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   9
            Left            =   720
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   14
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   10
            Left            =   960
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   13
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   11
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   12
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   12
            Left            =   1440
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   11
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   13
            Left            =   1680
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   10
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF00FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   14
            Left            =   1920
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   9
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   200
            Index           =   15
            Left            =   2160
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   8
            Top             =   220
            Width           =   200
         End
         Begin VB.PictureBox Pcolor 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   16
            Left            =   20
            ScaleHeight     =   420
            ScaleWidth      =   420
            TabIndex        =   7
            Top             =   0
            Width           =   420
         End
      End
      Begin VB.PictureBox PicFont 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   1095
         TabIndex        =   2
         Top             =   120
         Width           =   1095
         Begin VB.PictureBox PicCombox 
            AutoSize        =   -1  'True
            BackColor       =   &H00231C09&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   320
            Picture         =   "截屏.frx":28A4
            ScaleHeight     =   330
            ScaleWidth      =   750
            TabIndex        =   3
            Top             =   0
            Width           =   750
            Begin VB.TextBox Text1 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H00000000&
               Height          =   200
               Left            =   60
               TabIndex        =   4
               Text            =   "11"
               Top             =   60
               Width           =   340
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "截屏.frx":2CAD
            Left            =   320
            List            =   "截屏.frx":2CCC
            TabIndex        =   5
            Text            =   "11"
            Top             =   20
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   0
            Picture         =   "截屏.frx":2CF1
            Top             =   40
            Width           =   240
         End
         Begin VB.Image Image4 
            Height          =   330
            Index           =   0
            Left            =   120
            Picture         =   "截屏.frx":307B
            Top             =   480
            Width           =   750
         End
         Begin VB.Image Image4 
            Height          =   330
            Index           =   1
            Left            =   960
            Picture         =   "截屏.frx":3484
            Top             =   480
            Width           =   750
         End
         Begin VB.Image Image4 
            Height          =   330
            Index           =   2
            Left            =   1800
            Picture         =   "截屏.frx":41D6
            Top             =   480
            Width           =   750
         End
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   555
         Left            =   0
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   7
      Left            =   3705
      Picture         =   "截屏.frx":4F28
      Top             =   2460
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   6
      Left            =   4035
      Picture         =   "截屏.frx":4F85
      Top             =   2460
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   5
      Left            =   4305
      Picture         =   "截屏.frx":4FE2
      Top             =   2460
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   4
      Left            =   4635
      Picture         =   "截屏.frx":503F
      Top             =   2460
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   3
      Left            =   3630
      Picture         =   "截屏.frx":509C
      Top             =   2175
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   2
      Left            =   3960
      Picture         =   "截屏.frx":50F9
      Top             =   2175
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   1
      Left            =   4230
      Picture         =   "截屏.frx":5156
      Top             =   2175
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image DSB 
      Height          =   150
      Index           =   0
      Left            =   4560
      Picture         =   "截屏.frx":51B3
      Top             =   2205
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label LblPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "宽X高"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   3
      FillColor       =   &H0099FFFF&
      Height          =   480
      Left            =   120
      Top             =   1095
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Capture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------画文本用函数----------------------------------------------------
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'--------------画椎圆用函数----------------------------------------------------
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'-------------------------下拉列表框消息----------------------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const CB_SHOWDROPDOWN = &H14F
Const CB_GETDROPPEDSTATE = &H157
Private Type GUID
Data1 As Long
data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Private Type GdiplusStartupInput
GdiplusVersion As Long
DebugEventCallback As Long
SuppressBackgroundThread As Long
SuppressExternalCodecs As Long
End Type
Private Type EncoderParameter
GUID As GUID
NumberOfValues As Long
type As Long
Value As Long
End Type
Private Type EncoderParameters
Count As Long
Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal _
outputbuf As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal filename As Long, clsidEncoder As _
GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal cb As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'------------调用保存对话框--------------------------------
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Dim OriginalX As Single   '区域起点X坐标
Dim OriginalY As Single   '区域起点的Y坐标
Dim X1 As Single, y1 As Single, LeftL As Single, TopL As Single
Dim NewX As Single
Dim NewY As Single
Dim Status As String      '当前状态（正在选择区域或者拖动区域）
Dim ImgMove As String
Dim rc As RECT            '区域的范围
Dim MPoint As POINTAPI
Dim DPoint As POINTAPI
Dim ptInPic As Boolean     '鼠标是否位于pic上
Dim UnloadFrm As Long
Dim Edit As Boolean    '是否编辑状态
Dim EditStr As String  '编辑内容
Dim START As Boolean
Dim ImgIndex As Long   '记录单击Image2的索引
Dim X0 As Single, Y0 As Single

Private Type Bytes
 arr() As Byte
End Type
Dim Byt() As Bytes
Dim Indexs As Long, iCount As Long

Private Type POINTAPI
X As Long
Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Sub GetRGBColors(ByVal RGBColor As Long, ByRef RedColor As Long, ByRef GreenColor As Long, ByRef BlueColor As Long)
    RedColor = RGBColor Mod 256
    GreenColor = (RGBColor \ &H100) Mod 256
    BlueColor = (RGBColor \ &H10000) Mod 256
End Sub

Private Sub Form_Load()
    IS_CAPTURE = True
    Picture1.Top = -Picture1.Height
    Picture1.Visible = False
    Dim SourceDC As Long
    Me.AutoRedraw = True
    Me.ScaleMode = 3
    Screen.MousePointer = vbCrosshair      ' 将光标改为十字型
    SourceDC = CreateDC("DISPLAY", 0, 0, 0)
    BitBlt Me.hdc, 0, 0, Screen.Width / 15, Screen.Height / 15, SourceDC, 0, 0, &HCC0020  '拷贝当前屏幕到窗体
    DeleteDC SourceDC
    Me.WindowState = 2
    Me.PICTURE = Me.image
    Indexs = Indexs + 1
    ReDim Preserve Byt(Indexs)
    WriteP Byt(Indexs).arr(), Me.PICTURE
    Status = "draw"        '绘图状态
    Edit = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub DSB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Indexs = 1 Then
        Edit = False
    End If
    If Edit = False Then
        If Status = "move" Then
            ImgMove = "Start"
            GetCursorPos DPoint
            LeftL = Shape1.Left: TopL = Shape1.Top
            X1 = Shape1.Left + Shape1.Width: y1 = Shape1.Top + Shape1.Height
            Picture1.Visible = False
            Picture2.Visible = False
             Indexs = 1: iCount = 1
             ReDim Preserve Byt(Indexs)
        End If
    End If
End Sub

Private Sub DSB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Select Case Index
    Case 0: Screen.MousePointer = 8
    Case 1: Screen.MousePointer = 7
    Case 2: Screen.MousePointer = 6
    Case 3: Screen.MousePointer = 9
    Case 4: Screen.MousePointer = 6
    Case 5: Screen.MousePointer = 7
    Case 6: Screen.MousePointer = 8
    Case 7: Screen.MousePointer = 9
    End Select
    If ImgMove = "Start" Then
        GetCursorPos MPoint   '取得当前鼠标位置
        DSB(Index).Move MPoint.X, MPoint.Y
        Select Case Index
        Case 0   '左上移动
            Shape1.Move MPoint.X + DSB(Index).Width / 2, MPoint.Y + DSB(Index).Height / 2, X1 - MPoint.X, y1 - MPoint.Y
        Case 1   '上移动
            Shape1.Move LeftL, MPoint.Y + DSB(Index).Height / 2, X1 - LeftL, y1 - MPoint.Y
        Case 2   '右上移动
            Shape1.Move LeftL, MPoint.Y + DSB(Index).Height / 2, (MPoint.X - LeftL) + DSB(Index).Width / 2, y1 - MPoint.Y
        Case 3   '左移动
            Shape1.Move MPoint.X + DSB(Index).Width / 2, TopL, X1 - MPoint.X, y1 - TopL
        Case 4   '左下移动
            Shape1.Move MPoint.X + DSB(Index).Width / 2, TopL, X1 - MPoint.X, MPoint.Y - TopL
        Case 5  '下移动
            Shape1.Move LeftL, TopL, X1 - LeftL, MPoint.Y - TopL
        Case 6  '右下移动
            Shape1.Move LeftL, TopL, MPoint.X - LeftL, MPoint.Y - TopL
        Case 7  '右移动
            Shape1.Move LeftL, TopL, MPoint.X - LeftL, y1 - TopL
        End Select
        ImageMove
        LblPos.Caption = Shape1.Width & "x" & Shape1.Height
        LblPos.Move Shape1.Left + 2, Shape1.Top + 2
         lblInfo.Move Shape1.Left + 2, LblPos.Top + LblPos.Height + 2
        OriginalX = Shape1.Left
        OriginalY = Shape1.Top
    End If
End Sub

Private Sub DSB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMove = "Stop"
    If (Shape1.Top + Shape1.Height + 4 + Picture1.Height) > Screen.Height / 15 Then
        Picture1.Move (Shape1.Left + Shape1.Width) - Picture1.Width, (Shape1.Top + Shape1.Height) - Picture1.Height - 4
    Else
        Picture1.Move (Shape1.Left + Shape1.Width) - Picture1.Width, Shape1.Top + Shape1.Height + 4
    End If
    If Picture1.Left < 0 Then Picture1.Move 0
    Picture1.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 0
IS_CAPTURE = False
Call frmma.iCan
End Sub

Private Sub Image2_Click(Index As Integer)
    Dim I As Integer
    Select Case Index
    Case 0    '撤销
        Indexs = Indexs - 1
        If Indexs > 0 Then
            Set Me.PICTURE = ReadP(Byt(Indexs).arr())
        Else
            Status = "draw"
            Shape1.Visible = False
            Picture1.Visible = False
            Picture2.Visible = False
            LblPos.Visible = False
            lblInfo.Visible = False
            Shape1.Width = 0
            Shape1.Height = 0
            For I = 0 To 7
                DSB(I).Visible = False
            Next
            Indexs = 1
            Edit = False
            ' ReDim Byt(1)
        End If
    Case 2
        Dim PicBool As Boolean
        If Picture2.Visible = True Then PicBool = True Else PicBool = False
        Picture1.Visible = False         '如果选区包含部分提示图片，则需要把图片先隐藏.
        Picture2.Visible = False
        Sleep 10                         '有时候没有这两句会使得shape1也显示在截取的区域里
        DoEvents
        Shape1.Visible = False
        ScrnCap Shape1.Left, Shape1.Top, Shape1.Left + Shape1.Width, Shape1.Top + Shape1.Height

        Picture1.Visible = True
        If PicBool = True Then Picture2.Visible = True
        Shape1.Visible = True
        '------------------------------------------
        Call CutdSave   '保存截图
        Unload Me
    Case 3
        Unload Me
    Case 4
        Picture2.Left = Picture1.Left
        If (Picture1.Top + Picture1.Height + 3 + Picture2.Height) > Screen.Height / 15 Then
            Picture2.Top = Picture1.Top - Picture2.Height - 3
        Else
            Picture2.Top = Picture1.Top + Picture1.Height + 3
        End If
        PicFont.Visible = False
        PicColor.Left = PicFont.Left
        Picture2.Width = 3620 / 15 - PicFont.Width - 2
        Shape4.Width = Picture2.ScaleWidth
        If Picture2.Visible = False Then
            Picture2.Visible = True
        Else
            If ImgIndex = Index Then Picture2.Visible = False    '如果当前单击的按钮索引与记录索引相同就将Picture2隐藏
        End If
        Edit = True
        EditStr = Image2(Index).ToolTipText
    Case 5
        Picture2.Left = Picture1.Left
        If (Picture1.Top + Picture1.Height + 3 + Picture2.Height) > Screen.Height / 15 Then
            Picture2.Top = Picture1.Top - Picture2.Height - 3
        Else
            Picture2.Top = Picture1.Top + Picture1.Height + 3
        End If
        PicFont.Visible = True
        PicColor.Left = 78.667
        Picture2.Width = 3620 / 15
        Shape4.Width = Picture2.ScaleWidth
        If Picture2.Visible = False Then
            Picture2.Visible = True
        Else
            If ImgIndex = Index Then Picture2.Visible = False   '如果当前单击的按钮索引与记录索引相同就将Picture2隐藏
        End If
        Edit = True
        EditStr = Image2(Index).ToolTipText
    Case 6
        Picture2.Left = Picture1.Left
        If (Picture1.Top + Picture1.Height + 3 + Picture2.Height) > Screen.Height / 15 Then
            Picture2.Top = Picture1.Top - Picture2.Height - 3
        Else
            Picture2.Top = Picture1.Top + Picture1.Height + 3
        End If
        PicFont.Visible = False
        PicColor.Left = PicFont.Left
        Picture2.Width = 3620 / 15 - PicFont.Width - 2
        Shape4.Width = Picture2.ScaleWidth
        If Picture2.Visible = False Then
            Picture2.Visible = True
        Else
            If ImgIndex = Index Then Picture2.Visible = False   '如果当前单击的按钮索引与记录索引相同就将Picture2隐藏
        End If
        Edit = True
        EditStr = Image2(Index).ToolTipText
    Case 7
        Picture2.Left = Picture1.Left
        If (Picture1.Top + Picture1.Height + 3 + Picture2.Height) > Screen.Height / 15 Then
            Picture2.Top = Picture1.Top - Picture2.Height - 3
        Else
            Picture2.Top = Picture1.Top + Picture1.Height + 3
        End If
        PicFont.Visible = False
        PicColor.Left = PicFont.Left
        Picture2.Width = 3620 / 15 - PicFont.Width - 2
        Shape4.Width = Picture2.ScaleWidth
        If Picture2.Visible = False Then
            Picture2.Visible = True
        Else
            If ImgIndex = Index Then Picture2.Visible = False  '如果当前单击的按钮索引与记录索引相同就将Picture2隐藏
        End If
        Edit = True
        EditStr = Image2(Index).ToolTipText
    Case 1    '重做
        Indexs = Indexs + 1
        If Indexs <= iCount Then
            Set Me.PICTURE = ReadP(Byt(Indexs).arr())
        Else
            Indexs = iCount
        End If
    End Select
    ImgIndex = Index
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Edit = False Then
        Picture1.Visible = False
        If Status = "draw" Then          '如果是抓取状态
            Shape1.Visible = True
            Shape1.Width = 0
            Shape1.Height = 0
            OriginalX = X
            OriginalY = Y                '起点坐标
            Shape1.Left = OriginalX
            Shape1.Top = OriginalY
            If Button = 2 Then
                UnloadFrm = 0
            End If
        Else
            Screen.MousePointer = vbCrosshair      ' 将光标改为十字型
            rc.Left = Shape1.Left
            rc.Right = Shape1.Left + Shape1.Width
            rc.Top = Shape1.Top
            rc.Bottom = Shape1.Top + Shape1.Height
            If PtInRect(rc, X, Y) Then     '如果按下的点位于区域内
                NewX = X
                NewY = Y               '则移动区域
                If Button = 2 Then
                    Shape1.Width = 0
                    Shape1.Height = 0
                    OriginalX = X
                    OriginalY = Y
                    Shape1.Left = OriginalX
                    Shape1.Top = OriginalY
                    Shape1.Visible = False
                    LblPos.Visible = False
                    lblInfo.Visible = False
                    For I = 0 To 7
                        DSB(I).Visible = False
                    Next
                    Screen.MousePointer = 0
                    Status = "draw"            '状态恢复到抓取
                    UnloadFrm = 1
                End If
            Else                           '否则重新画一个区域
                Shape1.Width = 0
                Shape1.Height = 0
                OriginalX = X
                OriginalY = Y
                Shape1.Left = OriginalX
                Shape1.Top = OriginalY
                Shape1.Visible = False
                LblPos.Visible = False
                lblInfo.Visible = False
                For I = 0 To 7
                    DSB(I).Visible = False
                Next
                Screen.MousePointer = 0
                Status = "draw"            '状态恢复到抓取
                UnloadFrm = 1
            End If
        End If
    Else
        If X > Shape1.Left And X < Shape1.Left + Shape1.Width And Y > Shape1.Top And Y < Shape1.Top + Shape1.Height Then
            START = True: Me.AutoRedraw = False: X0 = X: Y0 = Y
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If START = True Then    '开始编辑
        Me.AutoRedraw = True
        Dim X2 As Single, y2 As Single
        If START = True Then
            Me.Cls
            If X > Shape1.Left + Shape1.Width Then
                X2 = Shape1.Left + Shape1.Width
            ElseIf X < Shape1.Left Then
                X2 = Shape1.Left
            Else
                X2 = X
            End If
            If Y > Shape1.Top + Shape1.Height Then
                y2 = Shape1.Top + Shape1.Height
            ElseIf Y < Shape1.Top Then
                y2 = Shape1.Top
            Else
                y2 = Y
            End If
            Select Case EditStr
            Case "添加箭头"
                Call Arrow(Me, X0, Y0, X2, y2, 10, Pcolor(16).BackColor)
                '----------将当前窗体Image设置为窗体Picture----------
                START = False
                Me.PICTURE = Me.image
                Indexs = Indexs + 1
                iCount = Indexs
                ReDim Preserve Byt(Indexs)
                WriteP Byt(Indexs).arr(), Me.PICTURE
            Case "添加文字"
                TextEdit.FOREColor = Pcolor(16).BackColor
                Me.FontSize = Text1.Text
                TextEdit.FontSize = Text1.Text
                If TextEdit.Visible = False Then
                    TextEdit.Left = X: TextEdit.Top = Y
                    TextEdit.Width = 375 / 15
                    TextEdit.Visible = True: TextEdit.SetFocus
                Else
                    SetTextColor Me.hdc, Pcolor(16).BackColor
                    TextOut Me.hdc, TextEdit.Left, TextEdit.Top, TextEdit, LenB(StrConv(TextEdit, vbFromUnicode))
                    TextEdit.Visible = False: TextEdit = ""
                    '----------将当前窗体Image设置为窗体Picture----------
                    START = False
                    Me.PICTURE = Me.image
                    Indexs = Indexs + 1
                    iCount = Indexs
                    ReDim Preserve Byt(Indexs)
                    WriteP Byt(Indexs).arr(), Me.PICTURE
                End If
            Case "添加椎圆"
                'Call MoveCircle(Me, x0, y0, X2, Y2, Pcolor(16).BackColor)
                Dim tmppen As Long
                Dim pen As Long
                pen = CreatePen(0, 1, Pcolor(16).BackColor)  '创建一个画笔
                tmppen = SelectObject(Me.hdc, pen)      '选定一个刷子
                Ellipse Me.hdc, X0, Y0, X2, y2       '画图
                SelectObject Me.hdc, tmppen    '删除对象
                DeleteObject pen
                '----------将当前窗体Image设置为窗体Picture----------
                START = False
                Me.PICTURE = Me.image
                Indexs = Indexs + 1
                iCount = Indexs
                ReDim Preserve Byt(Indexs)
                WriteP Byt(Indexs).arr(), Me.PICTURE
            Case "添加矩形"
                Me.Line (X0, Y0)-(X2, y2), Pcolor(16).BackColor, B     '画矩形
                '----------将当前窗体Image设置为窗体Picture----------
                START = False
                Me.PICTURE = Me.image
                Indexs = Indexs + 1
                iCount = Indexs
                ReDim Preserve Byt(Indexs)
                WriteP Byt(Indexs).arr(), Me.PICTURE
            End Select
        End If
    End If

    If Button = 1 Then
        If Status = "draw" Then
            Status = "move"
        End If
        OriginalX = Shape1.Left   '更新OriginalX，因为选择区域时可能会出现shape的right点大于left点
        OriginalY = Shape1.Top
        If (Shape1.Top + Shape1.Height + 4 + Picture1.Height) > Screen.Height / 15 Then
            Picture1.Move (Shape1.Left + Shape1.Width) - Picture1.Width, (Shape1.Top + Shape1.Height) - Picture1.Height - 4
        Else
            Picture1.Move (Shape1.Left + Shape1.Width) - Picture1.Width, Shape1.Top + Shape1.Height + 4
        End If
        If Picture1.Left < 0 Then Picture1.Move 0
        Picture1.Visible = True
    Else
        If UnloadFrm = 0 Then Unload Me
    End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RGBColor As Long, Red As Long, Green As Long, Blue As Long
    X1 = X: y1 = Y
    RGBColor = GetPixel(Me.hdc, X, Y)
    GetRGBColors RGBColor, Red, Green, Blue
    lblInfo.Caption = "RGB(" & Red & "," & Green & "," & Blue & ")"
    Dim Info As String
    If Edit = False Then    '编辑状态
        If X > Shape1.Left And Y > Shape1.Top And X < (Shape1.Left + Shape1.Width) And Y < (Shape1.Top + Shape1.Height) Then
            Screen.MousePointer = 5
        Else
            Screen.MousePointer = vbCrosshair
        End If
        If Button = 1 Then
            Shape1.Visible = False
            LblPos.Visible = False
            lblInfo.Visible = False
            If Status = "draw" Then            '如果是绘图状态
                If X > OriginalX And Y > OriginalY Then           '根据鼠标位置调整shape1的大小和位置
                    Shape1.Move OriginalX, OriginalY, X - OriginalX, Y - OriginalY
                ElseIf X < OriginalX And Y > OriginalY Then
                    Shape1.Move X, OriginalY, OriginalX - X, Y - OriginalY
                ElseIf X > OriginalX And Y < OriginalY Then
                    Shape1.Move OriginalX, Y, X - OriginalX, OriginalY - Y
                ElseIf X < OriginalX And Y < OriginalY Then
                    Shape1.Move X, Y, OriginalX - X, OriginalY - Y
                End If
                Info = Shape1.Width & "x" & Shape1.Height             '显示当前区域的大小
                LblPos.Move Shape1.Left + 2, Shape1.Top + 2
                lblInfo.Move Shape1.Left + 2, LblPos.Top + LblPos.Height + 2
                LblPos.Caption = Info
                Screen.MousePointer = vbCrosshair
            Else                               '如果是移动状态
                Screen.MousePointer = 5
                Shape1.Left = OriginalX - (NewX - X)
                Shape1.Top = OriginalY - (NewY - Y)
                If Shape1.Left < 0 Then Shape1.Left = 0   '使区域不超过屏幕
                If Shape1.Top < 0 Then Shape1.Top = 0
                If Shape1.Left + Shape1.Width > Screen.Width / 15 Then Shape1.Left = Screen.Width / 15 - Shape1.Width
                If Shape1.Top + Shape1.Height > Screen.Height / 15 Then Shape1.Top = Screen.Height / 15 - Shape1.Height
                LblPos.Move Shape1.Left + 2, Shape1.Top + 2
                lblInfo.Move Shape1.Left + 2, LblPos.Top + LblPos.Height + 2
            End If
            Call ImageMove
            Shape1.Visible = True
            LblPos.Visible = True
            lblInfo.Visible = True
        End If
    Else
        Dim X2 As Single, y2 As Single
        If START = True Then     '开始编辑
            Me.Cls
            If X > Shape1.Left + Shape1.Width Then
                X2 = Shape1.Left + Shape1.Width
            ElseIf X < Shape1.Left Then
                X2 = Shape1.Left
            Else
                X2 = X
            End If
            If Y > Shape1.Top + Shape1.Height Then
                y2 = Shape1.Top + Shape1.Height
            ElseIf Y < Shape1.Top Then
                y2 = Shape1.Top
            Else
                y2 = Y
            End If
            Select Case EditStr
            Case "添加箭头"
             Call Arrow(Me, X0, Y0, X2, y2, 10, Pcolor(16).BackColor)
            Case "添加文字"
            Case "添加椎圆"
            'Call MoveCircle(Me, x0, y0, X2, Y2, Pcolor(16).BackColor)
                Dim tmppen As Long
                Dim pen As Long
                pen = CreatePen(0, 1, Pcolor(16).BackColor)  '创建一个画笔
                tmppen = SelectObject(Me.hdc, pen)      '选定一个刷子
                Ellipse Me.hdc, X0, Y0, X2, y2       '画图
                SelectObject Me.hdc, tmppen    '删除对象
                DeleteObject pen
            Case "添加矩形"
                Me.Line (X0, Y0)-(X2, y2), Pcolor(16).BackColor, B     '画矩形
            End Select
        End If
    End If
End Sub

Private Sub Form_DblClick()
    If PtInRect(rc, NewX, NewY) Then     '看是否在区域内
        Picture1.Visible = False         '如果选区包含部分提示图片，则需要把图片先隐藏.
        Sleep 10                         '有时候没有这两句会使得shape1也显示在截取的区域里
        DoEvents
        Shape1.Visible = False
        ScrnCap Shape1.Left, Shape1.Top, Shape1.Left + Shape1.Width, Shape1.Top + Shape1.Height
        'MsgBox "图象已经保存到剪贴板中", vbInformation, "提示"
        Unload Me
    End If
End Sub
Public Sub ScrnCap(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim I As Integer
    Shape1.Visible = False               '不需要拷贝shape
    LblPos.Visible = False
    lblInfo.Visible = False
    For I = 0 To 7
        DSB(I).Visible = False
    Next
    DoEvents
    Dim rWidth As Long
    Dim rHeight As Long
    Dim SourceDC As Long
    Dim DestDC As Long
    Dim BHandle As Long
    Dim Wnd As Long
    Dim DHandle As Long
    rWidth = Right - Left
    rHeight = Bottom - Top
    SourceDC = CreateDC("DISPLAY", 0, 0, 0)
    DestDC = CreateCompatibleDC(SourceDC)
    BHandle = CreateCompatibleBitmap(SourceDC, rWidth, rHeight)
    SelectObject DestDC, BHandle
    BitBlt DestDC, 0, 0, rWidth, rHeight, SourceDC, Left, Top, &HCC0020
    Wnd = GetDesktopWindow
    OpenClipboard Wnd
    EmptyClipboard
    SetClipboardData 2, BHandle
    CloseClipboard
    DeleteDC DestDC
    ReleaseDC DHandle, SourceDC
End Sub

Public Sub MDown(X As Single, Y As Single)

End Sub

'--------------------保存截图-----------------------------------------
Public Function CutdSave()
    Dim sFile As String
    Dim SaveOpen As OPENFILENAME
    Dim PicType As String
    SaveOpen.lStructSize = Len(SaveOpen)
    SaveOpen.hwndOwner = 0&
    SaveOpen.lpstrFile = String$(255, 0)
    SaveOpen.nMaxFile = 255
    SaveOpen.lpstrInitialDir = App.Path
    SaveOpen.lpstrFilter = "PNG文件(*.PNG)" + Chr$(0) + "*.PNG" + Chr$(0) + "JPEG文件(*.jpg;*.jpeg)" + Chr$(0) + "*.jpg" + Chr$(0) + "位图文件(*.Bmp)" + Chr$(0) + "*.Bmp" + Chr$(0) + "GIF文件(*.gif)" + Chr$(0) + "*.gif" + Chr$(0) + "TIFF文件(*.TIFF)" + Chr$(0) + "*.tiff" + Chr$(0) + "所有文件(*.*)" + Chr$(0) + "*.*" + Chr$(0)
    SaveOpen.lpstrTitle = "保存为"
    SaveOpen.nFilterIndex = 1    '设置默认选择扩展类型
    SaveOpen.lpstrDefExt = "PNG"   '初始化扩展名
    'SaveOpen.lpstrFile = FileName    '保存文件名称
    If GetSaveFileName(SaveOpen) <> 0 Then
        sFile = Left(SaveOpen.lpstrFile, InStr(SaveOpen.lpstrFile, Chr$(0)) - 1)
    Else
        Exit Function
    End If
    'SavePicture Clipboard.GetData(), sFile
    PicType = Right(sFile, Len(sFile) - InStrRev(sFile, "."))
    SavePic Clipboard.GetData(), sFile, PicType
    Clipboard.Clear   ' 清除剪贴板
End Function

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
For I = 0 To 7
Shape3(I).Visible = False
Next
Shape3(Index).Visible = True
End Sub


Private Sub Pcolor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pcolor(16).BackColor = Pcolor(Index).BackColor
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
For I = 0 To 7
Shape3(I).Visible = False
Next
End Sub
Public Function ImageMove()
    Dim I As Integer
    DSB(0).Move Shape1.Left - (DSB(0).Width / 2), Shape1.Top - (DSB(0).Height / 2)
    DSB(1).Move (Shape1.Left + (Shape1.Width / 2)) - (DSB(0).Width / 2), Shape1.Top - (DSB(0).Height / 2)
    DSB(2).Move (Shape1.Left + (Shape1.Width)) - (DSB(2).Width / 1.5), Shape1.Top - (DSB(2).Height / 2)
    DSB(3).Move Shape1.Left - (DSB(3).Width / 2), Shape1.Top + (Shape1.Height / 2) - (DSB(3).Height / 2)
    DSB(4).Move Shape1.Left - (DSB(4).Width / 2), Shape1.Top + (Shape1.Height) - (DSB(4).Height / 2)
    DSB(5).Move (Shape1.Left + (Shape1.Width / 2)) - (DSB(5).Width / 2), Shape1.Top + (Shape1.Height) - (DSB(5).Height / 2)
    DSB(6).Move (Shape1.Left + (Shape1.Width)) - (DSB(6).Width / 2), Shape1.Top + (Shape1.Height) - (DSB(6).Height / 2)
    DSB(7).Move (Shape1.Left + (Shape1.Width)) - (DSB(7).Width / 2), Shape1.Top + (Shape1.Height / 2) - (DSB(7).Height / 2)
    For I = 0 To 7
        DSB(I).Visible = True
    Next
End Function

'-----------------------------增加---------------------------------
Private Sub PicCombox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicCombox.PICTURE = Image4(2).PICTURE
OpenCombo Combo1.hwnd
End Sub

Private Sub PicCombox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicCombox.PICTURE = Image4(1).PICTURE
End Sub


Private Sub PicFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicCombox.PICTURE = Image4(0).PICTURE
End Sub

Private Sub Combo1_Click()
    Text1.Text = Combo1.Text
End Sub

Private Sub Pcolor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> 16 Then
Pcolor(Index).BorderStyle = 1
Else
For I = 0 To 15
Pcolor(I).BorderStyle = 0
Next
End If
End Sub

Private Sub PicColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For I = 0 To 15
Pcolor(I).BorderStyle = 0
Next
End Sub

Public Sub OpenCombo(chwnd As Long)   '强制弹出Combo1的下拉列表
   Dim rc As Long
    rc = SendMessage(chwnd, CB_GETDROPPEDSTATE, 0, 0)
    If rc = 0 Then
        SendMessage chwnd, CB_SHOWDROPDOWN, True, 0
    Else
        SendMessage chwnd, CB_SHOWDROPDOWN, False, 0
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = 1
End Sub


'-------------------画箭头---------------------
Sub Arrow(pic As Object, X0 As Single, Y0 As Single, X1 As Single, y1 As Single, ArrowLen As Single, Optional Color As Long = 0)
Dim Xa As Single, Ya As Single, Xb As Single, Yb As Single, d As Double
d = Sqr((y1 - Y0) * (y1 - Y0) + (X1 - X0) * (X1 - X0))
If d > 0.0000000001 Then
Xa = X1 + ArrowLen * ((X0 - X1) + (Y0 - y1) / 2) / d
Ya = y1 + ArrowLen * ((Y0 - y1) - (X0 - X1) / 2) / d
Xb = X1 + ArrowLen * ((X0 - X1) - (Y0 - y1) / 2) / d
Yb = y1 + ArrowLen * ((Y0 - y1) + (X0 - X1) / 2) / d
pic.Line (Xa, Ya)-(X1, y1), Color
pic.Line (Xb, Yb)-(X1, y1), Color
pic.Line (X0, Y0)-(X1, y1), Color '如果仅画箭头，此句可删除
End If
End Sub

Private Sub TextEdit_Change()
If Me.TextWidth(TextEdit) > TextEdit.Width Then TextEdit.Width = Me.TextWidth(TextEdit)
End Sub
Public Function WriteP(ByRef pc() As Byte, pic As StdPicture)
'VB鼠标画圆并用数组保存每次操作，做到撤销，重做
Dim pbag As New PropertyBag
pbag.WriteProperty "pic", pic
pc = pbag.Contents
End Function
Public Function ReadP(ByRef rc() As Byte)
Dim pbagb As New PropertyBag
pbagb.Contents = rc
Set ReadP = pbagb.ReadProperty("pic")
End Function

Public Sub SavePic(ByVal pict As StdPicture, ByVal filename As String, PicType As String, _
Optional ByVal Quality As Byte = 80, _
Optional ByVal TIFF_ColorDepth As Long = 24, _
Optional ByVal TIFF_Compression As Long = 6)
Screen.MousePointer = vbHourglass
Dim tSI As GdiplusStartupInput
Dim lRes As Long
Dim lGDIP As Long
Dim lBitmap As Long
Dim aEncParams() As Byte
On Error GoTo ErrHandle:
tSI.GdiplusVersion = 1 ' 初始化 GDI+
lRes = GdiplusStartup(lGDIP, tSI)
If lRes = 0 Then ' 从句柄创建 GDI+ 图像
lRes = GdipCreateBitmapFromHBITMAP(pict.handle, 0, lBitmap)
If lRes = 0 Then
Dim tJpgEncoder As GUID
Dim tParams As EncoderParameters '初始化解码器的GUID标识
Select Case LCase(PicType)
Case "jpg"
CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
tParams.Count = 1 ' 设置解码器参数
With tParams.Parameter ' Quality
CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID ' 得到Quality参数的GUID标识
.NumberOfValues = 1
.type = 4
.Value = VarPtr(Quality)
End With
ReDim aEncParams(1 To Len(tParams))
Call CopyMemory(aEncParams(1), tParams, Len(tParams))
Case "png"
CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
ReDim aEncParams(1 To Len(tParams))
Case "gif"
CLSIDFromString StrPtr("{557CF402-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
ReDim aEncParams(1 To Len(tParams))
Case "tiff"
CLSIDFromString StrPtr("{557CF405-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
tParams.Count = 2
ReDim aEncParams(1 To Len(tParams) + Len(tParams.Parameter))
With tParams.Parameter
.NumberOfValues = 1
.type = 4
CLSIDFromString StrPtr("{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"), .GUID ' 得到ColorDepth参数的GUID标识
.Value = VarPtr(TIFF_Compression)
End With
Call CopyMemory(aEncParams(1), tParams, Len(tParams))
With tParams.Parameter
.NumberOfValues = 1
.type = 4
CLSIDFromString StrPtr("{66087055-AD66-4C7C-9A18-38A2310B8337}"), .GUID ' 得到Compression参数的GUID标识
.Value = VarPtr(TIFF_ColorDepth)
End With
Call CopyMemory(aEncParams(Len(tParams) + 1), tParams.Parameter, Len(tParams.Parameter))
Case "bmp" '可以提前写保存为BMP的代码，因为并没有用GDI+
SavePicture pict, filename
Screen.MousePointer = vbDefault
Exit Sub
End Select
lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, aEncParams(1)) '保存图像
GdipDisposeImage lBitmap ' 销毁GDI+图像
End If
GdiplusShutdown lGDIP '销毁 GDI+
End If
Screen.MousePointer = vbDefault
Erase aEncParams
Exit Sub
ErrHandle:
Screen.MousePointer = vbDefault
End Sub


