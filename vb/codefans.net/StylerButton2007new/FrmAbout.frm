VERSION 5.00
Begin VB.Form FrmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   39.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00757B48&
   LinkTopic       =   "Form1"
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin StylerButton2007.StylerButton V 
      Height          =   600
      Left            =   6180
      TabIndex        =   0
      Top             =   3750
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   582
      _ExtentY        =   212
      Caption         =   "OK"
      ForeColor       =   9404976
      CaptionDisableColor=   12236471
      CaptionEffectColor=   15393985
      CaptionEffect   =   4
      IconDisableColor=   11711154
      Theme           =   3
      FocusDottedRect =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBLC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1125
      Left            =   1890
      TabIndex        =   3
      Top             =   5640
      Width           =   270
   End
   Begin VB.Label LBLB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   -195
      TabIndex        =   2
      Top             =   7170
      Width           =   240
   End
   Begin VB.Label LBLA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   -30
      TabIndex        =   1
      Top             =   4350
      Width           =   105
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Const LWA_ALPHA = &H2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Sub TOPFORM(hWnd As Long, Action As Boolean)
    If Action = True Then
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    LBLB.Left = Me.ScaleWidth / 2 - LBLB.Width / 2 - 1
    LBLB.Top = 100

    lonRect = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 20, 20)
    SetWindowRgn Me.hWnd, lonRect, True
    V.DrawGradientFourColour Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight / 2 - 1, RGB(255, 255, 255), RGB(255, 255, 255), RGB(168, 208, 229), RGB(168, 208, 229)
    V.DrawGradientFourColour Me.hDC, 0, Me.ScaleHeight / 2 - 1, Me.ScaleWidth, Me.ScaleHeight / 2, RGB(146, 193, 217), RGB(146, 193, 217), RGB(213, 236, 247), RGB(213, 236, 247)
    RoundRectBorder Me, 0, 0, Me.ScaleWidth - 1, Me.ScaleHeight - 1, 20, 20, vbWhite
    FormFadeIn Me, 0, 240, 4

    TA
End Sub

Private Sub RoundRectBorder(nObject As Object, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, nColor As ColorConstants)
    Dim A As Variant
    A = nObject.ForeColor
    nObject.ForeColor = nColor
    RoundRect nObject.hDC, X1, Y1, X2, Y2, X3, Y3
    nObject.ForeColor = A
End Sub

Private Sub FormFadeIn(ByRef nForm As Form, Optional ByVal nFadeStart As Byte = 0, Optional ByVal nFadeEnd As Byte = 255, Optional ByVal nFadeInSpeed As Byte = 5)
    Dim c
    Dim ne As Integer, en(32767) As Boolean
    For Each c In nForm.Controls
        ne = ne + 1
        en(ne) = c.Enabled
        c.Enabled = False
    Next
    If nFadeEnd = 0 Then
        nFadeEnd = 255
    End If
    If nFadeInSpeed = 0 Then
        nFadeInSpeed = 5
    End If
    If nFadeStart >= nFadeEnd Then
        nFadeStart = 0
    ElseIf nFadeEnd <= nFadeStart Then
        nFadeEnd = 255
    End If

    TransparentsForm nForm.hWnd, 0
    nForm.Show
    Dim i As Long
    For i = nFadeStart To nFadeEnd Step nFadeInSpeed
        TransparentsForm nForm.hWnd, CByte(i)
        DoEvents
        Call Sleep(5)
    Next
    TransparentsForm nForm.hWnd, nFadeEnd
    i = 0
    For Each c In nForm.Controls
        i = i + 1
        c.Enabled = en(i)
    Next
End Sub

Private Function TransparentsForm(FormhWnd As Long, Alpha As Byte) As Boolean
    SetWindowLong FormhWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes FormhWnd, 0, Alpha, LWA_ALPHA
    LastAlpha = Alpha
End Function

Private Sub FormFadeOut(ByRef nForm As Form)
    On Local Error Resume Next
    Dim c
    Dim S As Integer
    For Each c In nForm.Controls
        c.Enabled = False
    Next

    Dim i As Long
    For i = 240 To 0 Step -5
        TransparentsForm nForm.hWnd, CByte(i)
        DoEvents
        Call Sleep(5)
    Next

End Sub

Private Sub TA()
    On Local Error Resume Next

    LA
    '==============================

    Me.FontSize = 40
    Me.FontBold = True
    Me.FontName = "Georgia"
    Me.ForeColor = RGB(72, 123, 117)

    For i = 0 To 18
        Me.CurrentX = 10
        Me.CurrentY = 10
        Me.Print Mid$("Styler Button 2007", 1, CByte(i))
        DoEvents
        Call Sleep(50)
    Next

    Me.FontBold = False
    Me.FontSize = 8
    Me.CurrentX = 520
    Me.CurrentY = 17
    Me.Print "?"

    Me.ForeColor = 0
    Me.FontSize = 8
    Me.FontName = "Times New Roman"

    Me.CurrentX = 437
    Me.CurrentY = 58

    Me.Print "VER " & App.Major & "." & App.Minor & "." & App.Revision

    Me.FontBold = True
    Me.ForeColor = 0
    Me.FontSize = 15
    Me.FontName = "Georgia"

    For i = 0 To 13
        Me.CurrentX = 10
        Me.CurrentY = 70
        Me.Print Mid$("Developed BY:", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next

    Call Sleep(200)

    Me.ForeColor = vbRed

    For i = 0 To 9
        Me.CurrentX = 165
        Me.CurrentY = 70
        Me.Print Mid$("UMAIR_11D", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next

    Call Sleep(100)

    Me.ForeColor = vbBlack
    Me.FontSize = 10
    For i = 0 To 63
        Me.CurrentX = 12
        Me.CurrentY = 100
        Me.Print Mid$("Styler Button 2007 is Very Quick, Powerful & New styles Botton.", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next

    Call Sleep(100)

    For i = 0 To 67
        Me.CurrentX = 12
        Me.CurrentY = 120
        Me.Print Mid$("Styler Button 2007 enables you to customize the appearance of your", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next
    Call Sleep(100)

    For i = 0 To 44
        Me.CurrentX = 12
        Me.CurrentY = 140
        Me.Print Mid$("applications to suit your individual needs.", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next

    Call Sleep(100)

    For i = 0 To 30
        Me.CurrentX = 12
        Me.CurrentY = 180
        Me.Print Mid$("If You Find Any Problems/Bug.", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next
    Call Sleep(100)

    For i = 0 To 32
        Me.CurrentX = 12
        Me.CurrentY = 200
        Me.Print Mid$("Any Questions For This Project.", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next

    Call Sleep(100)

    For i = 0 To 26
        Me.CurrentX = 12
        Me.CurrentY = 220
        Me.Print Mid$("Email:Umair_11D@Yahoo.com", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next
    Call Sleep(100)

    For i = 0 To 41
        Me.CurrentX = 12
        Me.CurrentY = 240
        Me.Print Mid$("Voice NO.:+923453021375 , +923002242573", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next

    Call Sleep(100)
    Me.FontBold = False
    Me.FontSize = 8
    For i = 0 To 32
        Me.CurrentX = 12
        Me.CurrentY = 280
        Me.Print Mid$("Copyright ?2004-2007 UMAIR_11D.", 1, CByte(i))
        DoEvents
        Call Sleep(10)
    Next

    V.Visible = True
    V.SetFocus
End Sub

Private Sub V_Click()
    FormFadeOut Me
    TOPFORM Me.hWnd, False
    Unload Me
End Sub

Private Sub LA()
    LBLA.Top = 10
    LBLA.Left = 10
    LBLB.Left = (Me.ScaleWidth / 2) - 284 / 2
    LBLB.Top = 100
    LBLC.Top = 105
    LBLC.Left = 142

    For i = 0 To 15
        LBLA.Caption = Mid$("正在初始化中...", 1, CByte(i))
        DoEvents
        Sleep (40)
    Next i

    Sleep (100)
    MA

    LBLB.Caption = ""
    LBLC.Caption = "    确定"

    Sleep (100)

    LBLC.Caption = ""
    For i = 0 To 15
        LBLA.Caption = Mid$("Loading...", 1, CByte(i))
        DoEvents
        Sleep (40)
    Next i
    Sleep (100)
    MA
    LBLB.Caption = ""
    LBLC.Caption = "    OK"

    Sleep (100)

    LBLC.Caption = ""
    LBLA.Caption = ""

End Sub

Private Sub MA()
    For i = 0 To 12
        LBLB.Caption = Mid$("[__________]", 1, CByte(i))
        DoEvents
        Sleep (10)
    Next i

    q = True

    For A = 0 To 5
        If q = True Then
            q = False
        ElseIf q = False Then
            q = True
            qq = 13
        End If
        For i = 0 To 12
            If q = True Then
                qq = qq - 1
                LBLC.Caption = Space$(qq) & "*"
            ElseIf q = False Then
                LBLC.Caption = Space$(i) & "*"
            End If

            DoEvents
            Sleep (5)
        Next i
    Next A
End Sub
