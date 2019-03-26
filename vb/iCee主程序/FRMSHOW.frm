VERSION 5.00
Begin VB.Form FRMSHOW 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "歌词秀"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   903
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "FRMSHOW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Private Const ULW_OPAQUE = &H4
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER = &H0
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type Size
    cx As Long
    cy As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
'
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type GdiFont
    FontName        As String
    FontSize        As Single
    FontBold        As Boolean
    BackColor1      As OLE_COLOR            '歌词颜色
    BackColor2      As OLE_COLOR
    ForeColor1      As OLE_COLOR            '卡拉OK字幕颜色
    ForeColor2      As OLE_COLOR
    LineColor       As OLE_COLOR            '描边线条颜色
End Type
'
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Dim mDC             As Long
Dim tempBI          As BITMAPINFO
Dim mainBitmap      As Long
Dim graphics        As Long, brush      As Long, pen       As Long
Dim fontFam         As Long, strFormat  As Long  '字体相关
Dim strpath         As Long
Dim rclayout        As RECTL
Dim MyRect          As RECTF
Dim pos As POINTAPI '定义这个变量是取得鼠标坐标

Dim blendFunc32bpp  As BLENDFUNCTION            '混合位图功能

Private strLastText         As String               '最近画边的文件 //方便重绘
Private iLastWidth          As Long
Private m_OK                As Boolean              'GDI+是否初始化成功
Private m_Font              As GdiFont              '文字及颜色信息
Private m_FontStyle         As Long                 '是否是粗体
Private m_FontSize          As Single               '换算成像素的尺寸

Private Sub Form_Load()
D_L_SHOW = True
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    '调整窗体大小及位置
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, (Screen.Height - GetTaskbarHeight) / Screen.TwipsPerPixelY - 70, _
                        Screen.Width / Screen.TwipsPerPixelX, 70, 0  ' SWP_NOMOVE Or SWP_NOSIZE
 Call TrForm(Me)
    '更新窗口样式
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED Or GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    m_OK = GDIPlusInitialize              '初始化 GDI+
    If Not m_OK Then Exit Sub
    
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32            '色深32位
        .biHeight = Me.ScaleHeight
        .biWidth = Me.ScaleWidth
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
    End With
    
    mDC = CreateCompatibleDC(Me.hdc)
    GdipCreateStringFormat 0, 0, strFormat                          '创建字体样式
    GdipSetStringFormatAlign strFormat, StringAlignmentCenter       '设置字体样式

    '初始化文字信息
    With m_Font
        .FontName = "微软雅黑"
        .FontSize = 18
        .LineColor = &H30000000
    End With
    m_FontSize = ConvertPointsToPixels(32)
    
    GdipCreateFontFamilyFromName StrPtr(m_Font.FontName), 0, fontFam '创建一个字体家族 2指定该字体属于哪个字体集（如果有），如果没有，则为NULL 3指向生成的字体家族的指针
    MyRect.Height = m_FontSize
    MyRect.Width = m_FontSize
    MyRect.Top = 0
    MyRect.Left = 0
    
    rclayout.Right = Screen.Width / 15
    rclayout.Bottom = 0
    
    
End Sub

Public Sub ReDrawText()
    DrawText strLastText, iLastWidth
End Sub


Public Sub DrawText(ByVal Text As String, Optional ByVal ShowWidth As Long = 0)
    'ShowWidth:         卡拉OK唱过的宽度
    If Not m_OK Then Exit Sub                                       'GDI+ 初始化失败，不绘制
    If Len(Text) = 0 And Len(strLastText) = 0 Then Exit Sub         '空内容，不绘制
    
    strLastText = Text
    iLastWidth = ShowWidth
    
    '每次重新创建 mainBitmap，不然之前的内容不会消失
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    DeleteObject SelectObject(mDC, mainBitmap)                      '删除DC中原来位图
    
    Call GdipCreateFromHDC(mDC, graphics)                           '功能:创建设备场景相对应的绘图区域（相当于给设备场景创建一个画板）
                                                                    'graphics  我们要创建的画板，创建成功后的画板的句柄存放在此
    GdipSetSmoothingMode graphics, SmoothingModeHighQuality         '消除锯齿
    
    '画 N 层阴影
    '***********************************
    '***********************************
    rclayout.Left = 1
    rclayout.Top = 1
    Call GdipCreateLineBrushFromRect(MyRect, &H90000000, &H90000000, LinearGradientModeVertical, WrapModeTileFlipY, brush) '创建一个渐变填充笔刷
    GdipCreatePath FillModeAlternate, strpath                       '创建一个路径
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
    GdipFillPath graphics, brush, strpath                           '填充路径
    GdipDeletePath strpath                                          '删除路径，为下一次填充做准备
'    '***********************************
'    '***********************************
    rclayout.Left = 1.5                                             '换位置 / 颜色
    rclayout.Top = 1.5
    Call GdipCreateLineBrushFromRect(MyRect, &H30000000, &H30000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
'    '***********************************
'    '***********************************
    rclayout.Left = 2
    rclayout.Top = 2
    Call GdipCreateLineBrushFromRect(MyRect, &H20000000, &H20000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
'    '***********************************
'    '***********************************
    rclayout.Left = 2.2
    rclayout.Top = 2.2
    Call GdipCreateLineBrushFromRect(MyRect, &H10000000, &H10000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
    '***********************************
    '***********************************
    rclayout.Left = -1
    rclayout.Top = -1
    Call GdipCreateLineBrushFromRect(MyRect, &H40000000, &H40000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
    '***********************************
    '***********************************
    rclayout.Left = -1.5
    rclayout.Top = -1.5
    Call GdipCreateLineBrushFromRect(MyRect, &H30000000, &H30000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath       '创建填充路径
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
    '#######################
    '#######################
    '--最后，画过渡文字------
    '#######################
    rclayout.Left = 0
    rclayout.Top = 0
    Call GdipCreateLineBrushFromRect(MyRect, m_Font.BackColor1, m_Font.BackColor2, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
    GdipFillPath graphics, brush, strpath
    
    If ShowWidth > 0 Then
        GdipDeletePath strpath
        '#######################
        '--画 唱过的歌词--------
        '#######################
        Call GdipCreateLineBrushFromRect(MyRect, m_Font.ForeColor1, m_Font.ForeColor2, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
        GdipCreatePath FillModeAlternate, strpath
        Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '添加文字
        '设置区域
        ShowWidth = (Me.Width / Screen.TwipsPerPixelX - Me.TextWidth(Text)) / 2 + ShowWidth
        GdipSetClipRectI graphics, 0, 0, ShowWidth, 50, CombineModeReplace        '设定剪辑区域，以达到类似卡拉OK效果
            GdipFillPath graphics, brush, strpath
        GdipResetClip graphics                                              '取消剪辑区域
    End If
    
    GdipCreatePen1 m_Font.LineColor, 1, UnitDocument, pen                  '创建一个描边的笔刷
    GdipDrawPath graphics, pen, strpath                              '文字描边
    GdipDeletePen pen
    
    GdipDeletePath strpath
    DeleteObject mainBitmap
    GdipDeleteGraphics graphics
    '===========================
    '===========================
    '===========================
    '更新分层的窗口的位置，大小，形状，内容和半透明度
    Dim winSize As Size
    winSize.cx = Me.ScaleWidth:  winSize.cy = Me.ScaleHeight
    Call UpdateLayeredWindow(Me.hwnd, Me.hdc, ByVal 0&, winSize, mDC, 0@, 0, blendFunc32bpp, ULW_ALPHA)
    'pptSrc不是一个可选的参数，要传POINT(0,0)，因此我们可以使用一个(Currency)0来填充这个参数指向的8个字节同时保证内容为0.
End Sub

Private Sub Form_Unload(Cancel As Integer)
D_L_SHOW = False
Call ClearGDI
End Sub

Public Sub ClearGDI()
    m_OK = False
    'GDI+ 完成
    GdipDeleteGraphics graphics             '删除画板
    GdipDeleteFontFamily fontFam            '删除字体样式
   
    DeleteObject mainBitmap                 '删除
    DeleteDC mDC
    
    GdipDeleteStringFormat strFormat
    GdipDeletePath strpath
    GdipDeleteBrush brush
    
    GDIPlusTerminate                        '析构 GDI +
    Debug.Print "析构GDI+"
End Sub


Public Property Get LrcFontName() As String
    LrcFontName = m_Font.FontName
End Property

Public Property Let LrcFontName(ByVal NewName As String)
    m_Font.FontName = NewName
    Me.FontName = NewName
    '改变了字体，重新创建字体家族
    GdipDeleteFontFamily fontFam
    GdipCreateFontFamilyFromName StrPtr(m_Font.FontName), 0, fontFam '创建一个字体家族 2指定该字体属于哪个字体集（如果有），如果没有，则为NULL 3指向生成的字体家族的指针
End Property

Public Property Get LrcFontSize() As Single
    LrcFontSize = m_Font.FontSize
End Property

Public Property Let LrcFontSize(ByVal NewSize As Single)
    m_Font.FontSize = NewSize
    m_FontSize = ConvertPointsToPixels(NewSize)         '由磅转换为像素
    MyRect.Height = m_FontSize
    MyRect.Width = m_FontSize
    
    Me.FontSize = NewSize
End Property

Public Property Get LrcFontBold() As Boolean
    LrcFontBold = m_Font.FontBold
End Property

Public Property Let LrcFontBold(ByVal Bold As Boolean)
    m_Font.FontBold = Bold
    m_FontStyle = IIf(Bold, FontStyle.FontStyleBold, 0)
    Me.FontBold = Bold
End Property

Public Property Get BackColor1() As OLE_COLOR
    BackColor1 = m_Font.BackColor1
End Property

Public Property Let BackColor1(ByVal NewColor As OLE_COLOR)
    m_Font.BackColor1 = NewColor
End Property

Public Property Get BackColor2() As OLE_COLOR
    BackColor2 = m_Font.BackColor2
End Property

Public Property Let BackColor2(ByVal NewColor As OLE_COLOR)
    m_Font.BackColor2 = NewColor
End Property

Public Property Get ForeColor1() As OLE_COLOR
    ForeColor1 = m_Font.ForeColor1
End Property

Public Property Let ForeColor1(ByVal NewColor As OLE_COLOR)
    m_Font.ForeColor1 = NewColor
End Property

Public Property Get ForeColor2() As OLE_COLOR
    ForeColor2 = m_Font.ForeColor2
End Property

Public Property Let ForeColor2(ByVal NewColor As OLE_COLOR)
    m_Font.ForeColor2 = NewColor
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = m_Font.LineColor
End Property

Public Property Let LineColor(ByVal NewColor As OLE_COLOR)
    m_Font.LineColor = NewColor
End Property

