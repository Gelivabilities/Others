VERSION 5.00
Begin VB.Form FRMSHOW 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "�����"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   903
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
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
    BackColor1      As OLE_COLOR            '�����ɫ
    BackColor2      As OLE_COLOR
    ForeColor1      As OLE_COLOR            '����OK��Ļ��ɫ
    ForeColor2      As OLE_COLOR
    LineColor       As OLE_COLOR            '���������ɫ
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
Dim fontFam         As Long, strFormat  As Long  '�������
Dim strpath         As Long
Dim rclayout        As RECTL
Dim MyRect          As RECTF
Dim pos As POINTAPI '�������������ȡ���������

Dim blendFunc32bpp  As BLENDFUNCTION            '���λͼ����

Private strLastText         As String               '������ߵ��ļ� //�����ػ�
Private iLastWidth          As Long
Private m_OK                As Boolean              'GDI+�Ƿ��ʼ���ɹ�
Private m_Font              As GdiFont              '���ּ���ɫ��Ϣ
Private m_FontStyle         As Long                 '�Ƿ��Ǵ���
Private m_FontSize          As Single               '��������صĳߴ�

Private Sub Form_Load()
D_L_SHOW = True
RESL = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    '���������С��λ��
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, (Screen.Height - GetTaskbarHeight) / Screen.TwipsPerPixelY - 70, _
                        Screen.Width / Screen.TwipsPerPixelX, 70, 0  ' SWP_NOMOVE Or SWP_NOSIZE
 Call TrForm(Me)
    '���´�����ʽ
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED Or GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    m_OK = GDIPlusInitialize              '��ʼ�� GDI+
    If Not m_OK Then Exit Sub
    
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32            'ɫ��32λ
        .biHeight = Me.ScaleHeight
        .biWidth = Me.ScaleWidth
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
    End With
    
    mDC = CreateCompatibleDC(Me.hdc)
    GdipCreateStringFormat 0, 0, strFormat                          '����������ʽ
    GdipSetStringFormatAlign strFormat, StringAlignmentCenter       '����������ʽ

    '��ʼ��������Ϣ
    With m_Font
        .FontName = "΢���ź�"
        .FontSize = 18
        .LineColor = &H30000000
    End With
    m_FontSize = ConvertPointsToPixels(32)
    
    GdipCreateFontFamilyFromName StrPtr(m_Font.FontName), 0, fontFam '����һ��������� 2ָ�������������ĸ����弯������У������û�У���ΪNULL 3ָ�����ɵ���������ָ��
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
    'ShowWidth:         ����OK�����Ŀ��
    If Not m_OK Then Exit Sub                                       'GDI+ ��ʼ��ʧ�ܣ�������
    If Len(Text) = 0 And Len(strLastText) = 0 Then Exit Sub         '�����ݣ�������
    
    strLastText = Text
    iLastWidth = ShowWidth
    
    'ÿ�����´��� mainBitmap����Ȼ֮ǰ�����ݲ�����ʧ
    mainBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    DeleteObject SelectObject(mDC, mainBitmap)                      'ɾ��DC��ԭ��λͼ
    
    Call GdipCreateFromHDC(mDC, graphics)                           '����:�����豸�������Ӧ�Ļ�ͼ�����൱�ڸ��豸��������һ�����壩
                                                                    'graphics  ����Ҫ�����Ļ��壬�����ɹ���Ļ���ľ������ڴ�
    GdipSetSmoothingMode graphics, SmoothingModeHighQuality         '�������
    
    '�� N ����Ӱ
    '***********************************
    '***********************************
    rclayout.Left = 1
    rclayout.Top = 1
    Call GdipCreateLineBrushFromRect(MyRect, &H90000000, &H90000000, LinearGradientModeVertical, WrapModeTileFlipY, brush) '����һ����������ˢ
    GdipCreatePath FillModeAlternate, strpath                       '����һ��·��
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
    GdipFillPath graphics, brush, strpath                           '���·��
    GdipDeletePath strpath                                          'ɾ��·����Ϊ��һ�������׼��
'    '***********************************
'    '***********************************
    rclayout.Left = 1.5                                             '��λ�� / ��ɫ
    rclayout.Top = 1.5
    Call GdipCreateLineBrushFromRect(MyRect, &H30000000, &H30000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
'    '***********************************
'    '***********************************
    rclayout.Left = 2
    rclayout.Top = 2
    Call GdipCreateLineBrushFromRect(MyRect, &H20000000, &H20000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
'    '***********************************
'    '***********************************
    rclayout.Left = 2.2
    rclayout.Top = 2.2
    Call GdipCreateLineBrushFromRect(MyRect, &H10000000, &H10000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
    '***********************************
    '***********************************
    rclayout.Left = -1
    rclayout.Top = -1
    Call GdipCreateLineBrushFromRect(MyRect, &H40000000, &H40000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
    '***********************************
    '***********************************
    rclayout.Left = -1.5
    rclayout.Top = -1.5
    Call GdipCreateLineBrushFromRect(MyRect, &H30000000, &H30000000, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath       '�������·��
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
    GdipFillPath graphics, brush, strpath
    GdipDeletePath strpath
    '#######################
    '#######################
    '--��󣬻���������------
    '#######################
    rclayout.Left = 0
    rclayout.Top = 0
    Call GdipCreateLineBrushFromRect(MyRect, m_Font.BackColor1, m_Font.BackColor2, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
    GdipCreatePath FillModeAlternate, strpath
    Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
    GdipFillPath graphics, brush, strpath
    
    If ShowWidth > 0 Then
        GdipDeletePath strpath
        '#######################
        '--�� �����ĸ��--------
        '#######################
        Call GdipCreateLineBrushFromRect(MyRect, m_Font.ForeColor1, m_Font.ForeColor2, LinearGradientModeVertical, WrapModeTileFlipXY, brush)
        GdipCreatePath FillModeAlternate, strpath
        Call GdipAddPathStringI(strpath, StrPtr(Text), -1, fontFam, m_FontStyle, m_FontSize, rclayout, strFormat)  '�������
        '��������
        ShowWidth = (Me.Width / Screen.TwipsPerPixelX - Me.TextWidth(Text)) / 2 + ShowWidth
        GdipSetClipRectI graphics, 0, 0, ShowWidth, 50, CombineModeReplace        '�趨���������Դﵽ���ƿ���OKЧ��
            GdipFillPath graphics, brush, strpath
        GdipResetClip graphics                                              'ȡ����������
    End If
    
    GdipCreatePen1 m_Font.LineColor, 1, UnitDocument, pen                  '����һ����ߵı�ˢ
    GdipDrawPath graphics, pen, strpath                              '�������
    GdipDeletePen pen
    
    GdipDeletePath strpath
    DeleteObject mainBitmap
    GdipDeleteGraphics graphics
    '===========================
    '===========================
    '===========================
    '���·ֲ�Ĵ��ڵ�λ�ã���С����״�����ݺͰ�͸����
    Dim winSize As Size
    winSize.cx = Me.ScaleWidth:  winSize.cy = Me.ScaleHeight
    Call UpdateLayeredWindow(Me.hwnd, Me.hdc, ByVal 0&, winSize, mDC, 0@, 0, blendFunc32bpp, ULW_ALPHA)
    'pptSrc����һ����ѡ�Ĳ�����Ҫ��POINT(0,0)��������ǿ���ʹ��һ��(Currency)0������������ָ���8���ֽ�ͬʱ��֤����Ϊ0.
End Sub

Private Sub Form_Unload(Cancel As Integer)
D_L_SHOW = False
Call ClearGDI
End Sub

Public Sub ClearGDI()
    m_OK = False
    'GDI+ ���
    GdipDeleteGraphics graphics             'ɾ������
    GdipDeleteFontFamily fontFam            'ɾ��������ʽ
   
    DeleteObject mainBitmap                 'ɾ��
    DeleteDC mDC
    
    GdipDeleteStringFormat strFormat
    GdipDeletePath strpath
    GdipDeleteBrush brush
    
    GDIPlusTerminate                        '���� GDI +
    Debug.Print "����GDI+"
End Sub


Public Property Get LrcFontName() As String
    LrcFontName = m_Font.FontName
End Property

Public Property Let LrcFontName(ByVal NewName As String)
    m_Font.FontName = NewName
    Me.FontName = NewName
    '�ı������壬���´����������
    GdipDeleteFontFamily fontFam
    GdipCreateFontFamilyFromName StrPtr(m_Font.FontName), 0, fontFam '����һ��������� 2ָ�������������ĸ����弯������У������û�У���ΪNULL 3ָ�����ɵ���������ָ��
End Property

Public Property Get LrcFontSize() As Single
    LrcFontSize = m_Font.FontSize
End Property

Public Property Let LrcFontSize(ByVal NewSize As Single)
    m_Font.FontSize = NewSize
    m_FontSize = ConvertPointsToPixels(NewSize)         '�ɰ�ת��Ϊ����
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

