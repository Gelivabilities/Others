Attribute VB_Name = "ͼ����"
'���ģ�����й�ͼ�����ģ��
Option Explicit
Public tR As Integer
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Enum UseAPIPaintPicture
    APIBitBlt = 1
    APISetDIBitsToDevice = 2
    APIStretchBlt = 3
    APIStretchDIBits = 4
End Enum
Public Type Clsid
    Data1         As Long
    data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type
Private Const CBM_INIT = &H4
Public Const Bmp_MAGIC_COOKIE As Integer = 19778
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'���ɶ����Handleȡ�ö������ݽṹ��API����
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Public CurrentEntryIndex As Integer
Public DecIndex As Integer
Public ErrorMessage As String
Public Help As Boolean
Public InError As Boolean
Public InputString As String
Public OutputString As String
Public OutputValue As Double
Public PrevAnswer As Double
Public PrevEntry As String
Public SetVariable As Boolean
Public Value As Double
Public ValueString As String
Public WindowCount As Integer
Public Char As String
Public Const PI = 3.14159265358979
Public MainArray() As Double
Public ValueArray() As Double
Public VariableArray() As String
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
Public Const WS_EX_LAYERED = &H80000
Public Const d_Bg = &HF6F6F6
Public Const d_Border = &H800000
Public Const d_Title1 = &HD68759
Public Const d_Title2 = &H9A400C
Public Const d_Bar1 = &HFAE2D0
Public Const d_Bar2 = &HE2A981
Public Const d_Hl1 = &HD0FCFD
Public Const d_Hl2 = &H9DDFFD
Public Const d_Checked1 = &H7DDDFA
Public Const d_Checked2 = &H4EBCF5
Public Const d_Pressed1 = &H5586F8
Public Const d_Pressed2 = &HA37D2
Public Const d_Sprt1 = &HCB8C6A
Public Const d_Sprt2 = vbWhite
Public Const d_Text = vbBlack
Public Const d_TextHl = vbBlack
Public Const d_TextDis = &HCB8C6A
Public Const d_Chevron1 = &HF1A675
Public Const d_Chevron2 = &H913500
Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type RGBTriple
    Red As Byte
    Green As Byte
    Blue As Byte
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Public Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Const WM_WINDOWPOSCHANGING = &H46
Public Type WINDOWPOS
        hwnd As Long
        hWndInsertAfter As Long
        X As Long  '������λ��X����
        Y As Long  '������λ��Y����
        cx As Long  '������λ�Ŀ��
        cy As Long '������λ�ĸ߶�
        flags As Long
End Type
Public Prewininf As Long
Public Type rBlendProps
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Type TRIVERTEX
   X As Long
   Y As Long
   Red0 As Byte
   Red1 As Byte
   Green0 As Byte
   Green1 As Byte
   Blue0 As Byte
   Blue1 As Byte
   Alpha0 As Byte
   Alpha1 As Byte
End Type
Public Type BmpRGB
    'ͼ��������ɫ��
    Blue As Byte                  '��
    Green As Byte                 '��
    Red As Byte                   '��
    'ͼ���������������ϣ���������
End Type
Public Type BmpFileHeard
    'λͼ�ļ�ͷ
    BmpType As String * 2         'λͼ��־
    BmpFileSize As Long           'λͼ�ļ������ֽ���
    BmpReserved As Long           '�����ֽ�
    BmpOffBits As Long            'λͼ���е���ʼλ��
End Type

Public Type BmpPictureHeard
    'λͼ��Ϣͷ
    BmpFileHeardLong As Long      '��Ϣͷ�ĳ���
    BmpWidth As Long              '��(����)
    BmpHeight As Long             '��(����)
    BmpPlanes As Integer          'λͼ�豸����
    BmpBitCount As Integer        '��ɫ��
    BmpCompression As Long        'ѹ������(0��ʾ��ѹ��)
    BmpSizeImage As Long          'λͼ���б��ֽ���
    BmpXPlesPerMeter As Long      'ˮƽ�ֱ���
    BmpYPlesPerMeter As Long      '��ֱ�ֱ���
    BmpClrUsed As Long            'λͼʵ��ʹ�õ���ɫ���е���ɫ��ַ��
    BmpClrImportant As Long       'λͼ��ʾ�����б���Ϊ��Ҫ��ɫ��ַ��
End Type
Public Type BmpFile
    'λͼ�ļ�����Ϣ
    Bmp_BmpFileHeard As BmpFileHeard
    Bmp_BmpPictureHeard As BmpPictureHeard
    Bmp_Bmp() As BmpRGB
End Type
Private Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type

Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As Any, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long

Public Enum GradientFillStyle
 GRADIENT_FILL_RECT_H = 0&
 GRADIENT_FILL_RECT_V = 1&
End Enum

Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CreateActCtxW Lib "kernel32.dll" (ByRef pActCtx As ACTCTXW) As Long
Private Declare Function ActivateActCtx Lib "kernel32.dll" (ByVal hActCtx As Long, ByRef lplpCookie As Long) As Long

Private Type ACTCTXW
 CBSIZE As Long
 dwFlags As Long
 lpcwstrSource As Long
 wProcessorArchitecture As Integer
 wLangId As Integer
 lpcwstrAssemblyDirectory As Long
 lpcwstrResourceName As Long
 lpcwstrApplicationName As Long
 hModule As Long
End Type

Private Const ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID As Long = 1
Private Const ACTCTX_FLAG_LANGID_VALID As Long = 2
Private Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID As Long = 4
Private Const ACTCTX_FLAG_RESOURCE_NAME_VALID As Long = 8
Private Const ACTCTX_FLAG_SET_PROCESS_DEFAULT As Long = 16
Private Const ACTCTX_FLAG_APPLICATION_NAME_VALID As Long = 32
Private Const ACTCTX_FLAG_HMODULE_VALID As Long = 128
Private Declare Function CreateDIBitmap_1 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_1, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_2 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_2, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_4 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_4, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_8 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_8, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_16 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_16, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24a Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24a, ByVal wUsage As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As Any, ByVal wUsage As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function GetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
Public Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Public Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Const STRETCH_ANDSCANS = 1    'Ĭ������.�޳����߶���ʣ�µ��߶ν���AND����.���ģʽͨ��Ӧ���ڲ����˰�ɫ�����ĵ�ɫλͼ
Public Const STRETCH_ORSCANS = 2     '�޳����߶α��򵥵����.���ģʽͨ�����ڲ�ɫλͼ
Public Const STRETCH_DELETESCANS = 3 '�޳����߶���ʣ�µ��߶ν���OR����.���ģʽͨ��Ӧ���ڲ����˰�ɫ�����ĵ�ɫλͼ
Public Const STRETCH_HALFTONE = 4    'Ŀ��λͼ�ϵ����ؿ鱻��ΪԴλͼ�ϴ��½��ƵĿ�.���ģʽҪ������������ģʽ
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function CountClipboardFormats Lib "user32" () As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const CF_BITMAP = 2
Public Const CF_DIB = 8
Public Const BI_RGB = 0&
Public Const BI_RLE4 = 2&
Public Const BI_RLE8 = 1&
Public Const BI_BitFields = 3&
Public Const BI_JPEG = 4&
Public Const BI_PNG = 5&
Public Type RECTL
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type BitmapData
  Width As Long
  Height As Long
  stride As Long
  PixelFormat As Long
  scan0 As Long
  Reserved As Long
End Type
Public Type PictureGDIBuffer
    GdipBitmap As Long
    ScaleWidth As Long
    ScaleHeight As Long
    GdipDC As Long
    GdipRect As RECTL
    GdipBitmapInto As BitmapData
End Type
Public Type PicturePNGBuffer
    hdc   As Long               '��ͼ
    ScaleWidth  As Long         '��ͼ�����أ�
    ScaleHeight As Long         '��ͼ�ߣ����أ�
    PictureColor() As Byte      '��ͼ��ɫ����
    aHDC   As Long              '(Alpha)ͼ
    aPictureColor() As Byte     '(Alpha)ͼ��ɫ����
    mHDC   As Long              '(Map)ͼ
    mPictureColor() As Byte     '(Map)ͼ��ɫ����
    ClipWidthHN As Integer      '����ͼƬ��
    ClipHeightVN As Integer     '����ͼƬ��
    ClipDC As Long              '����ͼ
    ClipColor() As Byte         '����ͼ��ɫ����
    ClipIndex As Integer        '�ڼ�����ͼ
    ClipCount As Integer        '��ͼ����
    ClipScaleWidth As Long      '��ͼ��
    ClipScaleHeight As Long     '��ͼ��
    ClipSetWidth As Long        '��ͼ����
    ClipSetHeight As Long       '��ͼ����
    ClipRenderInter As Integer  'ˢ�¼��
    aClipDC As Long             '(Alpha)����ͼ
    aClipColor() As Byte        '(Alpha)����ͼ��ɫ����
    mClipDC As Long             '(Map)����ͼ
    mClipColor() As Byte        '(Map)����ͼ��ɫ����
End Type
Public Type MoveDataNow
Height As Byte
Width As Byte
Temp2 As Byte
Temp1 As Byte
End Type
Public Const DIB_PAL_COLORS As Long = 1    ' BITMAPINFO������һ��16λ��ɫ������������
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const INVALID_HANDLE_VALUE = -1
Public Const CREATE_ALWAYS = 2
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long '����λͼ����
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Public Type BITMAPINFOHEADER
    biSize                                 As Long    '/* �ṹ���� */
    biWidth                                As Long    '/* ָ��λͼ�Ŀ�ȣ�������Ϊ��λ */
    biHeight                               As Long    '/* ָ��λͼ�ĸ߶ȣ�������Ϊ��λ */
    biPlanes                               As Integer '/* ָ��Ŀ���豸�ļ���(����Ϊ 1 ) */
    biBitCount                             As Integer '/* λͼ����ɫλ��,ÿһ�����ص�λ(1��4��8��16��24��32) */
    biCompression                          As Long    '/* ָ��ѹ������(BI_RGB Ϊ��ѹ��) */
    biSizeImage                            As Long    '/* ͼ��Ĵ�С,���ֽ�Ϊ��λ,����BI_RGB��ʽ��,������Ϊ0 */
    biXPelsPerMeter                        As Long    '/* ָ���豸ˮ׼�ֱ��ʣ���ÿ�׵�����Ϊ��λ */
    biYPelsPerMeter                        As Long    '/* ��ֱ�ֱ��ʣ�����ͬ�� */
    biClrUsed                              As Long    '/* ˵��λͼʵ��ʹ�õĲ�ɫ���е���ɫ������,��Ϊ0�Ļ�,˵��ʹ�����е�ɫ���� */
    biClrImportant                         As Long    '/* ˵����ͼ����ʾ����ҪӰ�����ɫ��������Ŀ�������0����ʾ����Ҫ */
End Type
Public Type RGBQUAD
    rgbBlue                                As Byte
    rgbGreen                               As Byte
    rgbRed                                 As Byte
    rgbReserved                            As Byte    '/* '����������Ϊ 0 */
End Type
Public Enum KhanImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum
Public Enum KhanImageFalgs
    LR_COLOR = &H2                               '/*
    LR_COPYRETURNORG = &H4                       '/* ��ʾ����һ��ͼ��ľ�ȷ�����������Բ���cxDesired��cyDesired
    LR_COPYDELETEORG = &H8                       '/* ��ʾ����һ��������ɾ��ԭʼͼ��.
    LR_CREATEDIBSECTION = &H2000                 '/* ������uTypeָ��ΪIMAGE_BITMAPʱ��ʹ�ú�������һ��DIB����λͼ��������һ�����ݵ�λͼ.�����־��װ��һ��λͼ��������ӳ��������ɫ����ʾ�豸ʱ�ǳ�����.
    LR_DEFAULTCOLOR = &H0                        '/* �Գ��淽ʽ����ͼ��
    LR_DEFAULTSIZE = &H40                        '/* �� cxDesired��cyDesiredδ����Ϊ�㣬ʹ��ϵͳָ���Ĺ���ֵ��ʶ����ͼ��Ŀ�͸�.��������������������cxDesired��cyDesired����Ϊ�㣬����ʹ��ʵ����Դ�ߴ�.�����Դ�������ͼ����ʹ�õ�һ��ͼ��Ĵ�С.
    LR_LOADFROMFILE = &H10                       '/* ���ݲ���lpszName��ֵװ��ͼ��.�����δ��������lpszName��ֵΪ��Դ����.
    LR_LOADMAP3DCOLORS = &H1000                  '/* ��ͼ���е����(Dk Gray RGB��128��128��128��).��(Gray RGB��192��192��192��).�Լ�ǳ��(Gray RGB��223��223��223��)���ض��滻��COLOR_3DSHADOW��COLOR_3DFACE�Լ�COLOR_3DLIGHT�ĵ�ǰ����
    LR_LOADTRANSPARENT = &H20                    '/* ��fuLoad����LR_LOADTRANSPARENT��LR_LOADMAP3DCOLORS����ֵ����LRLOADTRANSPARENT����.���ǣ���ɫ��ӿ���COLOR_3DFACE�����������COLOR_WINDOW.
    LR_MONOCHROME = &H1                          '/* ��ͼ��ת���ɵ�ɫ
    LR_SHARED = &H8000                           '/* ��ͼ�񽫱����װ������.���LR_SHAREDδ�����ã�������ͬһ����Դ�ڶ��ε������ͼ���Ǿͻ���װ���Ա����ͼ���ҷ��ز�ͬ�ľ��.
    LR_COPYFROMRESOURCE = &H4000                 '/*
    LR_VGACOLOR = &H80                           '/* ʹ�����ɫ.Uses true VGA colors.
End Enum
Type BITMAP
    bmType                                 As Long    '/* Type of bitmap */
    bmWidth                                As Long    '/* Pixel width */
    bmHeight                               As Long    '/* Pixel height */
    bmWidthBytes                           As Long    '/* Byte width = 3 x Pixel width */
    bmPlanes                               As Integer '/* Color depth of bitmap */
    bmBitsPixel                            As Integer '/* Bits per pixel, must be 16 or 24 */
    bmBits                                 As Long    '/* This is the pointer to the bitmap data */
End Type
Public Type BITMAPFILEHEADER
    bfType                                 As Integer '/* ָ���ļ����ͣ����� BM("magic cookie" - must be "BM" (19778)) */
    bfSize                                 As Long    '/* ָ��λͼ�ļ���С����λԪ��Ϊ��λ */
    bfReserved1                            As Integer '/* ������������Ϊ0 */
    bfReserved2                            As Integer '/* ͬ�� */
    bfOffBits                              As Long    '/* �Ӵ˼ܹ���λͼ����λ��λԪ��ƫ���� */
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" _
                Alias "CreateFontIndirectA" _
                (lpLogFont As LOGFONT) _
                As Long
                
Private Declare Function TextOut Lib "gdi32" _
                Alias "TextOutA" _
                (ByVal hdc As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal lpString As String, _
                ByVal nCount As Long) _
                As Long

Private Declare Function SetBkMode Lib "gdi32" _
                (ByVal hdc As Long, _
                ByVal nBkMode As Long) _
                As Long

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 50
End Type
Private Type ScTw
Width As Long
Height As Long
End Type
Private Type BITMAPINFO_1
bmiHeader As BITMAPINFOHEADER
bmiColors(1) As RGBQUAD
End Type
Private Type BITMAPINFO_2
bmiHeader As BITMAPINFOHEADER
bmiColors(3) As RGBQUAD
End Type
Private Type BITMAPINFO_4
bmiHeader As BITMAPINFOHEADER
bmiColors(15) As RGBQUAD
End Type
Private Type BITMAPINFO_8
bmiHeader As BITMAPINFOHEADER
bmiColors(255) As RGBQUAD
End Type
Private Type BITMAPINFO_16
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type BITMAPINFO_24
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type BITMAPINFO_24a
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTriple
End Type

'header
Private bm1 As BITMAPINFO_1
Private bm2 As BITMAPINFO_2
Private bm4 As BITMAPINFO_4
Private bm8 As BITMAPINFO_8
Private bm16 As BITMAPINFO_16
Private bm24 As BITMAPINFO_24
Private bm24a As BITMAPINFO_24a

'bitmap handle.
Private hBmp As Long

Dim RF As LOGFONT
Dim NewFont As Long
Dim OldFont As Long


Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, hImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As Long
Private Const UnitPixel As Long = &H2&
'�������Զ������ųߴ�
Public Sub DrawPicture(ByVal hdcDraw As Long, ByVal filename As String, Optional ByVal nLeft As Long = 0, Optional ByVal nTop As Long = 0, Optional MAXWIDTH As Long = 100, Optional MAXHEIGH As Long = 100)
    Dim hImage As Long
    Dim graphics As Long
    Dim Token As Long
    Dim GdipInput As GdiplusStartupInput
    Dim nWidth As Long
    Dim nHeight As Long
    
    GdipInput.GdiplusVersion = 1
    GdiplusStartup Token, GdipInput
    GdipLoadImageFromFile StrPtr(filename), hImage
     
'    GdipGetImageWidth hImage, nWidth
'    nWidth = nWidth * nScale
'    GdipGetImageHeight hImage, nHeight
'    nHeight = nHeight * nScale

    GdipCreateFromHDC hdcDraw, graphics
    GdipDrawImageRect graphics, hImage, nLeft, nTop, MAXWIDTH, MAXHEIGH
    GdipDeleteGraphics graphics
    GdipDisposeImage hImage
    GdiplusShutdown Token
End Sub
'�����ǰ���������
Public Sub DrawPictureByNum(ByVal hdcDraw As Long, ByVal filename As String, Optional ByVal nLeft As Long = 0, Optional ByVal nTop As Long = 0, Optional nScale As Double = 1)
    Dim hImage As Long
    Dim graphics As Long
    Dim Token As Long
    Dim GdipInput As GdiplusStartupInput
    Dim nWidth As Long
    Dim nHeight As Long

    GdipInput.GdiplusVersion = 1
    GdiplusStartup Token, GdipInput
    GdipLoadImageFromFile StrPtr(filename), hImage

    GdipGetImageWidth hImage, nWidth
    nWidth = nWidth * nScale
    GdipGetImageHeight hImage, nHeight
    nHeight = nHeight * nScale
    GdipCreateFromHDC hdcDraw, graphics
    GdipDrawImageRect graphics, hImage, nLeft, nTop, nWidth, nHeight
    GdipDeleteGraphics graphics
    GdipDisposeImage hImage
    GdiplusShutdown Token
End Sub
Public Sub DoTheStuff(ByVal hwnd As Long)
    SetWindowLong hwnd, -20, &H80000 '����͸��,��Щ��ɫΪ&H00000000&�ı�Ϊ͸��
    SetLayeredWindowAttributes hwnd, 0, 0, 1
End Sub
Sub FrmTrans(frm As Form)
tR = GetSetting("ICEE", "Main", "Tr", 0)
If tR = 1 Then
MakeTransparent frm.hwnd, 255
Else
MakeOpaque frm.hwnd
End If
End Sub
Public Function IsTransparent(ByVal hwnd As Long) As Boolean
On Error Resume Next
Dim Msg As Long
Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
  IsTransparent = True
Else
  IsTransparent = False
End If
If ERR Then IsTransparent = False
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If ERR Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
Dim Msg As Long
On Error Resume Next
Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
Msg = Msg And Not WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, Msg
SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If ERR Then MakeOpaque = 2
End Function
Private Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lwd As Long, hwd As Long
    Dim winpos As WINDOWPOS
    If Msg = WM_WINDOWPOSCHANGING Then
        CopyMemory winpos, ByVal lParam, Len(winpos)
        If winpos.X < 0 Then
            winpos.X = 0
            CopyMemory ByVal lParam, winpos, Len(winpos)
        End If
    End If
    WndProc = CallWindowProc(Prewininf, hwnd, Msg, wParam, lParam)
End Function
Sub SeekMe(TheFrm As Form)
    Dim Ret As Long
    '��¼�������Ϣ
    Prewininf = GetWindowLong(TheFrm.hwnd, GWL_WNDPROC)
    '���ƴ����λ��
    Ret = SetWindowLong(TheFrm.hwnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

Sub DRAWIT(ByVal hdc As Long, TXT As String, X, Y, ES)
     SetBkMode hdc, 1
     RF.lfHeight = 30
     '�����ַ��߶�
     RF.lfWidth = 10
     '�����ַ�ƽ�����
     RF.lfEscapement = 0
     '�����ı���б��
     RF.lfWeight = 400
     '�������������
     RF.lfItalic = 0
     '���岻��б
     RF.lfUnderline = 0
     '���岻���»���
     RF.lfStrikeOut = 0
     '���岻��ɾ����
     RF.lfOutPrecision = 0
     '�����������
     RF.lfClipPrecision = 0
     '���ü�������
     RF.lfQuality = 0
     '�����������
     RF.lfPitchAndFamily = 0
     '����������־��������
     RF.lfCharSet = 0
     '�����ַ���
     RF.lfFaceName = "Arial" + Chr(0)
     '������������
     Dim Throw As Long
     RF.lfEscapement = ES
    '�����ı���б��
     '�����������
     NewFont = CreateFontIndirect(RF)
     '����������
     OldFont = SelectObject(hdc, NewFont)
     'Ӧ��������

     'ѡ����ʾ�ı������
     Throw = TextOut(hdc, X, Y, TXT, Len(TXT))
     '��ʾ�ı�
     NewFont = SelectObject(hdc, OldFont)
     'ѡ�������
     Throw = DeleteObject(NewFont)
     'ɾ��������
End Sub

'''����Transparent() ����Դλͼ������������ X,Y λ�ã�ʹ��һ������͸��.
'''Transparent()�����������:һ����Ҫ���͸����Դλͼ,һ��Ŀ�� picturebox�ؼ� (PictDest),
'''һ��RGB��ɫֵ�����������������ԭλͼ��Ŀ�ĵ�����(destX �� destY��������Ϊ��λ).

Public Sub TRANSPARENT(ByVal sourceBmp As Long, Dest As Control, ByVal DestX As Integer, ByVal DestY As Integer, ByVal TransColor As Long)
    Const PIXEL = 3
    Dim SourceDC As Long 'Դλͼ
    Dim destScale As Long
    Dim maskDC As Long 'maskλͼ (monochrome)
    Dim saveDC As Long 'Դλͼ�ı���
    Dim resultDC As Long 'Դλͼ�뱳���ĺϲ�
    Dim invDC As Long 'Maskλͼ�ķ���ͼ
    Dim OrigColor As Long '����ɫ
    Dim Success As Long '���� Windows API�Ľ��
    Dim Bmp As BITMAP 'ԭλͼ�����ݽṹ˵��
    Dim hResultBmp As Long 'Դ�뱳����λͼ�ϲ�
    Dim hSaveBmp As Long 'ԭλͼ�Ŀ���
    Dim hSrcPrevBmp As Long
    Dim hDestPrevBmp As Long
    Dim hInvBmp As Long '��ת����λͼ (monochrome)
    Dim hPrevBmp As Long
    Dim hInvPrevBmp As Long
    Dim hSavePrevBmp As Long
    Dim hMaskBmp As Long
    Dim hMaskPrevBmp As Long
    
    
    destScale = Dest.ScaleMode '���� ScaleMode�Ա����ָ�
    Dest.ScaleMode = PIXEL '���� ScaleMode
    
    
    SourceDC = CreateCompatibleDC(Dest.hdc) '�����洢��DC
    saveDC = CreateCompatibleDC(Dest.hdc) '�����洢��DC
    
    invDC = CreateCompatibleDC(Dest.hdc) '�����洢��DC
    maskDC = CreateCompatibleDC(Dest.hdc) '�����洢��DC
    resultDC = CreateCompatibleDC(Dest.hdc) '�����洢��DC
    '����Դλͼ�õ����ĵĿ�Ⱥͳ��� (Bmp.bmWIDTH , Bmp.bmHeight)
    Success = GetObject(sourceBmp, Len(Bmp), Bmp)
    '������ɫ����λͼ
    hMaskBmp = CreateBitmap(Bmp.bmWidth, Bmp.bmHeight, 1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(Bmp.bmWidth, Bmp.bmHeight, 1, 1, ByVal 0&)
    
    hResultBmp = CreateCompatibleBitmap(Dest.hdc, Bmp.bmWidth, _
    Bmp.bmHeight)
    hSaveBmp = CreateCompatibleBitmap(Dest.hdc, Bmp.bmWidth, _
    Bmp.bmHeight)
    hSrcPrevBmp = SelectObject(SourceDC, sourceBmp)
    hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
    hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
    hInvPrevBmp = SelectObject(invDC, hInvBmp)
    hDestPrevBmp = SelectObject(resultDC, hResultBmp) 'ѡ��λͼ
    Success = BitBlt(saveDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, SourceDC, _
    0, 0, vbSrcCopy) '����Դλͼ�Ŀ����Ա����ָ�
    
    OrigColor = SetBkColor(SourceDC, TransColor)
    Success = BitBlt(maskDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, SourceDC, _
    0, 0, vbSrcCopy)
    TransColor = SetBkColor(SourceDC, OrigColor)
    
    Success = BitBlt(invDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, maskDC, _
    0, 0, vbNotSrcCopy)
    '��������ͼ���������յ�͸��λͼ
    Success = BitBlt(resultDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, _
    Dest.hdc, DestX, DestY, vbSrcCopy)
    
    Success = BitBlt(resultDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, _
    maskDC, 0, 0, vbSrcAnd)
    Success = BitBlt(SourceDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, invDC, _
    0, 0, vbSrcAnd)
    
    Success = BitBlt(resultDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, _
    SourceDC, 0, 0, vbSrcInvert)
    
    Success = BitBlt(Dest.hdc, DestX, DestY, Bmp.bmWidth, Bmp.bmHeight, _
    resultDC, 0, 0, vbSrcCopy) '�ڱ�������ʾ͸��λͼ
    
    Success = BitBlt(SourceDC, 0, 0, Bmp.bmWidth, Bmp.bmHeight, saveDC, _
    0, 0, vbSrcCopy) '�ָ�λͼ
    'ѡ������Ա��ͷ�
    hPrevBmp = SelectObject(resultDC, hDestPrevBmp)
    hPrevBmp = SelectObject(SourceDC, hSrcPrevBmp)
    hPrevBmp = SelectObject(saveDC, hSavePrevBmp)
    hPrevBmp = SelectObject(invDC, hInvPrevBmp)
    hPrevBmp = SelectObject(maskDC, hMaskPrevBmp)
    '�ͷ���Դ
    Success = DeleteDC(saveDC)
    Success = DeleteDC(invDC)
    Success = DeleteDC(resultDC)
    Success = DeleteObject(hSaveBmp)
    Success = DeleteObject(hMaskBmp)
    Success = DeleteObject(hInvBmp)
    Success = DeleteDC(SourceDC)
    Success = DeleteDC(maskDC)
    Success = DeleteObject(hResultBmp)
    Dest.ScaleMode = destScale '�ָ� ScaleMode
End Sub

Public Sub InitColorTable_1(Optional Sorting As Integer = 1)
Dim Fb1 As Byte
Dim Fb2 As Byte
Select Case Sorting
Case 0
Fb1 = 255
Fb2 = 0
Case 1
Fb1 = 0
Fb2 = 255
End Select
bm1.bmiColors(0).rgbRed = Fb1
bm1.bmiColors(0).rgbGreen = Fb1
bm1.bmiColors(0).rgbBlue = Fb1
bm1.bmiColors(0).rgbReserved = 0
bm1.bmiColors(1).rgbRed = Fb2
bm1.bmiColors(1).rgbGreen = Fb2
bm1.bmiColors(1).rgbBlue = Fb2
bm1.bmiColors(1).rgbReserved = 0

End Sub
Public Sub InitColorTable_1Palette(Palettenbyte() As Byte)
If UBound(Palettenbyte) = 5 Then
bm1.bmiColors(0).rgbRed = Palettenbyte(0)
bm1.bmiColors(0).rgbGreen = Palettenbyte(1)
bm1.bmiColors(0).rgbBlue = Palettenbyte(2)
bm1.bmiColors(0).rgbReserved = 0
bm1.bmiColors(1).rgbRed = Palettenbyte(3)
bm1.bmiColors(1).rgbGreen = Palettenbyte(4)
bm1.bmiColors(1).rgbBlue = Palettenbyte(5)
bm1.bmiColors(1).rgbReserved = 0
Else
InitColorTable_1
End If
End Sub

Public Sub InitColorTable_8(ByteArray() As Byte)
'Construct the palette
'==================================================
    Dim Palette8() As RGBTriple
        ReDim Palette8(255)
        CopyMemory Palette8(0), ByteArray(0), UBound(ByteArray) + 1
    Dim nCount As Long
    On Error Resume Next
    'Create Palette
    For nCount = 0 To 255
    bm8.bmiColors(nCount).rgbBlue = Palette8(nCount).Blue
    bm8.bmiColors(nCount).rgbGreen = Palette8(nCount).Green
    bm8.bmiColors(nCount).rgbRed = Palette8(nCount).Red
    bm8.bmiColors(nCount).rgbReserved = 0
    Next nCount
End Sub
Public Sub InitColorTable_4(ByteArray() As Byte)
    Dim Palette4() As RGBTriple
        ReDim Palette4(15)
        CopyMemory Palette4(0), ByteArray(0), UBound(ByteArray) + 1

Dim i As Integer
' Create a color table
For i = 0 To 15
bm4.bmiColors(i).rgbRed = Palette4(i).Red
bm4.bmiColors(i).rgbGreen = Palette4(i).Green
bm4.bmiColors(i).rgbBlue = Palette4(i).Blue
bm4.bmiColors(i).rgbReserved = 0
Next i

End Sub


Public Sub CreateBitmap_1(ByteArray() As Byte, BmpWidth As Long, BmpHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
' Create a 1bit Bitmap
Dim hdc As Long
With bm1.bmiHeader
.biSize = Len(bm1.bmiHeader)
.biWidth = BmpWidth
        If Orientation = 0 Then
        .biHeight = BmpHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -BmpHeight
        End If
.biPlanes = 1
.biBitCount = 1
.biCompression = BI_RGB
.biSizeImage = 0
.biXPelsPerMeter = 0
.biYPelsPerMeter = 0
.biClrUsed = Colorused
.biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_1(hdc, bm1.bmiHeader, CBM_INIT, ByteArray(0), bm1, DIB_RGB_COLORS)
End Sub
Public Sub CreateBitmap_2(ByteArray() As Byte, BmpWidth As Long, BmpHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
' Create a 2bit Bitmap
Dim hdc As Long
With bm1.bmiHeader
.biSize = Len(bm1.bmiHeader)
.biWidth = BmpWidth
        If Orientation = 0 Then
        .biHeight = BmpHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -BmpHeight
        End If
.biPlanes = 1
.biBitCount = 2
.biCompression = BI_RGB
.biSizeImage = 0
.biXPelsPerMeter = 0
.biYPelsPerMeter = 0
.biClrUsed = Colorused
.biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_2(hdc, bm2.bmiHeader, CBM_INIT, ByteArray(0), bm2, DIB_RGB_COLORS)
End Sub

Public Sub CreateBitmap_4(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)

Dim hdc As Long
With bm4.bmiHeader
.biSize = Len(bm1.bmiHeader)
.biWidth = PicWidth
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
.biPlanes = 1
.biBitCount = 4
.biCompression = BI_RGB
.biSizeImage = 0
.biXPelsPerMeter = 0
.biYPelsPerMeter = 0
.biClrUsed = Colorused
.biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_4(hdc, bm4.bmiHeader, CBM_INIT, ByteArray(0), bm4, DIB_RGB_COLORS)
End Sub

Public Sub CreateBitmap_8(BitmapArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)

Dim hdc As Long
With bm8.bmiHeader
.biSize = Len(bm8.bmiHeader)
.biWidth = PicWidth
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
.biPlanes = 1
.biBitCount = 8
.biCompression = BI_RGB
.biSizeImage = 0
.biXPelsPerMeter = 0
.biYPelsPerMeter = 0
.biClrUsed = Colorused
.biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_8(hdc, bm8.bmiHeader, CBM_INIT, BitmapArray(0), bm8, DIB_RGB_COLORS)
End Sub

Public Sub DrawBitmap(PicWidth As Long, PicHeight As Long, PicObject As Object, Scalierung As Boolean, Optional X As Long = 0, Optional Y As Long = 0, Optional DrawToBG As Boolean = False)
On Error Resume Next
Dim cDC As Long
Dim a As Long
Dim b As Long
Dim Tergabe As ScTw
Dim realheight As Long
Dim realwidth As Long
PicObject.Cls
If TypeOf PicObject Is Form Then
'change ScaleMode direct
Else
b = PicObject.Parent.ScaleMode
PicObject.Parent.ScaleMode = 1
End If

a = PicObject.ScaleMode
PicObject.ScaleMode = 1
Select Case Scalierung
Case True
Tergabe = PixelToTwips(PicWidth, PicHeight)
If DrawToBG = False Then
PicObject.Height = Tergabe.Height
PicObject.Width = Tergabe.Width
End If
Case False
End Select
If DrawToBG = False Then
If PicObject.Height <> PicObject.ScaleHeight Then 'with Boarders
Tergabe = Twipstopixel(PicObject.Width, PicObject.Height)
realheight = Tergabe.Height
realwidth = Tergabe.Width
PicObject.Height = PicObject.Height + (PicObject.Height - PicObject.ScaleHeight)
PicObject.Width = PicObject.Width + (PicObject.Width - PicObject.ScaleWidth)
Else
PicObject.ScaleMode = 3
realheight = PicObject.ScaleHeight
realwidth = PicObject.ScaleWidth
End If
Else
realheight = Tergabe.Height
realwidth = Tergabe.Width
PicHeight = realheight
PicWidth = realwidth
End If
If hBmp Then
cDC = CreateCompatibleDC(PicObject.hdc)
SelectObject cDC, hBmp
Call StretchBlt(PicObject.hdc, X, Y, realwidth, realheight, cDC, 0, 0, PicWidth, PicHeight, SRCCOPY)
DeleteDC cDC
DeleteObject hBmp
hBmp = 0
End If
If TypeOf PicObject Is Form Then
'change ScaleMode direct
Else
PicObject.Parent.ScaleMode = b
End If
PicObject.ScaleMode = a
PicObject.PICTURE = PicObject.image
End Sub






Public Sub CreateBitmap_24(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional ThreeToOrToFour As Integer = 0)

Dim hdc As Long
Dim Bits() As RGBQUAD
Dim BitsA() As RGBTriple
Select Case ThreeToOrToFour
Case 0
ReDim Bits((UBound(ByteArray) / 4) - 1)
CopyMemory Bits(0), ByteArray(0), UBound(ByteArray)
    With bm24.bmiHeader
        .biSize = Len(bm24.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth        'Bitmap Width
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
        .biBitCount = 32                        '32 bit alignment
        .biPlanes = 1                           'Single plane
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With

Case 1
ReDim BitsA((UBound(ByteArray) / 3) - 1)
CopyMemory BitsA(0), ByteArray(0), UBound(ByteArray)

    With bm24a.bmiHeader
        .biSize = Len(bm24.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth        'Bitmap Width
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
        .biBitCount = 24                        '24 bit alignment
        .biPlanes = 1                           'Single plane
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With
End Select
' Get the DC.
hdc = GetDC(0)
Select Case ThreeToOrToFour
Case 0
hBmp = CreateDIBitmap_24(hdc, bm24.bmiHeader, CBM_INIT, Bits(0), bm24, DIB_RGB_COLORS)
Case 1
hBmp = CreateDIBitmap_24a(hdc, bm24a.bmiHeader, CBM_INIT, BitsA(0), bm24a, DIB_RGB_COLORS)
End Select
End Sub
Public Sub CreateBitmap_16(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
Dim hdc As Long

    With bm16.bmiHeader
        .biSize = Len(bm16.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth                       'Bitmap Width
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
        .biPlanes = 1                           'Single plane
        .biBitCount = 16                        '32 bit alignment
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_16(hdc, bm16.bmiHeader, CBM_INIT, ByteArray(0), bm16, DIB_RGB_COLORS)
End Sub

Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
PixelToTwips.Width = xwert * ux
uy = Screen.TwipsPerPixelY
PixelToTwips.Height = ywert * uy
End Function



Public Function Twipstopixel(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
Twipstopixel.Width = xwert / ux
uy = Screen.TwipsPerPixelY
Twipstopixel.Height = ywert / uy
End Function

Public Function InitColorTable_Grey(BitDepth As Integer, Optional To8Bit As Boolean = False) As Byte()
    Dim CurLevel As Integer
    Dim Tergabe() As Byte
    Dim n As Long
    Dim LevelDiff As Byte
    Dim Tbl() As RGBQUAD
    Dim Table3() As RGBTriple
    Erase bm8.bmiColors
    If BitDepth <> 16 Then
        ReDim Tbl(2 ^ BitDepth - 1)
        ReDim Table3(2 ^ BitDepth - 1)
    Else
        ReDim Tbl(255)
        ReDim Table3(255)
    End If
    LevelDiff = 255 / UBound(Tbl)
    
    For n = 0 To UBound(Tbl)
        With Tbl(n)
            .rgbRed = CurLevel
            .rgbGreen = CurLevel
            .rgbBlue = CurLevel
        End With
        With Table3(n)
            .Red = CurLevel
            .Green = CurLevel
            .Blue = CurLevel
        End With
        CurLevel = CurLevel + LevelDiff
        
    Next n
  Select Case BitDepth
  Case 1
  If To8Bit = True Then
   CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 8
  End If
  Case 2
   CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 16
  Case 4
    If To8Bit = True Then
   CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 64
  Else
     CopyMemory ByVal VarPtr(bm4.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 64
  End If
  Case 8
 CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 1024
  End Select
  ReDim Tergabe(((UBound(Table3) + 1) * 3) - 1)
  CopyMemory Tergabe(0), ByVal VarPtr(Table3(0).Red), ((UBound(Table3) + 1) * 3)
  InitColorTable_Grey = Tergabe
End Function

Public Sub GetRGBColors(ByVal RGBColor As Long, ByRef RedColor As Long, ByRef GreenColor As Long, ByRef BlueColor As Long)
    RedColor = RGBColor Mod 256
    GreenColor = (RGBColor \ &H100) Mod 256
    BlueColor = (RGBColor \ &H10000) Mod 256
End Sub

Public Sub GetBmpFile(Bmp_BmpFileName As String, Bmp_BmpFile As BmpFile)
    '��ȡBMP�ļ�
    Dim Bmp_RD As Integer
    
    Bmp_RD = FreeFile
    Open Bmp_BmpFileName For Binary As #Bmp_RD
    
    Get #Bmp_RD, 1, Bmp_BmpFile.Bmp_BmpFileHeard      '��ȡλͼ�ļ�ͷ
    Get #Bmp_RD, 15, Bmp_BmpFile.Bmp_BmpPictureHeard  '��ȡ��Ϣͷ
    
    '����ͼ��ĸ߶����ȳ�ʼ��ͼ�����е��±�
    ReDim Bmp_BmpFile.Bmp_Bmp(1 To Bmp_BmpFile.Bmp_BmpPictureHeard.BmpWidth, _
        1 To Bmp_BmpFile.Bmp_BmpPictureHeard.BmpHeight)
    '��ȡͼ������
    Get #Bmp_RD, Bmp_BmpFile.Bmp_BmpFileHeard.BmpOffBits + 1, _
        Bmp_BmpFile.Bmp_Bmp
    'Bmp_BmpFile.Bmp_BmpFileHeard.BmpOffBits + 1����Ϊ
    'Bmp_BmpFile.Bmp_BmpFileHeard.BmpOffBits��¼��������
    '��0Ϊ��ʼ������������1Ϊ��ʼ������Ҫ��1
    
    Close #Bmp_RD
End Sub

Public Sub PutBmpFile(Bmp_BmpFileName As String, Bmp_BmpFile As BmpFile)
    'дBMP�ļ�
    Dim Bmp_WR As Integer
    
    Bmp_WR = FreeFile
    Open Bmp_BmpFileName For Binary As #Bmp_WR
    
    Put #Bmp_WR, 1, Bmp_BmpFile.Bmp_BmpFileHeard      'дλͼ�ļ�ͷ
    Put #Bmp_WR, 15, Bmp_BmpFile.Bmp_BmpPictureHeard  'д��Ϣͷ
    Put #Bmp_WR, Bmp_BmpFile.Bmp_BmpFileHeard.BmpOffBits + 1, _
        Bmp_BmpFile.Bmp_Bmp   'дͼ������
    'Bmp_BmpFile.Bmp_BmpFileHeard.BmpOffBits + 1����Ϊ
    'Bmp_BmpFile.Bmp_BmpFileHeard.BmpOffBits��¼������
    '����0Ϊ��ʼ������������1Ϊ��ʼ������Ҫ��1
    
    Close #Bmp_WR
End Sub
Public Sub YouHua(Bmp_SBmpFileName As BmpFile, Bmp_DBmpFileName As BmpFile, _
    Bmp_BmpSize As Integer)
    '�ͻ�Ч��
    Dim i As Integer, j As Integer, a As Integer, b As Integer
    Bmp_DBmpFileName = Bmp_SBmpFileName
    For i = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth - Bmp_BmpSize
        For j = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight - Bmp_BmpSize
            a = Rnd() * (Bmp_BmpSize - 1) + 1
            b = Rnd() * (Bmp_BmpSize - 1) + 1
            Bmp_DBmpFileName.Bmp_Bmp(i, j).Red = _
                Bmp_DBmpFileName.Bmp_Bmp(i + a, j + b).Red
            Bmp_DBmpFileName.Bmp_Bmp(i, j).Green = _
                Bmp_DBmpFileName.Bmp_Bmp(i + a, j + b).Green
            Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue = _
                Bmp_DBmpFileName.Bmp_Bmp(i + a, j + b).Blue
        Next j
    Next i
End Sub

Public Sub MuKe(Bmp_SBmpFileName As BmpFile, Bmp_DBmpFileName As BmpFile)
    'ľ��Ч��
    Dim i As Integer, j As Integer
    Bmp_DBmpFileName = Bmp_SBmpFileName
    For i = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth
        For j = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight
            If (CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Red) + _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Green) + _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue)) / 3 > 128 Then
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Red = 0
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Green = 0
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue = 0
            Else
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Red = &HFF
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Green = &HFF
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue = &HFF
            End If
        Next j
    Next i
End Sub

Public Sub FuDiao(Bmp_SBmpFileName As BmpFile, Bmp_DBmpFileName As BmpFile)
    '����Ч��
    Dim i As Long, j As Long
    Bmp_DBmpFileName = Bmp_SBmpFileName
    For i = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth - 1
        For j = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight - 1
            Bmp_DBmpFileName.Bmp_Bmp(i, j).Red = _
                IIf((CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Red) - _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Red) + 128) _
                > 255, 255, IIf((CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Red) _
                - CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Red) + 128) _
                < 0, 0, (CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Red) - _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Red) + 128)))
            Bmp_DBmpFileName.Bmp_Bmp(i, j).Green = _
                IIf((CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Green) - _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Green) + 128) _
                > 255, 255, IIf((CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Green) _
                - CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Green) + 128) _
                < 0, 0, (CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Green) - _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Green) + 128)))
            Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue = _
                IIf((CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue) - _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Blue) + 128) _
                > 255, 255, IIf((CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue) _
                - CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Blue) + 128) _
                < 0, 0, (CLng(Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue) - _
                CLng(Bmp_DBmpFileName.Bmp_Bmp(i + 1, j + 1).Blue) + 128)))
        Next j
    Next i
  
    Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight = _
        Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight - 1
    Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth = _
        Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth - 1
End Sub

Public Sub DengGuang(Bmp_SBmpFileName As BmpFile, Bmp_DBmpFileName As BmpFile, _
    X As Long, Y As Long, m As Long, n As Long)
    '�ƹ�Ч��
    Dim i As Long, j As Long, r As Long, G As Long, b As Long
    Bmp_DBmpFileName = Bmp_SBmpFileName
    For i = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth - 1
        For j = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight - 1
            If Sqr((i - X) ^ 2 + (j - Y) ^ 2) - 60 < 0 Then
                r = Bmp_DBmpFileName.Bmp_Bmp(i, j).Red + (m * (1 - _
                    (Sqr((i - X) ^ 2 + (j - Y) ^ 2) + n) / Y))
                G = Bmp_DBmpFileName.Bmp_Bmp(i, j).Green + (m * (1 _
                    - (Sqr((i - X) ^ 2 + (j - Y) ^ 2) + n) / Y))
                b = Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue + (m * (1 - _
                    (Sqr((i - X) ^ 2 + (j - Y) ^ 2) + n) / Y))
                
                If r < 0 Then r = 0
                If r > 255 Then r = 255
                
                If G < 0 Then G = 0
                If G > 255 Then G = 255
                
                If b < 0 Then b = 0
                If b > 255 Then b = 255
                
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Red = r
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Green = G
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue = b
              
            End If
        Next j
    Next i
End Sub

Public Sub MoShu(Bmp_SBmpFileName As BmpFile, Bmp_DBmpFileName As BmpFile)
    'ͼ��ħ��
    Dim i As Integer, j As Integer
    Dim r As Integer, G As Integer, b As Integer
    Dim Y As Long, Cr As Long, cb As Long
    Bmp_DBmpFileName = Bmp_SBmpFileName
    For i = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth
        For j = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight
            If (i + (Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight - j + _
                1)) Mod 2 = 0 Then
                '��Ϊ�еĶ�ȡ���������ϵģ������������еĵ�һ����ʵ�����һ��
                'Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight - j + 1��
                '�����һ������Ӧ����ʵ���к�
                
                r = Bmp_DBmpFileName.Bmp_Bmp(i, j).Red
                G = Bmp_DBmpFileName.Bmp_Bmp(i, j).Green
                b = Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue
                
                Y = 0.299 * r + 0.587 * G + 0.114 * b
                cb = -0.1687 * r - 0.3313 * G + 0.5 * b
                Cr = 0.5 * r - 0.4187 * G - 0.0813 * b
                
                cb = -cb
                Cr = -Cr
                
                r = Y + 1.402 * Cr
                G = Y - 0.34414 * cb - 0.71414 * Cr
                b = Y + 1.772 * cb
                
                If r > 255 Then r = 255
                If G > 255 Then G = 255
                If b > 255 Then b = 255
                
                If r < 0 Then r = 0
                If G < 0 Then G = 0
                If b < 0 Then b = 0
                
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Red = r
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Green = G
                Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue = b
                
            End If
        Next j
    Next i
End Sub

Public Sub YYMS(Bmp_SBmpFileName As BmpFile, Bmp_DBmpFileName As BmpFile, _
    File_File() As Byte)
    '����
    Dim i As Long, j As Long, Z As Integer, a As Integer, K As Long
    Dim RMask As Byte, GMask As Byte, BMask As Byte
    Dim File_Bin() As Byte
    ReDim File_Bin(1 To UBound(File_File) * 8) As Byte
    
    RMask = 2: GMask = 1: BMask = 3
    
    Bmp_DBmpFileName = Bmp_SBmpFileName
      
    For i = 1 To UBound(File_File)
        For j = 1 To 8
            a = 2 ^ (8 - j)
            
            If (File_File(i) And a) <> 0 Then
                File_Bin((i - 1) * 8 + j) = 1
            Else
                File_Bin((i - 1) * 8 + j) = 0
            End If
        Next j
    Next i
    
    K = 1
    For i = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpWidth
        For j = 1 To Bmp_DBmpFileName.Bmp_BmpPictureHeard.BmpHeight
            YXMS_Mask Bmp_DBmpFileName.Bmp_Bmp(i, j).Red, RMask, File_Bin, K
            YXMS_Mask Bmp_DBmpFileName.Bmp_Bmp(i, j).Green, GMask, File_Bin, K
            YXMS_Mask Bmp_DBmpFileName.Bmp_Bmp(i, j).Blue, BMask, File_Bin, K
        Next j
    Next i
    Bmp_DBmpFileName.Bmp_BmpFileHeard.BmpReserved = UBound(File_File)
End Sub

Private Sub YXMS_Mask(Bmp_Color As Byte, Mask As Byte, File_Bin() _
    As Byte, K As Long)
    '���͵����ݲ���
    Dim Z As Integer, a As Byte
    For Z = 1 To Mask
        If K <= UBound(File_Bin) Then
            a = 2 ^ (Mask - Z)
            If File_Bin(K) = 1 Then
                Bmp_Color = (Bmp_Color Or a)
            Else
                Bmp_Color = (Bmp_Color And (a Xor &HFF))
            End If
            K = K + 1
        End If
    Next Z
End Sub

Public Sub XXMS(Bmp_SBmpFileName As BmpFile, File_File() As Byte)
    '����
    Dim i As Long, j As Long, Z As Integer, a As Integer, K As Long
    Dim RMask As Byte, GMask As Byte, BMask As Byte
    Dim File_Bin() As Byte
    
    ReDim File_File(1 To Bmp_SBmpFileName.Bmp_BmpFileHeard.BmpReserved) As Byte
    ReDim File_Bin(1 To UBound(File_File) * 8) As Byte
    
    RMask = 2: GMask = 1: BMask = 3
    
    K = 1
    For i = 1 To Bmp_SBmpFileName.Bmp_BmpPictureHeard.BmpWidth
        For j = 1 To Bmp_SBmpFileName.Bmp_BmpPictureHeard.BmpHeight
            XXMS_Mask Bmp_SBmpFileName.Bmp_Bmp(i, j).Red, RMask, File_Bin, K
            XXMS_Mask Bmp_SBmpFileName.Bmp_Bmp(i, j).Green, GMask, File_Bin, K
            XXMS_Mask Bmp_SBmpFileName.Bmp_Bmp(i, j).Blue, BMask, File_Bin, K
        Next j
    Next i
    
    For i = 1 To UBound(File_File)
        For j = 1 To 8
            a = 2 ^ (8 - j)
            
            If File_Bin((i - 1) * 8 + j) = 1 Then
                File_File(i) = (File_File(i) Or a)
            Else
                File_File(i) = (File_File(i) And (a Xor &HFF))
            End If
        Next j
    Next i
  
End Sub

Private Sub XXMS_Mask(Bmp_Color As Byte, Mask As Byte, File_Bin() _
    As Byte, K As Long)
    '���͵����ݲ���
    Dim Z As Integer, a As Byte
    For Z = 1 To Mask
        If K <= UBound(File_Bin) Then
            a = 2 ^ (Mask - Z)
            If (Bmp_Color And a) = 0 Then
                File_Bin(K) = 0
            Else
                File_Bin(K) = 1
            End If
            K = K + 1
        End If
    Next Z
End Sub

Public Sub ShadePicture(picSource As PictureBox, PicTarget As PictureBox, WithColor As Long, Thickness As Integer)
On Error Resume Next
Dim sRate, Col As Long
Dim X, Y As Single
Dim XMax, YMax As Single
Dim cBlue, cGreen, cRed As Double   'Determines the pixel color
Dim sBlue, sGreen, sRed As Double   'Determines the SHADING color
    'Getting the RGB values of selected color
    sBlue = Fix((WithColor / 256) / 256)
    sGreen = Fix((WithColor - ((sBlue * 256) * 256)) / 256)
    sRed = Fix(WithColor - ((sBlue * 256) * 256) - (sGreen * 256))
    'Calculate screen height & width of the image
    XMax = picSource.Width / Screen.TwipsPerPixelX - 1
    YMax = picSource.Height / Screen.TwipsPerPixelY - 1
    'Initialising Shading
    PicTarget.Cls
    sRate = Thickness / 10
    'Process all pixels and alter them accordingly
    For X = 0 To XMax
      For Y = 0 To YMax
        Col = GetPixel(picSource.hdc, X, Y)
        If Not Col = 0 Then     'Because black colors are usually the borders of an image and never change border color.It will affect the clarity.
            'Getting the RGB values of current pixel
            cBlue = Fix((Col / 256) / 256)
            cGreen = Fix((Col - ((cBlue * 256) * 256)) / 256)
            cRed = Fix(Col - ((cBlue * 256) * 256) - (cGreen * 256))
            'Resetting the RGB values of current pixel with  the  sRate of  shading
            cRed = cRed + (sRed - cRed) * sRate
            cGreen = cGreen + (sGreen - cGreen) * sRate
            cBlue = cBlue + (sBlue - cBlue) * sRate
            If Not Col = 12632256 Then SetPixel PicTarget.hdc, X, Y, RGB(cRed, cGreen, cBlue)   'Skipping transparent col and setting the pixel
        Else
            SetPixel PicTarget.hdc, X, Y, Col
        End If
      Next Y
    PicTarget.Refresh
Next X
End Sub


Public Function GetSysLvwHandler() As Long
    GetSysLvwHandler = FindWindow("Progman", "Program Manager")
    If (GetSysLvwHandler <> 0) Then
        GetSysLvwHandler = FindWindowEx(GetSysLvwHandler, 0&, "SHELLDLL_DefView", vbNullString)
        If (GetSysLvwHandler <> 0) Then
            GetSysLvwHandler = FindWindowEx(GetSysLvwHandler, 0&, "SysListView32", vbNullString)
        End If
    End If
End Function
Public Sub ShowTransparency(cSrc As PictureBox, cDest As PictureBox, ByVal nLevel As Byte)
    Dim LrProps As rBlendProps
    Dim LnBlendPtr As Long
    
    cDest.Cls
    LrProps.tBlendAmount = nLevel
    CopyMemory LnBlendPtr, LrProps, 4
    With cSrc
        AlphaBlend cDest.hdc, 0, 0, .ScaleWidth, .ScaleHeight, _
            .hdc, 0, 0, .ScaleWidth, .ScaleHeight, LnBlendPtr
    End With
    cDest.Refresh
End Sub
Public Sub MakeTaskbarTransparent(ByVal bLevel As Byte)
    Dim lOldStyle As Long
    Dim lhwnd As Long
    
    lhwnd = FindWindow("Shell_TrayWnd", vbNullString)
    If (lhwnd <> 0) Then
        lOldStyle = GetWindowLong(lhwnd, GWL_EXSTYLE)
        SetWindowLong lhwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
        SetLayeredWindowAttributes lhwnd, 0, bLevel, LWA_ALPHA
    End If
    lhwnd = FindWindow("BaseBar", vbNullString)
    If (lhwnd <> 0) Then
        lOldStyle = GetWindowLong(lhwnd, GWL_EXSTYLE)
        SetWindowLong lhwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
        SetLayeredWindowAttributes lhwnd, 0, bLevel, LWA_ALPHA
    End If
End Sub


Public Sub GradientFillRect( _
      ByVal lhDC As Long, _
      ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
      ByVal lStartColor As Long, _
      ByVal lEndColor As Long, _
      ByVal eDir As GradientFillStyle _
   )

    lStartColor = TranslateColor(lStartColor)
    lEndColor = TranslateColor(lEndColor)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR As GRADIENT_RECT

    pSetTriVertexColor tTV(0), lStartColor
    tTV(0).X = Left
    tTV(0).Y = Top
    pSetTriVertexColor tTV(1), lEndColor
    tTV(1).X = Right
    tTV(1).Y = Bottom

    tGR.UpperLeft = 0
    tGR.LowerRight = 1

    GradientFill lhDC, tTV(0), 2, tGR, 1, eDir

End Sub

Public Function TranslateColor(ByVal clr As Long) As Long
If clr < 0 Then
 TranslateColor = GetSysColor(clr And &HFFFFFF)
Else
 TranslateColor = clr
End If
End Function
Private Sub pSetTriVertexColor(tTV As TRIVERTEX, LColor As Long)
   tTV.Red1 = (LColor And &HFF&)
   tTV.Green1 = (LColor And &HFF00&) \ &H100&
   tTV.Blue1 = (LColor And &HFF0000) \ &H10000
End Sub
Public Function pCreateDC(ByVal Width As Long, ByVal Height As Long) As Long
    Dim TmpDC   As Long
    Dim rDC     As Long
    Dim rBmp    As Long
    TmpDC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If TmpDC Then
        rDC = CreateCompatibleDC(TmpDC)
        If rDC Then
            rBmp = CreateCompatibleBitmap(TmpDC, Width, Height)
            If rBmp Then
                DeleteObject SelectObject(rDC, rBmp)
                pCreateDC = rDC
                DeleteObject rBmp
            Else
                DeleteDC rDC
            End If
        End If
        DeleteDC TmpDC
    End If
End Function

Public Function pCreateDCByHandle(ByVal handle As Long) As Long
    Dim TmpDC   As Long
    Dim rDC     As Long
    Dim rBmp    As Long
    TmpDC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If TmpDC Then
        rDC = CreateCompatibleDC(TmpDC)
        If rDC Then
            DeleteObject SelectObject(rDC, handle)
            pCreateDCByHandle = rDC
        End If
        DeleteDC TmpDC
    End If
End Function
Public Sub ѣ��ͼ��(bPic As PictureBox)
Dim aa As Long, bb As Long
Dim pict() As Byte
   Dim av As Long
   Dim Ptr As Long
   Dim safe As SAFEARRAY1D, Bmp As BITMAP
    Call GetObject(bPic.PICTURE, Len(Bmp), Bmp)
    With safe
      .cbElements = 1
      .cDims = 1
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = Bmp.bmHeight * Bmp.bmWidthBytes
      .pvData = Bmp.bmBits
    End With
    Call CopyMemory(ByVal VarPtrArray(pict), VarPtr(safe), 4)
    On Error Resume Next
    Ptr = Bmp.bmWidthBytes + 3
    For aa = 1 To Bmp.bmHeight - 3
      For bb = 0 To Bmp.bmWidthBytes
        Ptr = Ptr + 1
        av = pict(Ptr - Bmp.bmWidthBytes)
        av = av + pict(Ptr - 30)
        av = av + pict(Ptr + 30) '����ͼ���ص�ֵ
        av = av + pict(Ptr + Bmp.bmWidthBytes)
        pict(Ptr) = av \ 4 'ֵԽ��,ͼƬԽ��
      Next bb
    Next aa
    Call CopyMemory(ByVal VarPtrArray(pict), 0&, 4)
End Sub
Public Sub ����ģ��(bPic As PictureBox)
Dim aa As Long, bb As Long
Dim pict() As Byte
   Dim av As Long
   Dim Ptr As Long
   Dim safe As SAFEARRAY1D, Bmp As BITMAP
    Call GetObject(bPic.PICTURE, Len(Bmp), Bmp)
    With safe
      .cbElements = 1
      .cDims = 1
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = Bmp.bmHeight * Bmp.bmWidthBytes
      .pvData = Bmp.bmBits
    End With
    Call CopyMemory(ByVal VarPtrArray(pict), VarPtr(safe), 4)
    On Error Resume Next
    Ptr = Bmp.bmWidthBytes
    For aa = 0 To Bmp.bmHeight
      For bb = 0 To Bmp.bmWidthBytes
        Ptr = Ptr + 1
        av = pict(Ptr - Bmp.bmWidthBytes)
        av = av + pict(Ptr - 4)
        av = av + pict(Ptr + 4) '����ͼ���ص�ֵ
        av = av + pict(Ptr + Bmp.bmWidthBytes)
        pict(Ptr) = av \ 4 'ֵԽ��,ͼƬԽ��
      Next bb
    Next aa
    Call CopyMemory(ByVal VarPtrArray(pict), 0&, 4)
End Sub
