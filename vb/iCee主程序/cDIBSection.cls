VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private m_hDIB As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the CLSPICDIBS:
Private m_hDC As Long
' Address of memory pointing to the CLSPICDIBS's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Const CF_BITMAP = 2
Private Const CF_DIB = 8

Private Const BITMAPTYPE As Integer = &H4D42
Private Type BITMAPFILEHEADER
   bfType As Integer '- type  ="BM" i.e &H4D42 - 2
   bfSize As Long ' - size in bytes of file - 6
   bfReserved1 As Integer ' - reserved, must be 0 - 8
   bfReserved2 As Integer ' - reserved, must be 0 - 10
   bfOffBits As Long ' offset from this structure to the bitmap bits - 14
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Function CopyToClipboard( _
        Optional ByVal bAsDIB As Boolean = True _
    ) As Boolean
Dim lhDCDesktop As Long
Dim lhDC As Long
Dim lhBmpOld As Long
Dim hObj As Long
Dim lFmt As Long
Dim b() As Byte
Dim tBI As BITMAPINFO
Dim lPtr As Long
Dim hDibCopy As Long

   lhDCDesktop = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    If (lhDCDesktop <> 0) Then
        lhDC = CreateCompatibleDC(lhDCDesktop)
        If (lhDC <> 0) Then
            If (bAsDIB) Then
               Call SHOWWRONG("I don't know how to put a DIB on the clipboard! Copy as bitmap instead!!!", 0)
                ' Create a duplicate CLSPICDIBS and copy
                ' to the clipboard:
                'LSet tBI = m_tBI
                'hDibCopy = CreateDIBSection( _
                '        lhDC, _
                '        m_tBI, _
                '        DIB_RGB_COLORS, _
                '        lPtr, _
                '        0, 0)
                'If (hDibCopy <> 0) Then
                '    lhBmpOld = SelectObject(lhDC, hObj)
                '    BitBlt lhDC, 0, 0, Width, Height, m_hDC, 0, 0, vbSrcCopy
                '    SelectObject lhDC, lhBmpOld
                '    lFmt = CF_DIB
                '
                '     '....
                                    
                'Else
                '    hObj = 0
                'End If
            Else
                ' Create a compatible bitmap and copy to
                ' the clipboard:
                hObj = CreateCompatibleBitmap(lhDCDesktop, Width, Height)
                If (hObj <> 0) Then
                    lhBmpOld = SelectObject(lhDC, hObj)
                    PaintPicture lhDC
                    SelectObject lhDC, lhBmpOld
                    lFmt = CF_BITMAP
                    ' Now set the clipboard to the bitmap:
                    If (OpenClipboard(0) <> 0) Then
                        EmptyClipboard
                        If (SetClipboardData(lFmt, hObj) <> 0) Then
                            CopyToClipboard = True
                        End If
                        CloseClipboard
                    End If
                End If
            End If
            DeleteDC lhDC
        End If
    End If
End Function

Public Function CreateDIB( _
        ByVal lhDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection( _
            lhDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
    CreateDIB = (hDib <> 0)
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim lhDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As BITMAP
    
    GetObjectAPI picThis.handle, Len(tBMP), tBMP
    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhDCDesktop = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        If (lhDCDesktop <> 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If (lhDC <> 0) Then
                lhBmpOld = SelectObject(lhDC, picThis.handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhBmpOld
                DeleteObject lhDC
            End If
        End If
    End If
End Function
Public Function CreateFromFile(ByVal filename As String)
CreateFromPicture LoadPicture(filename)
End Function
Public Sub Cls()
BitBlt m_hDC, 0, 0, Width, Height, 0, 0, 0, &H42&
End Sub
Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
Dim lHDCDesk As Long
    ClearUp
    lHDCDesk = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    m_hDC = CreateCompatibleDC(lHDCDesk)
    DeleteDC lHDCDesk
    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIB)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIB)
            Create = True
        Else
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If
End Function


Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lhDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop
End Sub


Public Sub PaintPicture( _
        ByVal lhDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hdc() As Long
    hdc = m_hDC
End Property
Public Property Get hDib() As Long
    hDib = m_hDIB
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
'Public Sub GetBits(b() As Byte) 'crash..
'Erase b
'Dim t As SAFEARRAY2D
'With t
'    .cbElements = 1
'    .cDims = 2
'    .Bounds(0).lLbound = 0
'    .Bounds(0).cElements = Height
'    .Bounds(1).lLbound = 0
'    .Bounds(1).cElements = BytesPerScanLine
'    .pvData = m_lPtr
'End With
'CopyMemory ByVal VarPtrArray(b()), VarPtr(t), 4
'End Sub
'Public Sub DestroyBits(b() As Byte)
'CopyMemory ByVal VarPtrArray(b()), 0&, 4
'End Sub
Public Sub RandomiseBits( _
        Optional ByVal bGray As Boolean = False _
    )
Dim bDib() As Byte
Dim X As Long, Y As Long
Dim LC As Long
Dim tSA As SAFEARRAY2D
Dim xEnd As Long
    
    ' Get the bits in the from DIB section:
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

    ' random:
    Randomize Timer
    
    xEnd = (Width - 1) * 3
    If (bGray) Then
        For Y = 0 To m_tBI.bmiHeader.biHeight - 1
            For X = 0 To xEnd Step 3
                LC = Rnd * 255
                bDib(X, Y) = LC
                bDib(X + 1, Y) = LC
                bDib(X + 2, Y) = LC
            Next X
        Next Y
    Else
        For X = 0 To xEnd Step 3
            For Y = 0 To m_tBI.bmiHeader.biHeight - 1
                bDib(X, Y) = Rnd * 255
                bDib(X + 1, Y) = Rnd * 255
                bDib(X + 2, Y) = Rnd * 255
            Next Y
        Next X
    End If
    
    ' Clear the temporary array descriptor
   ' This is necessary under NT4.
   CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIB
        End If
        DeleteObject m_hDC
    End If
    m_hDC = 0: m_hDIB = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Public Function Resample( _
        ByVal lNewHeight As Long, _
        ByVal lNewWidth As Long _
    ) As cDIBSection
Dim cDib As cDIBSection
    Set cDib = New cDIBSection
    If cDib.Create(lNewWidth, lNewHeight) Then
        If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
            ' Change in size, do resample:
            ResampleDib cDib
        Else
            ' No size change so just return a copy:
            cDib.LoadPictureBlt m_hDC
        End If
        Set Resample = cDib
    End If
End Function

Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
Dim bDibFrom() As Byte
Dim bDibTo() As Byte

Dim tSAFrom As SAFEARRAY2D
Dim tSATo As SAFEARRAY2D

    ' Get the bits in the from DIB section:
    With tSAFrom
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    ' Get the bits in the to DIB section:
    With tSATo
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibTo.BytesPerScanLine()
        .pvData = cDibTo.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

Dim xScale As Single
Dim yScale As Single

Dim X As Long, Y As Long, xEnd As Long, xOut As Long

Dim fX As Single, fY As Single
Dim ifY As Long, ifX As Long
Dim dX As Single, dY As Single
Dim r As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single
Dim G As Long, G1 As Single, g2 As Single, g3 As Single, g4 As Single
Dim b As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single
Dim ir1 As Long, ig1 As Long, ib1 As Long
Dim ir2 As Long, ig2 As Long, ib2 As Long

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    
    xEnd = cDibTo.Width - 1
        
    For Y = 0 To cDibTo.Height - 1
        
        fY = Y * yScale
        ifY = Int(fY)
        dY = fY - ifY
        
        For X = 0 To xEnd
            fX = X * xScale
            ifX = Int(fX)
            dX = fX - ifX
            
            ifX = ifX * 3
            ' Interpolate using the four nearest pixels in the source
            b1 = bDibFrom(ifX, ifY): G1 = bDibFrom(ifX + 1, ifY): r1 = bDibFrom(ifX + 2, ifY)
            b2 = bDibFrom(ifX + 3, ifY): g2 = bDibFrom(ifX + 4, ifY): r2 = bDibFrom(ifX + 5, ifY)
            b3 = bDibFrom(ifX, ifY + 1): g3 = bDibFrom(ifX + 1, ifY + 1): r3 = bDibFrom(ifX + 2, ifY + 1)
            b4 = bDibFrom(ifX + 3, ifY + 1): g4 = bDibFrom(ifX + 4, ifY + 1): r4 = bDibFrom(ifX + 5, ifY + 1)
            
            ' Interplate in x direction:
            ir1 = r1 * (1 - dY) + r3 * dY: ig1 = G1 * (1 - dY) + g3 * dY: ib1 = b1 * (1 - dY) + b3 * dY
            ir2 = r2 * (1 - dY) + r4 * dY: ig2 = g2 * (1 - dY) + g4 * dY: ib2 = b2 * (1 - dY) + b4 * dY
            ' Interpolate in y:
            r = ir1 * (1 - dX) + ir2 * dX: G = ig1 * (1 - dX) + ig2 * dX: b = ib1 * (1 - dX) + ib2 * dX
            
            ' Set output:
            If (r < 0) Then r = 0
            If (r > 255) Then r = 255
            If (G < 0) Then G = 0
            If (G > 255) Then G = 255
            If (b < 0) Then b = 0
            If (b > 255) Then b = 255
            xOut = X * 3
            bDibTo(xOut, Y) = b
            bDibTo(xOut + 1, Y) = G
            bDibTo(xOut + 2, Y) = r
            
        Next X
        
    Next Y

    ' Clear the temporary array descriptor
    ' This is necessary under NT4.
    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4


End Function


Public Function SavePicture(ByVal sFileName As String) As Boolean
Dim tBH As BITMAPFILEHEADER
Dim tRGBQ As RGBQUAD
Dim hFile As Long
Dim lBytesWritten As Long
Dim lSize As Long
Dim lR As Long
Dim bErr As Boolean
Dim hMem As Long, lPtr As Long
Dim lErr As Long

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = BITMAPTYPE
      .bfOffBits = 14 + Len(m_tBI.bmiHeader)
      .bfSize = .bfOffBits + m_tBI.bmiHeader.biSizeImage
   End With
   hFile = CreateFile(sFileName, _
                 GENERIC_READ Or GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  0)
   lErr = ERR.LastDllError
   If (hFile = INVALID_HANDLE_VALUE) Then
      ' error
      ERR.Raise 17, App.exename & ".cDIBSection", ApiError(lErr)
   Else
      
      ' Writing the BITMAPFILEINFOHEADER is somewhat painful
      ' due to non-byte alignment of structure...
      hMem = GlobalAlloc(GPTR, 14)
      lPtr = GlobalLock(hMem)
      CopyMemory ByVal lPtr, tBH.bfType, 2
      CopyMemory ByVal lPtr + 2, tBH.bfSize, 4
      CopyMemory ByVal lPtr + 6, 0&, 4
      CopyMemory ByVal lPtr + 10, tBH.bfOffBits, 4
      lSize = 14
      lR = WriteFile(hFile, ByVal lPtr, lSize, lBytesWritten, ByVal 0&)
      GlobalUnlock hMem
      GlobalFree hMem
      
      ' Add the BITMAPINFOHEADER:
      bErr = FileErrHandler(lR, lSize, lBytesWritten)
      If Not bErr Then
         lSize = Len(m_tBI.bmiHeader)
         lR = WriteFile(hFile, m_tBI.bmiHeader, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      ' There is no palette for a truecolour DIB
      
      If Not bErr Then
         ' Its easy to write the bitmap data, though...
         lSize = m_tBI.bmiHeader.biSizeImage
         lR = WriteFile(hFile, ByVal m_lPtr, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      
      CloseHandle hFile
      'SavePicture = Not (bErr)
   End If

End Function

Function ApiError(ByVal E As Long) As String
    Dim s As String, C As Long
    s = String(256, 0)
    C = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, E, 0&, s, Len(s), ByVal 0)
    If C Then ApiError = Left$(s, C)
End Function

Private Function FileErrHandler(ByVal lR As Long, ByVal lSize As Long, ByVal lBytes As Long) As Boolean
   If (lR = 0) Or Not (lSize = lBytes) Then
      'Err.Raise
      FileErrHandler = True
   End If
End Function


Private Sub Class_Terminate()
    ClearUp
End Sub

' Scans must align on dword boundaries:
Public Property Get BytesPerScanLine() As Long
    Dim lngNum As Long
    
    If m_tBI.bmiHeader.biWidth <= 0 Then
       BytesPerScanLine = 0
       Exit Function
    End If
    'BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
    BytesPerScanLine = ((m_tBI.bmiHeader.biWidth * m_tBI.bmiHeader.biBitCount + 31) \ 32) * 4
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC

End Property


Public Property Get ColorBit() As Long
    ColorBit = m_tBI.bmiHeader.biBitCount
End Property

