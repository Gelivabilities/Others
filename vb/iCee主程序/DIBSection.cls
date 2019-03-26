VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSPICDIBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit
'Ðý×ªÍ¼Ïñ
Private m_hDIB As Long
Private m_hBmpOld As Long
Private m_hDC As Long
Private m_lPtr As Long
Private m_tBI As BITMAPINFO
Public Property Get BytesPerScanLine() As Long
    Dim lngNum As Long
    
    If m_tBI.bmiHeader.biWidth <= 0 Then
       BytesPerScanLine = 0
       Exit Function
    End If
    BytesPerScanLine = ((m_tBI.bmiHeader.biWidth * m_tBI.bmiHeader.biBitCount + 31) \ 32) * 4
End Property

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            Call SelectObject(m_hDC, m_hBmpOld)
            Call DeleteObject(m_hDIB)
        End If
        Call DeleteDC(m_hDC)
    End If
    m_hDC = 0: m_hDIB = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Public Property Get ColorBit() As Long
    ColorBit = m_tBI.bmiHeader.biBitCount
End Property

Public Function CopyToClipboard(Optional ByVal lSrcLeft As Long = 0, _
                                Optional ByVal lSrcTop As Long = 0, _
                                Optional ByVal lSrcWidth As Long = -1, _
                                Optional ByVal lSrcHeight As Long = -1, _
                                Optional ByVal bAsDIB As Boolean = False) As Boolean
    Dim lhDCDesktop As Long
    Dim lhDC As Long
    Dim lhBmpOld As Long
    Dim hObj As Long
    Dim lFmt As Long
    Dim b() As Byte
    Dim tBI As BITMAPINFO
    Dim lPtr As Long
    Dim hDibCopy As Long
    
    If (lSrcWidth < 0) Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If (lSrcHeight < 0) Then lSrcHeight = m_tBI.bmiHeader.biHeight
    
    lhDCDesktop = GetDC(GetDesktopWindow())
    If (lhDCDesktop <> 0) Then
        lhDC = CreateCompatibleDC(0)
        If (lhDC <> 0) Then
            If (bAsDIB) Then
               'MsgBox "I don't know how to put a DIB on the clipboard! Copy as bitmap instead!!!"
            Else
                ' Create a compatible bitmap and copy to
                ' the clipboard:
                hObj = CreateCompatibleBitmap(lhDCDesktop, lSrcWidth, lSrcHeight)
                If (hObj <> 0) Then
                    lhBmpOld = SelectObject(lhDC, hObj)
                    Me.PaintPicture lhDC, APIBitBlt, 0, 0, lSrcWidth, lSrcHeight, lSrcLeft, lSrcTop
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
        DeleteDC lhDCDesktop
    End If
    CopyToClipboard = True
End Function

Public Function Create(ByVal Width As Long, ByVal Height As Long) As Boolean
    Call ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
       If CreateDIB(m_hDC, Width, Height, m_hDIB) Then
          m_hBmpOld = SelectObject(m_hDC, m_hDIB)
          Create = True
       Else
          DeleteDC m_hDC
          m_hDC = 0
       End If
    End If
End Function

Public Function CreateDIB(ByVal lhDC As Long, _
                          ByVal Width As Long, _
                          ByVal Height As Long, _
                          ByRef hDib As Long) As Boolean
    Dim lngI              As Long
    
    With m_tBI.bmiHeader
        .biBitCount = 32
        .biCompression = BI_RGB
        .biClrImportant = 0
        .biClrUsed = 0
        .biHeight = Height
        .biPlanes = 1
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = Width
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biSizeImage = Me.BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
    CreateDIB = (hDib <> 0)
End Function

Public Function CreateFromFile(ByVal sFileName As String) As Boolean
    Dim hMap                   As Long
    Dim TempPic                As IPictureDisp
    
    If Not ExistFile(sFileName) Then Exit Function
    
    hMap = LoadImage(0, sFileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_CREATEDIBSECTION Or LR_LOADFROMFILE)
    If hMap Then
        '
    Else
        On Error Resume Next
        Set TempPic = LoadPicture(sFileName)
        On Error GoTo 0
        If TempPic Is Nothing Then
        Else
            If TempPic.type = vbPicTypeBitmap Then
                hMap = TempPic.handle
            Else
                Set TempPic = Nothing
            End If
        End If
    End If
    
    If hMap Then CreateFromFile = CreateFromHandle(hMap)
    If hMap <> 0 Then Call DeleteObject(hMap)
End Function

Public Function CreateFromHandle(ByVal hBmp As Long, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1) As Boolean
    Dim lhDC                     As Long
    Dim lhBmpOld                 As Long
    Dim tBMP                     As BITMAP
    Dim blnOK                    As Boolean
    
    If lSrcWidth < 0 Or lSrcHeight < 0 Then
       GetObjectAPI hBmp, Len(tBMP), tBMP
       blnOK = Create(tBMP.bmWidth, tBMP.bmHeight)
    Else
       blnOK = Create(lSrcWidth, lSrcHeight)
    End If
    
    If blnOK Then
       lhDC = CreateCompatibleDC(0)
       If (lhDC <> 0) Then
           lhBmpOld = SelectObject(lhDC, hBmp)
           Call LoadPictureBlt(lhDC, 0, 0, lSrcWidth, lSrcHeight, lSrcLeft, lSrcTop, vbSrcCopy)
           Call SelectObject(lhDC, lhBmpOld)
           Call DeleteDC(lhDC)
           CreateFromHandle = True
       End If
    End If
End Function

Public Function CreateFromImage(ByRef hImage As Long, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1) As Boolean
    Dim lhDC                    As Long
    Dim lhBmpOld                As Long
    Dim tBMP                    As BITMAP
    Dim blnOK                   As Boolean
    Dim byBGR()                 As Byte
    Dim tBmi                    As BITMAPINFO
    
    If lSrcWidth < 0 Or lSrcHeight < 0 Then
       Call GetObjectAPI(hImage, Len(tBMP), tBMP)
       blnOK = Create(tBMP.bmWidth, tBMP.bmHeight)
    Else
       blnOK = Create(lSrcWidth, lSrcHeight)
    End If
    
    If blnOK Then
       lhDC = CreateCompatibleDC(0)
       If (lhDC <> 0) Then
           lhBmpOld = SelectObject(lhDC, hImage)
           With tBmi.bmiHeader
                .biBitCount = 32
                .biClrImportant = 0
                .biClrUsed = 0
                .biCompression = BI_RGB
                .biHeight = m_tBI.bmiHeader.biHeight
                .biPlanes = 1
                .biSize = Len(tBmi.bmiHeader)
                .biWidth = m_tBI.bmiHeader.biWidth
                .biXPelsPerMeter = 0
                .biYPelsPerMeter = 0
                .biSizeImage = (((.biWidth * 3) + 3) And &HFFFFFFFC) * .biHeight
           End With
           ReDim byBGR(1 To 4, 1 To m_tBI.bmiHeader.biWidth, 1 To m_tBI.bmiHeader.biHeight)
           Call GetDIBits(lhDC, hImage, 0, m_tBI.bmiHeader.biHeight, byBGR(1, 1, 1), tBmi, DIB_RGB_COLORS)
           Call SetDIBits(m_hDC, m_hDIB, 0, m_tBI.bmiHeader.biHeight, byBGR(1, 1, 1), tBmi, DIB_RGB_COLORS)
           Call SelectObject(lhDC, lhBmpOld)
           Call DeleteDC(lhDC)
           Erase byBGR
           CreateFromImage = True
       End If
    End If
End Function

Public Function CreateFromPackedDIBPointer(ByVal pDIB As Long) As Boolean
    Dim blnOK As Boolean
    
    Debug.Assert pDIB <> 0
    
    'Creates a full-color (no palette) DIB from a pointer to a full-color memory DIB
    
    'get the BitmapInfoHeader
    Call CopyMemory(ByVal VarPtr(m_tBI.bmiHeader.biSize), ByVal pDIB, Len(m_tBI.bmiHeader))
    If m_tBI.bmiHeader.biBitCount < 16 Then
       Debug.Print "Error! DIB was less than 16 colors."
       Exit Function 'only supports high-color or full-color dibs
    End If
    
    blnOK = Create(m_tBI.bmiHeader.biWidth, m_tBI.bmiHeader.biHeight)
    
    blnOK = SetDIBits(m_hDC, m_hDIB, 0, m_tBI.bmiHeader.biHeight, ByVal (pDIB + 40), m_tBI, DIB_RGB_COLORS)
    
    'now get the bitmap bits
    If m_tBI.bmiHeader.biSizeImage < 1 Then Exit Function 'return False
'    m_lPtr = pDIB + 40
    'and return True
    CreateFromPackedDIBPointer = blnOK
End Function

Public Function CreateFromPicture(ByRef picThis As StdPicture, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1) As Boolean
    CreateFromPicture = CreateFromHandle(picThis.handle, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight)
End Function

Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property

Public Property Get DataSize() As Long
    DataSize = m_tBI.bmiHeader.biSizeImage
End Property

Private Function ExistFile(ByVal filename As String) As Boolean
    On Error Resume Next
    Call FileLen(filename)
    ExistFile = (ERR = 0)
End Function

Private Function FileErrHandler(ByVal lR As Long, ByVal lSize As Long, ByVal lBytes As Long) As Boolean
   If (lR = 0) Or Not (lSize = lBytes) Then
      'Err.Raise
      FileErrHandler = True
   End If
End Function

Public Sub GetBestFitInfo(ByVal DstW As Long, ByVal DstH As Long, bfx As Long, bfy As Long, bfW As Long, bfH As Long, Optional ByVal StretchFit As Boolean = 0)
    Dim cW As Single
    Dim ch As Single
    
    bfx = 0
    bfy = 0
    If ((Me.Width > DstW Or Me.Height > DstH) Or StretchFit) Then
       cW = DstW / Me.Width
       ch = DstH / Me.Height
       If (cW < ch) Then
          bfW = DstW
          bfH = Me.Height * cW
       Else
          bfH = DstH
          bfW = Me.Width * ch
       End If
    Else
       bfW = Me.Width
       bfH = Me.Height
    End If
    bfx = (DstW - bfW) \ 2
    bfy = (DstH - bfH) \ 2
End Sub

Public Property Get hdc() As Long
    hdc = m_hDC
End Property

Public Property Get hDib() As Long
    hDib = m_hDIB
End Property

Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt(ByVal lSrcDC As Long, _
                          Optional ByVal lDestLeft As Long = 0, _
                          Optional ByVal lDestTop As Long = 0, _
                          Optional ByVal lDestWidth As Long = -1, _
                          Optional ByVal lDestHeight As Long = -1, _
                          Optional ByVal lSrcLeft As Long = 0, _
                          Optional ByVal lSrcTop As Long = 0, _
                          Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    If lDestWidth < 0 Then lDestWidth = m_tBI.bmiHeader.biWidth
    If lDestHeight < 0 Then lDestHeight = m_tBI.bmiHeader.biHeight
    Call BitBlt(m_hDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, lSrcDC, lSrcLeft, lSrcTop, eRop)
End Sub

Public Function PaintPicture(ByVal lhDC As Long, _
                             Optional ByVal APIProc As UseAPIPaintPicture = APIBitBlt, _
                             Optional ByVal lDestLeft As Long = 0, _
                             Optional ByVal lDestTop As Long = 0, _
                             Optional ByVal lDestWidth As Long = -1, _
                             Optional ByVal lDestHeight As Long = -1, _
                             Optional ByVal lSrcLeft As Long = 0, _
                             Optional ByVal lSrcTop As Long = 0, _
                             Optional ByVal lSrcWidth As Long = -1, _
                             Optional ByVal lSrcHeight As Long = -1, _
                             Optional ByVal eRop As RasterOpConstants = vbSrcCopy) As Boolean
    Dim lOldMode             As Long
    Dim blnOK                As Boolean
'    Dim lTime As Long
    
'    lTime = timeGetTime
    blnOK = False
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    If (lSrcWidth < 0) Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If (lSrcHeight < 0) Then lSrcHeight = m_tBI.bmiHeader.biHeight
    
    Select Case APIProc
           Case APIBitBlt
             blnOK = BitBlt(lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, _
                            m_hDC, lSrcLeft, lSrcTop, eRop)
'             Debug.Print "BitBlt " & timeGetTime - lTime
           Case APISetDIBitsToDevice
             blnOK = SetDIBitsToDevice(ByVal lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, _
                                       lSrcLeft, lSrcTop, 0, m_tBI.bmiHeader.biHeight, ByVal m_lPtr, m_tBI, _
                                       IIf(m_tBI.bmiHeader.biBitCount > 8, DIB_RGB_COLORS, DIB_PAL_COLORS))
'             Debug.Print "SetDIBitsToDevice " & timeGetTime - lTime
           Case APIStretchBlt
             lOldMode = GetStretchBltMode(lhDC)
             Call SetStretchBltMode(lhDC, STRETCH_HALFTONE) 'STRETCH_DELETESCANS
             blnOK = StretchBlt(lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, _
                                m_hDC, lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight, eRop)
             Call SetStretchBltMode(lhDC, lOldMode)
'             Debug.Print "StretchBlt " & timeGetTime - lTime
           Case APIStretchDIBits
             lOldMode = GetStretchBltMode(lhDC)
             Call SetStretchBltMode(lhDC, STRETCH_HALFTONE)
             blnOK = StretchDIBits(ByVal lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, _
                                   lSrcLeft, lSrcTop, lSrcWidth, lSrcHeight, ByVal m_lPtr, m_tBI, _
                                   IIf(m_tBI.bmiHeader.biBitCount > 8, DIB_RGB_COLORS, DIB_PAL_COLORS), eRop)
             Call SetStretchBltMode(lhDC, lOldMode)
'             Debug.Print "StretchDIBits " & timeGetTime - lTime
    End Select
    PaintPicture = blnOK
End Function

Public Function SaveBitmap(ByVal sFileName As String) As Boolean
    Dim hFile              As Long
    Dim lBytesWritten      As Long
    Dim lSize              As Long
    Dim lR                 As Long
    Dim hMem               As Long
    Dim lPtr               As Long
    Dim lErr               As Long
    Dim bErr               As Boolean
    Dim tBH                As BITMAPFILEHEADER

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = Bmp_MAGIC_COOKIE
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

   Else
      
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

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property

Private Sub Class_Initialize()
    'load all those mutthas...
    Call ClearUp
End Sub

Private Sub Class_Terminate()
    Call ClearUp
End Sub




