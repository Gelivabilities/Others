VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit
'Í¼ÏñÆØ¹âÂÊ
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbAlpha As Byte
End Type
Private Type BITMAPINFOHEADER
        bmSize As Long
        bmWidth As Long
        bmHeight As Long
        bmPlanes As Integer
        bmBitCount As Integer
        bmCompression As Long
        bmSizeImage As Long
        bmXPelsPerMeter As Long
        bmYPelsPerMeter As Long
        bmClrUsed As Long
        bmClrImportant As Long
End Type
Private Type BITMAPINFO
        bmHeader As BITMAPINFOHEADER
        bmColors(0 To 255) As RGBQUAD
End Type
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Function GetImageWidth(SrcPictureBox As PictureBox) As Long
    Dim bm As BITMAP
    GetObject SrcPictureBox.image, Len(bm), bm
    GetImageWidth = bm.bmWidth
End Function
Public Function GetImageHeight(SrcPictureBox As PictureBox) As Long
    Dim bm As BITMAP
    GetObject SrcPictureBox.image, Len(bm), bm
    GetImageHeight = bm.bmHeight
End Function
Public Function GetImageStreamLength(SrcPictureBox As PictureBox) As Long
    Dim bm As BITMAP
    GetObject SrcPictureBox.image, Len(bm), bm
    GetImageStreamLength = (bm.bmWidth * (bm.bmHeight + 1)) * 3
End Function
Public Sub GetImageData2D(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As BITMAP
    GetObject SrcPictureBox.image, Len(bm), bm
    Erase ImageData()
    Dim ArrayWidth As Long
    ArrayWidth = (bm.bmWidth * 3) - 1
    ArrayWidth = ArrayWidth + (bm.bmWidth Mod 4)
    ReDim ImageData(0 To ArrayWidth, 0 To bm.bmHeight) As Byte
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40
    bmi.bmHeader.bmPlanes = 1
    bmi.bmHeader.bmBitCount = 24
    bmi.bmHeader.bmCompression = 0
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.image, 0, bm.bmHeight, ImageData(0, 0), bmi, 0

End Sub

Public Sub SetImageData2D(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As BITMAP
    GetObject DstPictureBox.image, Len(bm), bm
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40
    bmi.bmHeader.bmPlanes = 1
    bmi.bmHeader.bmBitCount = 24
    bmi.bmHeader.bmCompression = 0
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0, 0), bmi, 0, vbSrcCopy
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.image
        DstPictureBox.Refresh
    End If
    DoEvents
End Sub
Public Sub GetImageData(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As BITMAP
    GetObject SrcPictureBox.image, Len(bm), bm
    Erase ImageData()
    ReDim ImageData(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40
    bmi.bmHeader.bmPlanes = 1
    bmi.bmHeader.bmBitCount = 24
    bmi.bmHeader.bmCompression = 0
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.image, 0, bm.bmHeight, ImageData(0, 0, 0), bmi, 0
End Sub
Public Sub SetImageData(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As BITMAP
    GetObject DstPictureBox.image, Len(bm), bm
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40
    bmi.bmHeader.bmPlanes = 1
    bmi.bmHeader.bmBitCount = 24
    bmi.bmHeader.bmCompression = 0
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0, 0, 0), bmi, 0, vbSrcCopy
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.image
        DstPictureBox.Refresh
    End If
    DoEvents
End Sub

Public Sub GetImageDataStream(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As BITMAP
    GetObject SrcPictureBox.image, Len(bm), bm
    Erase ImageData()
    ReDim ImageData(0 To GetImageStreamLength(SrcPictureBox))
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40
    bmi.bmHeader.bmPlanes = 1
    bmi.bmHeader.bmBitCount = 24
    bmi.bmHeader.bmCompression = 0
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.image, 0, bm.bmHeight, ImageData(0), bmi, 0
End Sub

Public Sub SetImageDataStream(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As BITMAP
    GetObject DstPictureBox.image, Len(bm), bm
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40
    bmi.bmHeader.bmPlanes = 1
    bmi.bmHeader.bmBitCount = 24
    bmi.bmHeader.bmCompression = 0
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0), bmi, 0, vbSrcCopy
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.image
        DstPictureBox.Refresh
    End If
    DoEvents
End Sub


