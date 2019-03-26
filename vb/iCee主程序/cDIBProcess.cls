VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'旋转图像进度条
Public Event InitProgress(ByVal Max As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)
Private Type OLECOLOR
    RedOrSys As Byte
    Green As Byte
    Blue As Byte
    type As Byte
End Type

Private Function GetB(ByRef Color As Long) As Byte
    Dim clr As OLECOLOR
    
    CopyMemory clr, WinColor(Color), Len(clr)
    
    GetB = clr.Blue
End Function

Private Function GetG(ByRef Color As Long) As Byte
    Dim clr As OLECOLOR
    
    CopyMemory clr, WinColor(Color), Len(clr)
    
    GetG = clr.Green
End Function

Private Function GetR(ByRef Color As Long) As Byte
    Dim clr As OLECOLOR
    
    CopyMemory clr, WinColor(Color), Len(clr)
    
    GetR = clr.RedOrSys
End Function

Private Sub pvBuildSA(ByRef tSA As SAFEARRAY2D, ByRef DIB As CLSPICDIBS)
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.Width
        .pvData = DIB.DIBSectionBitsPtr
    End With
End Sub

Private Function WinColor(ByRef VBColor As Long) As Long
    'Get a windows color from a VB color
    'Important for system colors like vbButtonFace.
    '
    Dim SysClr As OLECOLOR
    
    CopyMemory SysClr, VBColor, Len(SysClr)
    
    If SysClr.type = &H80 Then 'It is a system color
        'SysClr.RedOrSys is the index of the system color
        WinColor = GetSysColor(SysClr.RedOrSys)
    Else
        WinColor = VBColor
    End If
End Function

Private Function Maximum(ParamArray Vals())
    Dim n As Integer, MaxVal
    
    For n = 0 To UBound(Vals)
        If Vals(n) > MaxVal Then MaxVal = Vals(n)
    Next n

    Maximum = MaxVal
End Function

Private Function Minimum(ParamArray Vals())
    Dim n As Integer, MinVal
    
    MinVal = Vals(0)
    
    For n = 0 To UBound(Vals)
        If Vals(n) < MinVal Then MinVal = Vals(n)
    Next n

    Minimum = MinVal
End Function

Public Sub RotateDIB(DIB As CLSPICDIBS, Angle As Single, Optional BackClr As Long = vbBlack)
    Dim nWidth As Long, nHeight As Long, Cosine As Single
    Dim Sine As Single, X1 As Long, y1 As Long, X2 As Long
    Dim y2 As Long, X3 As Long, y3 As Long, MinX As Long
    Dim MinY As Long, MaxX As Long, MaxY As Long
    Dim X As Long, Y As Long, SourceX As Long, SourceY As Long
    Dim lngWidth As Long, lngHeight As Long
    Dim sBits() As RGBQUAD, dBits() As RGBQUAD, BackBit As RGBQUAD
    Dim stSA As SAFEARRAY2D, dtSA As SAFEARRAY2D
    Dim lTime As Long
    Dim sDib As New CLSPICDIBS
    
    lTime = timeGetTime()
    Set sDib = DIB
    Set DIB = Nothing
    Angle = Angle * ((4 * Atn(1)) / -180)
    
    nWidth = sDib.Width - 1
    nHeight = sDib.Height - 1

    Cosine = Cos(Angle): Sine = SIN(Angle)
    X1 = -nHeight * Sine
    y1 = nHeight * Cosine
    X2 = nWidth * Cosine - nHeight * Sine
    y2 = nHeight * Cosine + nWidth * Sine
    X3 = nWidth * Cosine
    y3 = nWidth * Sine

    MinX = Minimum(0, Minimum(X1, Minimum(X2, X3)))
    MinY = Minimum(0, Minimum(y1, Minimum(y2, y3)))
    MaxX = Maximum(X1, Maximum(X2, X3))
    MaxY = Maximum(y1, Maximum(y2, y3))

    lngWidth = MaxX - MinX: lngHeight = MaxY - MinY
    RaiseEvent InitProgress(lngHeight - 1)
    Set DIB = New CLSPICDIBS
    DIB.Create lngWidth, lngHeight
    Call pvBuildSA(stSA, sDib)
    CopyMemory ByVal VarPtrArray(sBits()), VarPtr(stSA), 4
    Call pvBuildSA(dtSA, DIB)
    CopyMemory ByVal VarPtrArray(dBits()), VarPtr(dtSA), 4
    With BackBit
         .rgbBlue = GetB(BackClr)
         .rgbGreen = GetG(BackClr)
         .rgbRed = GetR(BackClr)
    End With
    For Y = 0 To lngHeight - 1
        For X = 0 To lngWidth - 1
            SourceX = (X + MinX) * Cosine + (Y + MinY) * Sine
            SourceY = (Y + MinY) * Cosine - (X + MinX) * Sine
            If (SourceX >= 0) And (SourceX <= nWidth) And _
               (SourceY >= 0) And (SourceY <= nHeight) Then
               LSet dBits(X, Y) = sBits(SourceX, SourceY)
            Else
               LSet dBits(X, Y) = BackBit
            End If
        Next X
        RaiseEvent Progress(Y)
    Next Y
    CopyMemory ByVal VarPtrArray(sBits), 0&, 4
    CopyMemory ByVal VarPtrArray(dBits), 0&, 4
    Set sDib = Nothing
    RaiseEvent Complete(timeGetTime - lTime)
End Sub


