VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'»æÖÆÔ²ÐÎ±ýÍ¼
Private Const OUT_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0
Private Const DEFAULT_CHARSET = 1
Private Const LF_FACESIZE = 32

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type Segment
    Value As Double
    Colour As Long
    name As String
End Type

Private Type LOGFONT
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
        lfFaceName(1 To LF_FACESIZE) As Byte
End Type

Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal y3 As Long, ByVal x4 As Long, ByVal y4 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private LF As LOGFONT

Private intSegments As Long

Private PieChart() As Segment

Private Sub DrawLegend(ByVal lngColour, ByVal intYPosition, ByRef lngPichDC As Long, ByRef lngPichwnd As Long)
    Dim FontToUse As Long
    Dim rc As RECT
    Dim Oldhdc As Long
    Dim Dl As Long
    Dim LnghBrush As Long
    Dim TmpString As String
    
    On Error GoTo ErrHandle
    

    LnghBrush = CreateSolidBrush(lngColour)
    SelectObject lngPichDC, LnghBrush
    
    Rectangle lngPichDC, 220, 50 + (15 * intYPosition), 240, 60 + (15 * intYPosition)
    
    LF.lfHeight = 12: LF.lfWidth = 5: LF.lfEscapement = 0: LF.lfWeight = 800
    LF.lfItalic = 0: LF.lfUnderline = 0: LF.lfStrikeOut = 0
    LF.lfOutPrecision = OUT_DEFAULT_PRECIS: LF.lfClipPrecision = OUT_DEFAULT_PRECIS
    LF.lfQuality = DEFAULT_QUALITY: LF.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
    LF.lfCharSet = DEFAULT_CHARSET
    

    FontToUse = CreateFontIndirect(LF)
 
    Oldhdc = SelectObject(lngPichDC, FontToUse)
    Dl = GetClientRect(lngPichwnd, rc)
    TmpString = PieChart(intYPosition).name & " (" & PieChart(intYPosition).Value & ")"
    Dl = TextOut(lngPichDC, 245, 52 + (15 * intYPosition), TmpString, LenB(StrConv(TmpString, vbFromUnicode)))
    '****
    Dl = TextOut(lngPichDC, 245, 40, "", Len(""))
    '****
    Dl = SelectObject(lngPichDC, Oldhdc)

Exit Sub

ErrHandle:
    Exit Sub
End Sub

Private Sub DrawSegment(ByVal intPerc As Integer, ByVal lngColour As Long, ByRef lngPichDC As Long)
    Const PI = 3.14159265358979
    Dim X1, y1, X2, y2, X3, y3, x4, y4, rad, theta, beta As Double
    Dim LnghBrush As Long

    On Error GoTo ErrHandle
    
    X1 = 0: y1 = 20: X2 = 200: y2 = 220
    X3 = (X2 - X1) / 2: y4 = y1
    rad = (X2 - X1) / 2
    theta = (intPerc / 100) * 360
    beta = 180 - theta - 90
    x4 = rad + ((rad * (SIN(theta * (PI / 180)))) * 180 / PI)
    y4 = rad - ((rad * (SIN(beta * (PI / 180)))) * 180 / PI)

    LnghBrush = CreateSolidBrush(lngColour)
    SelectObject lngPichDC, LnghBrush
    
    Pie lngPichDC, CLng(X1), CLng(y1), CLng(X2), CLng(y2), CLng(x4), CLng(y4), CLng(X3), CLng(y3)
    
Exit Sub

ErrHandle:
    Exit Sub
End Sub

Public Function DrawPie(lngPichDC As Long, lngPichwnd As Long, blnShowLegend, strGraphTitle As String) As Long
    Dim PieTotal As Double
    Dim intCount As Integer
    Dim SegmentTotal As Double
    Dim FontToUse As Long
    Dim rc As RECT
    Dim Oldhdc As Long
    Dim Dl As Long
    
    On Error GoTo ErrHandle

    PieTotal = 0
    For intCount = 0 To UBound(PieChart): PieTotal = PieTotal + PieChart(intCount).Value: Next

    SegmentTotal = 100
    If blnShowLegend Then DrawLegend PieChart(0).Colour, 0, lngPichDC, lngPichwnd
    DrawSegment SegmentTotal, PieChart(0).Colour, lngPichDC

    For intCount = 0 To UBound(PieChart) - 1
        SegmentTotal = SegmentTotal - ((PieChart(intCount).Value / PieTotal) * 100)

        If blnShowLegend Then DrawLegend PieChart(intCount + 1).Colour, intCount + 1, lngPichDC, lngPichDC

        DrawSegment SegmentTotal, PieChart(intCount + 1).Colour, lngPichDC
    Next

    LF.lfHeight = 20: LF.lfWidth = 10: LF.lfEscapement = 0: LF.lfWeight = 800
    LF.lfItalic = 0: LF.lfUnderline = 1: LF.lfStrikeOut = 0
    LF.lfOutPrecision = OUT_DEFAULT_PRECIS: LF.lfClipPrecision = OUT_DEFAULT_PRECIS
    LF.lfQuality = DEFAULT_QUALITY: LF.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
    LF.lfCharSet = DEFAULT_CHARSET

    FontToUse = CreateFontIndirect(LF)

    Oldhdc = SelectObject(lngPichDC, FontToUse)
    Dl = GetClientRect(lngPichwnd, rc)
    Dl = TextOut(lngPichDC, 20, 0, strGraphTitle, LenB(StrConv(strGraphTitle, vbFromUnicode)))
    Dl = SelectObject(lngPichDC, Oldhdc)
    
    DrawPie = 1
Exit Function

ErrHandle: '
    DrawPie = 2
    Exit Function
End Function

Public Function AddSegment(dblPercentage As Double, StrName As String, lngColour As Long) As Long
    intSegments = intSegments + 1
    If intSegments > 0 Then

        ReDim Preserve PieChart(intSegments)
    End If

    PieChart(intSegments).Value = CLng(dblPercentage)
    PieChart(intSegments).Colour = lngColour
    PieChart(intSegments).name = StrName
End Function

Public Function Clear()

    intSegments = -1

    ReDim PieChart(0)
End Function

Private Sub Class_Initialize()
    Clear
End Sub



