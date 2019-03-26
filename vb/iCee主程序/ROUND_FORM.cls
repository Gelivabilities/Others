VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ROUND_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Ô²ÐÎ´°¿Ú»òÍ¼Æ¬¿ò
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private tipRC As RECT
Private TipBox As RECT
Private mlTipBox As Long
Private hBrush As Long

Public Function ROUND_USER(frmobj As UserControl, cornertwist As Integer, verticalborderwidth As Double, horizontalborderwidth As Double)
    Dim lRet As Long
    Dim iDrawTxt As Integer, sHelp As String
    Dim lTipWidth As Long, lTipHeight As Long
    Dim lCorner As Long
    lCorner = cornertwist
    lTipWidth = frmobj.ScaleWidth
    lTipHeight = frmobj.ScaleHeight
    mlTipBox = CreateRoundRectRgn(0, 0, lTipWidth, lTipHeight, lCorner, lCorner)
    hBrush = CreateSolidBrush(&H404040)
    lRet = FrameRgn(frmobj.hdc, mlTipBox, hBrush, verticalborderwidth, horizontalborderwidth)
    lRet = SetWindowRgn(frmobj.hwnd, mlTipBox, True)
    frmobj.Refresh
End Function

Public Function ROUND_FORM(frmobj As Form, cornertwist As Integer, verticalborderwidth As Double, horizontalborderwidth As Double)
    Dim lRet As Long
    Dim iDrawTxt As Integer, sHelp As String
    Dim lTipWidth As Long, lTipHeight As Long
    Dim lCorner As Long
    lCorner = cornertwist
    lTipWidth = frmobj.ScaleWidth
    lTipHeight = frmobj.ScaleHeight
    mlTipBox = CreateRoundRectRgn(0, 0, lTipWidth, lTipHeight, lCorner, lCorner)
    hBrush = CreateSolidBrush(&H404040)
    lRet = FrameRgn(frmobj.hdc, mlTipBox, hBrush, verticalborderwidth, horizontalborderwidth)
    lRet = SetWindowRgn(frmobj.hwnd, mlTipBox, True)
    frmobj.Refresh
End Function
Public Function ROUND_PIC(PIC As PictureBox, cornertwist As Integer, verticalborderwidth As Double, horizontalborderwidth As Double)
    Dim lRet As Long
    Dim iDrawTxt As Integer, sHelp As String
    Dim lTipWidth As Long, lTipHeight As Long
    Dim lCorner As Long
    lCorner = cornertwist
    lTipWidth = PIC.ScaleWidth
    lTipHeight = PIC.ScaleHeight
    mlTipBox = CreateRoundRectRgn(0, 0, lTipWidth, lTipHeight, lCorner, lCorner)
    'hBrush = CreateSolidBrush(&H800000)
    lRet = FrameRgn(PIC.hdc, mlTipBox, hBrush, verticalborderwidth, horizontalborderwidth)
    lRet = SetWindowRgn(PIC.hwnd, mlTipBox, True)
    PIC.Refresh
End Function

