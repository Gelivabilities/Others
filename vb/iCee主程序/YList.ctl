VERSION 5.00
Begin VB.UserControl IList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H001F1F1F&
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   237
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   227
   ToolboxBitmap   =   "YList.ctx":0000
   Begin ICEE.IVScroll YVScroll1 
      Height          =   2895
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   255
      _ExtentX        =   661
      _ExtentY        =   5106
      MinV            =   0
      MaxV            =   20
      Value           =   0
      SmallChange     =   1
      LargeChange     =   10
   End
End
Attribute VB_Name = "IList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim COLOR_N As Long, COLOR_H As Long
'�Զ����������
Private m_List() As String            '���ڱ����б����ı��Ķ�̬����
Private m_ListCount As Integer        '�����б������Ŀ
Private m_ItemHeight As Integer       '�б���ĸ߶�
Private m_PageCount As Integer        '�б��һҳ���б���������Ŀ
Private m_OldValue As Integer         '���ڱ����������ֵ
Private m_TopIndex As Integer         '�б���е�һ���б��������(���Ŵ�ֱ�������Ĺ����ᷢ���仯)
Private m_ListIndex As Integer        '�б���е�ǰѡ���������ֵ,���н�����б��������ֵ
Private nMaxLength As Integer         '�б����ı�����󳤶�
Private m_ScrollWidth As Integer      '����ˮƽ������,����ֵ1��ʾ���ҹ���1��ƽ���ַ����

'�¼�����:
Event Click()
Event DBClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
'�Զ��峣������
Private Const nH = 18               '�ʹ�ֱ�������ؼ��Ŀ��ֵ18����һ��
Private Const m_relLeft = 2         '�б�����ʾλ�õ���ƫ����
Private Const m_relTop = 2          '�б�����ʾλ�õ���ƫ����

'ϵͳ��������
Private Const GWL_WNDPROC           As Long = -4

Private Const WM_PAINT = &HF
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MOUSEWHEEL = &H20A

Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400

'ϵͳ��������
Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Private Type Size
        cx As Long
        cy As Long
End Type

Private Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type

Private Type tSubData
    hwnd        As Long
    nAddrSub    As Long
    nAddrOrig   As Long
    nMsgCntA    As Long
    nMsgCntB    As Long
    aMsgTblA()  As Long
    aMsgTblB()  As Long
End Type

Private Enum eMsgWhen
    MSG_AFTER = 1
    MSG_BEFORE = 2
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

Private Const WM_MOUSEMOVE  As Long = &H200

Private sc_aSubData()               As tSubData

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef hwnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    
    Dim PS As PAINTSTRUCT
    Dim hdc As Long
    Dim iIndex As Integer, i As Integer
    Dim rc As RECT
    Dim hBrush As Long, hBrushFrame As Long
    Dim TMP As Integer
    Dim dwX         As Integer
    Dim dwY         As Integer
    Dim ItemIdx     As Integer
    Static f           As Boolean   '��������Maxֵ��Valueֵ�Ƿ��Ѿ���ʼ����
    Dim tmpTopIndex As Integer
    Static IsHScroll As Boolean

    Select Case uMsg
        
        Case WM_PAINT

            '����ˢ��
            hBrush = CreateSolidBrush(COLOR_N) '����Ǳ�����ɫ
            hBrushFrame = CreateSolidBrush(COLOR_H)


            '��������б���,����һҳ���޷������������б���,��ô����ʾ��ֱ������,������ʾ
            If m_ListCount > 0 Then
                '�ж��Ƿ���Ҫ��ʾ������
                If m_ListCount > m_PageCount Then
                    '�������������߶�
                    TMP = UserControl.ScaleHeight - 6 - nH * 2
                    TMP = IIf(IsHScroll, TMP - nH - 1, TMP)
                    TMP = m_PageCount * TMP / m_ListCount
                    TMP = IIf(TMP <= 10, 10, TMP)
                    '��ʼ��������
                    YVScroll1.MaxV = m_ListCount - m_PageCount
                    YVScroll1.GlideHeight = TMP
                    YVScroll1.Visible = True
                Else
                    m_TopIndex = 0
                    YVScroll1.Value = 0
                    YVScroll1.Visible = False
                End If
            End If
  
            hdc = BeginPaint(hwnd, PS)
            
            '���б�򱳾�
            With rc
                .Left = 0
                .Top = 0
                .Right = UserControl.ScaleWidth
                .Bottom = UserControl.ScaleHeight
            End With
            FillRect hdc, rc, hBrush
            
            '��������б���
            If m_ListCount > 0 Then
                TMP = m_TopIndex + m_PageCount - 1
                If TMP > UBound(m_List) Then
                    TMP = UBound(m_List)
                End If
                '������б���
                For iIndex = m_TopIndex To TMP
                    With rc
                        .Left = m_relLeft - m_ScrollWidth
                        .Top = m_ItemHeight * i + m_relTop
                        .Right = UserControl.ScaleWidth - nH - 1 - m_relLeft
                        If YVScroll1.Visible = False Then .Right = UserControl.ScaleWidth - m_relLeft
                        .Bottom = .Top + m_ItemHeight
                    End With
                    '����ǽ�������ô�����߶���ʾ
                    If iIndex = ListIndex Then FillRect hdc, rc, hBrushFrame
                    '���б����ı�
                    DrawText hdc, m_List(iIndex), -1, rc, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
                    i = i + 1
                Next iIndex
            End If
            
            '���߿�
            'GetClientRect hWnd, rc
           ' pFrameRect hdc, rc, hBrushFrame
            
            DeleteObject hBrush
            DeleteObject hBrushFrame

            EndPaint hwnd, PS
            
            bHandled = True
            lReturn = 0
         
        Case WM_LBUTTONDOWN
            '������б����ϵ����ô����Click�¼�
            dwX = LOWORD(lParam)
            dwY = HIWORD(lParam)
            ItemIdx = HitTest(dwX, dwY)
            If ItemIdx = -1 Then Exit Sub
            ListIndex = ItemIdx
            RaiseEvent Click
            
        Case WM_LBUTTONDBLCLK
            RaiseEvent DBClick
        
        Case WM_MOUSEWHEEL
            If m_ListCount = 0 Then Exit Sub
            '����m_TopIndex��ֵ
            tmpTopIndex = m_TopIndex
            If HIWORD(wParam) < 0 Then '���������¹���
                tmpTopIndex = tmpTopIndex + 3
                tmpTopIndex = IIf(tmpTopIndex > UBound(m_List) - m_PageCount + 1, _
                                                       UBound(m_List) - m_PageCount + 1, tmpTopIndex)
                If m_TopIndex <> tmpTopIndex Then
                    m_TopIndex = tmpTopIndex
                    YVScroll1.Value = m_TopIndex
                    InvalidateRect UserControl.hwnd, 0, True
                End If
            Else                       '���Ϲ���
                tmpTopIndex = tmpTopIndex - 3
                tmpTopIndex = IIf(tmpTopIndex < 0, 0, tmpTopIndex)
                If m_TopIndex <> tmpTopIndex Then
                    m_TopIndex = tmpTopIndex
                    YVScroll1.Value = m_TopIndex
                    InvalidateRect UserControl.hwnd, 0, True
                End If
            End If
            
            bHandled = True
            lReturn = 0
        
        'Case WM_LBUTTONUP

       'Case WM_MOUSEMOVE
        'Case WM_MOUSELEAVE

    End Select
    
    '�����������������Ϣ
    'bHandled = True
    'lReturn = 0
End Sub

Private Function HIWORD(ByVal lParam As Long) As Integer
    HIWORD = lParam \ 65536
End Function
    
Private Function LOWORD(ByVal lParam As Long) As Integer
    LOWORD = lParam Mod 65536
End Function

'�жϵ�ǰ�������λ�ö�Ӧ�б��������ֵ�������б������򷵻�-1
Private Function HitTest(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim rc As RECT
    Dim iCount      As Integer
    Dim iHitItem    As Integer
    
    If ListCount = 0 Then
        HitTest = -1
        Exit Function
    End If
    
    If X < m_relLeft Or X > UserControl.ScaleWidth - nH - 1 - m_relLeft _
                               Or Y < m_relTop Or Y > UserControl.ScaleHeight - 1 - m_relTop Then
        HitTest = -1
        Exit Function
    End If
    
    iHitItem = m_TopIndex + Fix((Y - m_relTop) / m_ItemHeight)                  '�õ���ǰҳ��������б��������ֵ
    If iHitItem < 0 Or iHitItem > ListCount - 1 Then iHitItem = -1
    HitTest = iHitItem
    
End Function

Private Sub UserControl_Initialize()

    m_ListIndex = -1
    '��ʼ���б���ĸ߶�
    m_ItemHeight = 20
    '��ʼ����ֱ���������й�����
    YVScroll1.Visible = False
    '��ʼ��ˮƽ���������й�����
    COLOR_N = &H1F1F1F
    COLOR_H = &HCDC034
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  'Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Font)
  m_ItemHeight = 18
      Call Subclass_Start(hwnd)
      Call Subclass_AddMsg(hwnd, WM_PAINT, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_LBUTTONDOWN, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_LBUTTONDBLCLK, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_MOUSEWHEEL, MSG_BEFORE)
      'Call Subclass_AddMsg(hWnd, WM_LBUTTONUP, MSG_BEFORE)
      'Call Subclass_AddMsg(hWnd, WM_RBUTTONUP, MSG_BEFORE)
      'Call Subclass_AddMsg(hWnd, WM_ERASEBKGND, MSG_BEFORE)
      'Call Subclass_AddMsg(hWnd, WM_MOUSEMOVE, MSG_BEFORE)
      'Call Subclass_AddMsg(hWnd, WM_MOUSELEAVE, MSG_BEFORE)
      'Call Subclass_AddMsg(hWnd, WM_KILLFOCUS, MSG_BEFORE)
End Sub

Private Sub UserControl_Resize()
    If m_PageCount = 0 Then
        Call InitList
    Else
        YVScroll1.Move UserControl.ScaleWidth - nH - 2, 2, nH, UserControl.ScaleHeight - 4
    End If
End Sub

Private Sub UserControl_Terminate()
  On Error GoTo Catch
  Call Subclass_StopAll   'ֹͣ������Ϣ
  Exit Sub
Catch:
    ERR.Clear
End Sub

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  With sc_aSubData(zIdx(lng_hWnd))
    If (When) And (eMsgWhen.MSG_BEFORE) Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If (When) And (eMsgWhen.MSG_AFTER) Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  With sc_aSubData(zIdx(lng_hWnd))
    If (When) And (eMsgWhen.MSG_BEFORE) Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If (When) And (eMsgWhen.MSG_AFTER) Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
  Const CODE_LEN              As Long = 200
  Const FUNC_CWP              As String = "CallWindowProcA"
  Const FUNC_EBM              As String = "EbMode"
  Const FUNC_SWL              As String = "SetWindowLongA"
  Const MOD_USER              As String = "user32"
  Const MOD_VBA5              As String = "vba5"
  Const MOD_VBA6              As String = "vba6"
  Const PATCH_01              As Long = 18
  Const PATCH_02              As Long = 68
  Const PATCH_03              As Long = 78
  Const PATCH_06              As Long = 116
  Const PATCH_07              As Long = 121
  Const PATCH_0A              As Long = 186
  Static aBuf(1 To CODE_LEN)  As Byte
  Static pCWP                 As Long
  Static pEbMode              As Long
  Static pSWL                 As Long
  Dim i                       As Long
  Dim j                       As Long
  Dim nSubIdx                 As Long
  Dim sHex                    As String
  
  If (aBuf(1) = 0) Then
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
    i = 1
    Do While (j < CODE_LEN)
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
      i = i + 2
    Loop
    If (Subclass_InIDE = True) Then
      aBuf(16) = &H90
      aBuf(17) = &H90
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
      If (pEbMode = 0) Then pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
    End If
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
    ReDim sc_aSubData(0 To 0) As tSubData
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If (nSubIdx = -1) Then
      nSubIdx = UBound(sc_aSubData()) + 1
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
    End If
    Subclass_Start = nSubIdx
  End If
  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd
    .nAddrSub = GlobalAlloc(0, CODE_LEN)
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
  End With
End Function

Private Sub Subclass_StopAll()
  Dim i As Long
  
  On Error GoTo MyErr
  i = UBound(sc_aSubData())
  Do While (i >= 0)
    With sc_aSubData(i)
      If (.hwnd <> 0) Then Call Subclass_Stop(.hwnd)
    End With
    i = i - 1
  Loop
  Exit Sub
MyErr:
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)
    Call zPatchVal(.nAddrSub, 93, 0)
    Call zPatchVal(.nAddrSub, 137, 0)
    Call GlobalFree(.nAddrSub)
    .hwnd = 0
    .nMsgCntB = 0
    .nMsgCntA = 0
    Erase .aMsgTblB
    Erase .aMsgTblA
  End With
End Sub

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long, nOff1 As Long, nOff2 As Long
  
  If (uMsg = -1) Then
    nMsgCnt = -1
  Else
    Do While (nEntry < nMsgCnt)
      nEntry = nEntry + 1
      If (aMsgTbl(nEntry) = 0) Then
        aMsgTbl(nEntry) = uMsg
        Exit Sub
      ElseIf (aMsgTbl(nEntry) = uMsg) Then
        Exit Sub
      End If
    Loop
    nMsgCnt = nMsgCnt + 1
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
    aMsgTbl(nMsgCnt) = uMsg
  End If
  If (When = eMsgWhen.MSG_BEFORE) Then
    nOff1 = 88
    nOff2 = 93
  Else
    nOff1 = 132
    nOff2 = 137
  End If
  If (uMsg <> -1) Then Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
  Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If (uMsg = -1) Then
    nMsgCnt = 0
    If (When = eMsgWhen.MSG_BEFORE) Then
      nEntry = 93
    Else
      nEntry = 137
    End If
    Call zPatchVal(nAddr, nEntry, 0)
  Else
    Do While (nEntry < nMsgCnt)
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then
        aMsgTbl(nEntry) = 0
        Exit Do
      End If
    Loop
  End If
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
  zIdx = UBound(sc_aSubData)
  Do While (zIdx >= 0)
    With sc_aSubData(zIdx)
      If (.hwnd = lng_hWnd) And Not (bAdd = True) Then
        Exit Function
      ElseIf (.hwnd = 0) And (bAdd = True) Then
        Exit Function
      End If
    End With
    zIdx = zIdx - 1
  Loop
  If Not (bAdd = True) Then Debug.Assert False
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'�ڲ�����
Private Sub pFrameRect(ByVal hdc As Long, rc As RECT, ByVal hBrush As Long)
    FrameRect hdc, rc, hBrush
End Sub

'�ڶ���������ʱ����,ʵ��û�õ�
Public Sub AddItem(ByVal Item As String, Optional Index As Integer)
    '�б���������Ŀ�������޵�
    If m_ListCount = 32767 Then
        Exit Sub
    End If
    '���¶����б����ı�������
    ReDim Preserve m_List(m_ListCount)
    '�����б����ı�
    m_List(m_ListCount) = Item
    '�б������Ŀ��1
    m_ListCount = m_ListCount + 1
    '���𴰿��ػ�
    InvalidateRect UserControl.hwnd, 0, True
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
    Dim i As Integer
    Static f As Boolean      '�����ظ���ʼ�����������й�����
    Dim CntOfUp As Integer   'ɾ���б���ǰ�Ϸ��м������������б���
    Dim CntOfDown As Integer 'ɾ���б���ǰ�·��м������������б���
    
    '��������б���
    If m_ListCount > 0 Then
        '�����ɾ�����б�������ֵ��0��UBound(m_List)֮��
        If Index >= 0 And Index <= UBound(m_List) Then
            '����m_List���������
            If Index <> UBound(m_List) Then
                For i = Index To UBound(m_List) - 1
                    m_List(i) = m_List(i + 1)
                Next i
            End If
            '�б������Ŀ��1
            m_ListCount = m_ListCount - 1
            '���m_ListCountΪ0,��ô�ͷ�����
            If m_ListCount = 0 Then
                Erase m_List
                Exit Sub
            Else
                '���¶����б����ı�������
                ReDim Preserve m_List(m_ListCount - 1)
            End If
            '�������Ҫ��ʾ��������ô��ʼ�����������й�����
            If m_ListCount <= m_PageCount Then
                If f = False Then
                    m_TopIndex = 0
                    YVScroll1.Value = 0
                    YVScroll1.Visible = False
                    f = True
                End If
            Else
                f = False
            End If
            '���ݸ���ɾ�����ȷ��m_TopIndex,YVScroll1.Valueֵ,�Ա�ɾ������������ʾ
            '���m_TopIndex=0��ô����Ҫ��ʲô�������������ʾ,ֻҪ����m_TopIndex��Ϊ0�����(ֻҪm_TopIndex
            '��Ϊ0��˵��ɾ���б���󻹴��ڹ�����)
            If m_TopIndex <> 0 Then
                CntOfUp = m_TopIndex
                CntOfDown = m_ListCount - (m_TopIndex + m_PageCount - 1) '��ʱm_ListCount��ֵ��ɾ���б���ǰ
                                                                           '�б���������������ֵ
                If CntOfUp <= CntOfDown Then
                    If Index < m_TopIndex Then
                        m_TopIndex = m_TopIndex - 1
                        YVScroll1.Value = YVScroll1.Value - 1
                    End If
                ElseIf CntOfUp > CntOfDown Then
                    If CntOfDown = 0 Then
                        m_TopIndex = m_TopIndex - 1
                        YVScroll1.Value = YVScroll1.Value - 1
                    Else
                        If Index < m_TopIndex Then
                            m_TopIndex = m_TopIndex - 1
                            YVScroll1.Value = YVScroll1.Value - 1
                        End If
                    End If
                End If
            End If
            'ɾ���б����Խ�����Ĵ���
            If m_ListIndex > UBound(m_List) Then m_ListIndex = -1
            '���𴰿��ػ�
            InvalidateRect UserControl.hwnd, 0, True
        End If
    End If

End Sub

'�õ��б���ĸ߶�
Public Property Get ItemHeight() As Integer
    ItemHeight = m_ItemHeight
End Property

'�����б���ĸ߶�
Public Property Let ItemHeight(ByVal NewValue As Integer)
    Dim n As Integer
    Dim tm As TEXTMETRIC
    '�õ�����ĸ߶�
    GetTextMetrics UserControl.hdc, tm
    If NewValue >= tm.tmHeight Then
        m_ItemHeight = NewValue
        UserControl_Resize
        InvalidateRect UserControl.hwnd, 0, True
    End If
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(NewValue As StdFont)
    Dim tm As TEXTMETRIC
    Set UserControl.Font = NewValue
    '�õ�����ĸ߶�
    GetTextMetrics UserControl.hdc, tm
    If tm.tmHeight > m_ItemHeight Then
        m_ItemHeight = tm.tmHeight
    End If
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", UserControl.Font)
    Call PropBag.WriteProperty("ItemHeight", m_ItemHeight)
End Sub


Public Property Get ListIndex() As Integer
    ListIndex = m_ListIndex
End Property

Public Property Let ListIndex(ByVal va As Integer)
    If (va >= 0 And va <= m_ListCount - 1) Or va = -1 Then
        m_ListIndex = va
        '���𴰿��ػ�
        InvalidateRect UserControl.hwnd, 0, True
    End If
End Property

Public Property Get PageCount() As Integer
    PageCount = m_PageCount
End Property

Public Property Let PageCount(ByVal va As Integer)
    If va > 0 Then
        m_PageCount = va
        YVScroll1.Move UserControl.ScaleWidth - nH - 2, 2, nH, UserControl.ScaleHeight - 4
    End If
End Property

Public Property Get ListCount() As Integer
    ListCount = m_ListCount
End Property

Public Property Get List(ByVal Index As Integer) As String
    If m_ListCount > 0 Then
        If Index >= 0 And Index <= m_ListCount - 1 Then List = m_List(Index)
    End If
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub InitValue()
    YVScroll1.Value = 0
End Sub

'����б������󳤶���ȷ���Ƿ���Ҫˮƽ������,����0��ʾ����Ҫ,��0��ʾ��Ҫ
Public Function GetMaxItemLength() As Integer
    Dim i As Integer
    Dim MaxLength As Integer, TMP As Integer
    Dim lpSize As Size
    
    If m_ListCount = 0 Then Exit Function
    For i = 0 To UBound(m_List)
        GetTextExtentPoint32 UserControl.hdc, m_List(i), LenB(StrConv(m_List(i), vbFromUnicode)), lpSize
        If lpSize.cx > MaxLength Then MaxLength = lpSize.cx
    Next i
    If m_ListCount <= m_PageCount Then
        TMP = UserControl.ScaleWidth - m_relLeft * 2
    Else
        TMP = UserControl.ScaleWidth - nH - 1 - m_relLeft * 2
    End If
    If MaxLength > TMP Then GetMaxItemLength = MaxLength
End Function

'ȷ��ˮƽ��������MaxVֵ�Լ�����Ŀ��,���ı��б��߶�,Ȼ���ƶ�����ȷ��λ��
Public Sub SetHScrollInformation(ByVal MaxLength As Integer)
   Dim rel As Integer
   Dim tm As TEXTMETRIC
   Dim ActLength As Integer, TMP As Integer
   
   '�õ��豸�������ı����й���Ϣ(�����õ���tm.tmAveCharWidth)
   GetTextMetrics UserControl.hdc, tm
   
    If m_ListCount <= m_PageCount Then
        ActLength = UserControl.ScaleWidth - m_relLeft * 2
    Else
        ActLength = UserControl.ScaleWidth - nH - 1 - m_relLeft * 2
    End If
   
   rel = MaxLength - ActLength

   'ȷ������Ŀ��
   TMP = ActLength * (UserControl.ScaleWidth - 9 - nH * 3) / MaxLength
   TMP = IIf(TMP <= 10, 10, TMP)

   '�ı��б��ĸ߶�,Ȼ���ƶ�����ȷ��λ��
   If UserControl.ScaleHeight > m_PageCount * m_ItemHeight + m_relTop * 2 Then
       'ʲô������
   Else
       UserControl.Height = m_PageCount * m_ItemHeight + m_relTop * 2 + 19
   End If
   '����������Ԥ��1���صĿռ����ڻ��б��ı߿�
End Sub

Private Sub YVScroll1_Change()
    m_TopIndex = m_TopIndex + (YVScroll1.Value - YVScroll1.OldValue)
    'Debug.Print "����������ı��¼�,����ֵ�ı���" & YVScroll1.Value - YVScroll1.OldValue & "����λ"
    YVScroll1.OldValue = YVScroll1.Value
    '���𴰿��ػ�
    InvalidateRect UserControl.hwnd, 0, True
End Sub

Private Sub InitList()
    Dim nWidth As Integer
    Dim nHeight As Integer
    Dim TMP As Integer
    
    nWidth = UserControl.ScaleWidth - m_relLeft * 2 - nH - 1
    nHeight = UserControl.ScaleHeight - m_relTop * 2
    
    If nWidth < m_ItemHeight Then
        UserControl.Width = (m_ItemHeight + m_relLeft * 2 + nH + 1) * 15
    End If
    
    If nHeight < m_ItemHeight Then
        UserControl.Height = (m_ItemHeight + m_relTop * 2) * 15
        m_PageCount = 1
    Else
        If nHeight Mod m_ItemHeight Then
            TMP = (nHeight \ m_ItemHeight) + 1
            UserControl.Height = (TMP * m_ItemHeight + m_relTop * 2) * 15
            m_PageCount = TMP
        Else
            m_PageCount = nHeight \ m_ItemHeight
        End If
    End If
  
    YVScroll1.Move UserControl.ScaleWidth - nH - 2, 2, nH, UserControl.ScaleHeight - 4
End Sub

Public Sub Clear()
    Dim i As Integer, j As Integer
    If m_ListCount > 0 Then
        j = m_ListCount - 1
        For i = j To 0 Step -1
            RemoveItem i
        Next i
    End If
End Sub

Public Sub InitListHeight(ByVal iMaxPageCount As Integer)

    If m_ListCount = 0 Then
        UserControl.Height = (m_relTop * 2 + m_ItemHeight) * 15
        PageCount = 1
    ElseIf m_ListCount > 0 And m_ListCount <= iMaxPageCount Then
        UserControl.Height = (m_relTop * 2 + m_ItemHeight * m_ListCount) * 15
        PageCount = m_ListCount
    ElseIf m_ListCount > iMaxPageCount Then
        UserControl.Height = (m_relTop * 2 + m_ItemHeight * iMaxPageCount) * 15
        PageCount = iMaxPageCount
    End If
    
End Sub

Public Sub InitScrollValue()
    YVScroll1.Value = 0
    m_TopIndex = 0
    m_ScrollWidth = 0
End Sub

Sub SETCOLOR(BK As Long, CHOSE As Long)
COLOR_N = BK
COLOR_H = CHOSE
UserControl.Refresh
End Sub
