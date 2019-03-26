VERSION 5.00
Begin VB.UserControl IVScroll 
   BackColor       =   &H001F1F1F&
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   FillColor       =   &H001F1F1F&
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   Begin VB.Timer tmrIsLbuttonDown 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   600
   End
End
Attribute VB_Name = "IVScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const nH = 18   '滚动条向上向下按钮的高度
Private Const nMinGlideHeight = 10 '滑块的最小高度

Private Const GWL_WNDPROC           As Long = -4

Private Const TME_LEAVE                         As Long = &H2

Private Const WM_PAINT = &HF
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_MOVE = &H3
Private Const WM_MOUSELEAVE = &H2A3
Private Const WM_ERASEBKGND = &H14
Private Const WM_KILLFOCUS = &H8

Private Const STRETCH_HALFTONE = 4

Private Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type

Private Type TRACKMOUSEEVENTTYPE
    CBSIZE                                      As Long
    dwFlags                                     As Long
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
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

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, PT As POINTAPI) As Long
Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private m_MinV As Long   '最小值
Private m_MaxV As Long   '最大值
Private m_Value As Long   '当前值
Private m_SmallChange As Long
Private m_LargeChange As Long
Private m_GlideHeight As Long    '滑块高度

'按下鼠标时各种状态值
Private m_Index As Long     '1      按下向上按钮
                            '2      按下向下按钮
                            '3      按下滑块
                            '0      在不包括滑块的滚动区域按下

                         
'鼠标移动时各种状态值
Private m_HoverIndex As Long     '1      在向上按钮上
                                 '2      在向下按钮上
                                 '3      在滑块上
                                 '-1     不在滚动条上

            
Private m_OldY As Long

Private m_IsLbuttonDown As Boolean  '是否按下鼠标左键
Private m_IsFirstDown As Boolean    '该变量的作用是防止计时器事件第一次就触发
Private m_IsGlideDown As Boolean    '是否在滑块上按下鼠标左键
                                   
Private m_Count As Long
Private m_Mod As Long
Private m_IsMoreThan As Boolean


Private m_OldValue As Long
Private m_GlideTop As Long
Private m_srcVScroll As Long

Private m_rel As Long

'事件声明:
Public Event Change()

                         
'这个过程必须放在用户控件的最顶部
'要隐藏这个过程可以在"工具"菜单 -> "过程属性" -> "名称"选择(zSubclass_Proc) -> 点击"高级"按钮，选中"隐藏该成员."
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhwnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    
    Dim uPS As PAINTSTRUCT
    Dim dwY As Long
    Dim p As POINTAPI
    Dim TmpDC As Long
    Dim rc As RECT
    
    Select Case uMsg
    
        Case WM_PAINT

            BeginPaint lhwnd, uPS
            TmpDC = pCreateDC(ScaleWidth, ScaleHeight - 2 * nH)
            SetRect rc, 0, 0, ScaleWidth, ScaleHeight - 2 * nH
            pFillRect TmpDC, rc, vbWhite
            SetStretchBltMode TmpDC, STRETCH_HALFTONE
            SetStretchBltMode uPS.hdc, STRETCH_HALFTONE
            
            Call pDrawGlideAndBkg(TmpDC)
            Call BitBlt(uPS.hdc, 0, nH, ScaleWidth, ScaleHeight - 2 * nH, TmpDC, 0, 0, vbSrcCopy)
            
            Call pDrawTopButton(uPS.hdc)
            Call pDrawBottomButton(uPS.hdc)
             
            EndPaint lhwnd, uPS
            
            DeleteDC TmpDC
            bHandled = True
            lReturn = 0
         
        Case WM_KILLFOCUS
            m_IsLbuttonDown = False
            m_IsGlideDown = False
            tmrIsLbuttonDown.Enabled = False
         
        Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
            m_IsLbuttonDown = True
            m_IsFirstDown = True
            tmrIsLbuttonDown.Enabled = True
            dwY = HIWORD(lParam)
            m_OldY = dwY
            m_Index = DownTest(dwY)
            Call OnLbuttonDown(dwY)
            
        
        Case WM_LBUTTONUP
            m_IsLbuttonDown = False
            tmrIsLbuttonDown.Enabled = False
            m_IsGlideDown = False
            InvalidateRect lhwnd, 0, 0
            
        Case WM_MOUSEMOVE
            TrackMouseTracking lhwnd
            dwY = HIWORD(lParam)
            Dim tmpIndex As Long
            tmpIndex = MoveTest(dwY)
            If tmpIndex <> m_HoverIndex Then
                m_HoverIndex = tmpIndex
                InvalidateRect lhwnd, 0, 0
                If m_HoverIndex <> -1 Then InvalidateRect lhwnd, 0, 0
            End If
            
        Case WM_MOUSELEAVE
            m_HoverIndex = -1
            InvalidateRect lhwnd, 0, 0

    End Select

End Sub

Private Sub pDrawGlideAndBkg(ByVal mDC As Long)

    'Dim tmpGlideTop As Long

    '贴背景图
    pStretchBlt mDC, 0, 0, ScaleWidth, ScaleHeight - nH * 2, m_srcVScroll, 32 * 6, 0, 32, 32, 1
            
    '贴滑块
    If m_GlideTop < nH + 3 Then m_GlideTop = nH + 3
    If m_GlideTop > ScaleHeight - nH - 3 - GlideHeight Then
        m_GlideTop = ScaleHeight - nH - 3 - GlideHeight
    End If
    
    'If m_GlideTop > ScaleHeight - nH - GlideHeight - 6 Then
    '    tmpGlideTop = ScaleHeight - nH - GlideHeight - 6
    'Else
    '    tmpGlideTop = m_GlideTop
    'End If
                        
    If m_IsGlideDown Then
        '贴滑块按下的图
        pStretchBlt mDC, 0, m_GlideTop - nH, ScaleWidth, GlideHeight, m_srcVScroll, 32 * 4, 0, 32, 32, 1
    Else
        If m_HoverIndex = 3 Then
            '贴鼠标在上方的图
            pStretchBlt mDC, 0, m_GlideTop - nH, ScaleWidth, GlideHeight, m_srcVScroll, 32 * 4, 0, 32, 32, 1
            Exit Sub
        End If
        '贴正常时的图
        pStretchBlt mDC, 0, m_GlideTop - nH, ScaleWidth, GlideHeight, m_srcVScroll, 32 * 4, 0, 32, 32, 1
    End If
End Sub

Private Sub tmrIsLbuttonDown_Timer()

    Dim PT As POINTAPI
    GetCursorPos PT
    ScreenToClient hwnd, PT
    
    '拖动滑块时所作的处理
    If m_IsGlideDown Then
        If (PT.Y - m_OldY) Then
            Dim tmpTop As Long
            tmpTop = m_GlideTop
            m_GlideTop = m_GlideTop + (PT.Y - m_OldY)
            If m_GlideTop <= ScaleHeight - GlideHeight - nH And m_GlideTop >= nH Then
                m_OldY = PT.Y
            Else
                If m_GlideTop > ScaleHeight - GlideHeight - nH Then
                    m_GlideTop = ScaleHeight - GlideHeight - nH
                Else
                    m_GlideTop = nH
                End If
                m_OldY = m_rel + m_GlideTop
            End If
            Dim nValue As Long
            nValue = pGetValueByGlideTop(m_GlideTop)
            If m_OldValue <> nValue Then
                m_Value = nValue
                RaiseEvent Change
            End If
            If tmpTop <> m_GlideTop Then
                InvalidateRect hwnd, 0, 0
            End If
        End If
        Exit Sub
    End If
    
    '对鼠标移出滚动条外面所作的处理
    If PT.X < 0 Or PT.X > ScaleWidth Or PT.Y < 0 Or PT.Y > ScaleHeight Then
        m_HoverIndex = -1
        m_IsLbuttonDown = False
        tmrIsLbuttonDown.Enabled = False
        InvalidateRect hwnd, 0, 0
        Exit Sub
    Else
        m_HoverIndex = MoveTest(PT.Y)
    End If

    If m_IsFirstDown Then m_IsFirstDown = False: Exit Sub
    
    Dim tmpIndex As Long
    Dim IsStop As Boolean
    tmpIndex = DownTest(PT.Y)
    If tmpIndex <> m_Index Then
        m_IsLbuttonDown = False
        tmrIsLbuttonDown.Enabled = False
        InvalidateRect hwnd, 0, 0
        Exit Sub
    End If
    
    Select Case m_Index
        Case 1
            m_Value = m_Value - m_SmallChange
        Case 2
            m_Value = m_Value + m_SmallChange
        Case 0
            If PT.Y <= pGetGlideTop(m_Value) Then
                m_Value = m_Value - m_LargeChange
                If PT.Y > pGetGlideTop(m_Value) Then IsStop = True
            ElseIf PT.Y >= pGetGlideTop(m_Value) + GlideHeight Then
                m_Value = m_Value + m_LargeChange
                If PT.Y < pGetGlideTop(m_Value) + GlideHeight Then IsStop = True
            End If
    End Select

    If m_Value <= m_MinV Then m_Value = m_MinV
    If m_Value >= m_MaxV Then m_Value = m_MaxV
    
    If m_OldValue <> m_Value Then
        m_GlideTop = pGetGlideTop(m_Value)
        InvalidateRect hwnd, 0, 0
        RaiseEvent Change
    End If
    
    If IsStop Then
        m_IsLbuttonDown = False
        tmrIsLbuttonDown.Enabled = False
    End If
    
End Sub


Private Function DownTest(ByVal dwY As Long) As Long
    If dwY >= 0 And dwY <= nH Then
        DownTest = 1
    ElseIf dwY >= ScaleHeight - nH And dwY <= ScaleHeight Then
        DownTest = 2
    ElseIf dwY > m_GlideTop And dwY < m_GlideTop + GlideHeight Then
        DownTest = 3
    Else
        DownTest = 0
    End If
End Function

Private Sub OnLbuttonDown(ByVal dwY As Long)

    Select Case m_Index
        Case 1
            m_Value = m_Value - m_SmallChange
        Case 2
            m_Value = m_Value + m_SmallChange
        Case 3
            m_IsGlideDown = True
            m_rel = dwY - pGetGlideTop(Value)
            Exit Sub
        Case 0
            If dwY <= pGetGlideTop(Value) Then
                m_Value = m_Value - m_LargeChange
            ElseIf dwY >= pGetGlideTop(Value) + GlideHeight Then
                m_Value = m_Value + m_LargeChange
            End If
    End Select
    
    If m_Value < m_MinV Then m_Value = m_MinV
    If m_Value > m_MaxV Then m_Value = m_MaxV
    
    If m_OldValue <> m_Value Then
        RaiseEvent Change
    End If
    
    m_GlideTop = pGetGlideTop(Value)
    InvalidateRect hwnd, 0, 0
    
End Sub

Private Function MoveTest(ByVal dwY As Long) As Long
    MoveTest = -1
    If dwY >= 0 And dwY <= nH Then
        MoveTest = 1
    ElseIf dwY >= ScaleHeight - nH And dwY <= ScaleHeight Then
        MoveTest = 2
    ElseIf dwY > m_GlideTop And dwY < m_GlideTop + GlideHeight Then
        MoveTest = 3
    End If
End Function

Private Sub TrackMouseTracking(ByVal hwnd As Long)
    Dim lpEventTrack As TRACKMOUSEEVENTTYPE
    With lpEventTrack
        .CBSIZE = Len(lpEventTrack)
        .dwFlags = TME_LEAVE
        .hwndTrack = hwnd
    End With
    TrackMouseEvent lpEventTrack
End Sub

Private Function HIWORD(ByVal lParam As Long) As Long
    HIWORD = lParam \ 65536
End Function
    
Private Function LOWORD(ByVal lParam As Long) As Long
    LOWORD = lParam Mod 65536
End Function

Private Sub pDrawScroll(ByVal mDC As Long)
    Call pDrawTopButton(mDC)
    Call pDrawBottomButton(mDC)
    Call pDrawGlideArea(mDC)
    Call pDrawGlide(mDC)
End Sub

'画向上按钮
Private Sub pDrawTopButton(ByVal hdc As Long)
    If m_IsLbuttonDown Then
        '贴按下时的图
        If m_Index = 1 Then
            pStretchBlt hdc, 0, 0, ScaleWidth, nH, m_srcVScroll, 32, 0, 32, 32, 2
        End If
    Else
        If m_HoverIndex = 1 Then
            '贴鼠标在上方的图
            pStretchBlt hdc, 0, 0, ScaleWidth, nH, m_srcVScroll, 32, 0, 32, 32, 2
        Else
            '贴正常时的图
            pStretchBlt hdc, 0, 0, ScaleWidth, nH, m_srcVScroll, 0, 0, 32, 32, 2
        End If
    End If
End Sub

'画向下按钮
Private Sub pDrawBottomButton(ByVal hdc As Long)
    If m_IsLbuttonDown Then
        '贴按下时的图
        If m_Index = 2 Then
            pStretchBlt hdc, 0, ScaleHeight - nH, ScaleWidth, nH, m_srcVScroll, 32 * 3, 0, 32, 32, 2
        End If
    Else
        If m_HoverIndex = 2 Then
            '贴鼠标在上方的图
            pStretchBlt hdc, 0, ScaleHeight - nH, ScaleWidth, nH, m_srcVScroll, 32 * 3, 0, 32, 32, 2
        Else
            '贴正常时的图
            pStretchBlt hdc, 0, ScaleHeight - nH, ScaleWidth, nH, m_srcVScroll, 32 * 2, 0, 32, 32, 2
        End If
    End If

End Sub

Private Function pStretchBlt(ByVal DesDC As Long, ByVal DesX As Long, ByVal DesY As Long, ByVal DesWidth As Long, _
    ByVal DesHeight As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, _
    ByVal SrcHeight As Long, ByVal n As Long) As Long
    
    StretchBlt DesDC, DesX, DesY, DesWidth, n, SrcDC, SrcX, SrcY, SrcWidth, n, vbSrcCopy  '上边
    
    StretchBlt DesDC, DesX, DesY + DesHeight - n, DesWidth, n, SrcDC, SrcX, SrcY + SrcHeight - n, SrcWidth, n, _
                                                                                                  vbSrcCopy '下边
                                                                                          
    StretchBlt DesDC, DesX, DesY + n, n, DesHeight - 2 * n, SrcDC, SrcX, SrcY + n, n, SrcHeight - 2 * n, _
                                                                                                  vbSrcCopy '左边
                                                                                                  
    StretchBlt DesDC, DesX + DesWidth - n, DesY + n, n, DesHeight - 2 * n, SrcDC, SrcX + SrcWidth - n, _
                                                                  SrcY + n, n, SrcHeight - 2 * n, vbSrcCopy '右边
    
    StretchBlt DesDC, DesX + n, DesY + n, DesWidth - 2 * n, DesHeight - 2 * n, SrcDC, SrcX + n, _
                                       SrcY + n, SrcWidth - 2 * n, SrcHeight - 2 * n, vbSrcCopy '右边
    
End Function

'画滑动区域
Private Sub pDrawGlideArea(ByVal hdc As Long)
    '贴背景图
    pStretchBlt hdc, 0, nH + 1, ScaleWidth, ScaleHeight - nH * 2 - 2, m_srcVScroll, 32 * 6, 0, 32, 32, 1
End Sub

'画滑块
Private Sub pDrawGlide(ByVal hdc As Long)
    
    If m_GlideTop < nH + 3 Then m_GlideTop = nH + 3
    If m_GlideTop > ScaleHeight - nH - 3 - GlideHeight Then
        m_GlideTop = ScaleHeight - nH - 3 - GlideHeight
    End If
                     
    If m_IsGlideDown Then
        '贴滑块按下的图
        pStretchBlt hdc, 2, m_GlideTop, ScaleWidth - 4, GlideHeight, m_srcVScroll, 32 * 5, 0, 32, 32, 1
    Else
        If m_HoverIndex = 3 Then
            '贴鼠标在上方的图
            pStretchBlt hdc, 2, m_GlideTop, ScaleWidth - 4, GlideHeight, m_srcVScroll, 32 * 5, 0, 32, 32, 1
            Exit Sub
        End If
        '贴正常时的图
        pStretchBlt hdc, 2, m_GlideTop, ScaleWidth - 4, GlideHeight, m_srcVScroll, 32 * 4, 0, 32, 32, 1
    End If
    
End Sub

'得到滑块的Top值(比较抽象,用数学归纳法找出规律)
Public Function pGetGlideTop(ByVal CurValue As Long) As Long
    
    If CurValue = MinV Then
        pGetGlideTop = nH + 3
        Exit Function
    ElseIf CurValue = MaxV Then
        pGetGlideTop = ScaleHeight - GlideHeight - nH - 3
        Exit Function
    End If
    
    Dim TMP As Long
    If m_IsMoreThan Then
        If m_Mod Then
            TMP = MinV + m_Mod
            If CurValue > MinV And CurValue <= TMP Then
                pGetGlideTop = (CurValue - MinV) * (m_Count + 1) + (nH + 3)
            Else
                pGetGlideTop = pGetGlideTop(TMP) + (CurValue - TMP) * m_Count
            End If
        Else
            pGetGlideTop = (CurValue - MinV) * m_Count + (nH + 3)
        End If
    Else
        If m_Mod Then
            TMP = MinV + m_Mod * (m_Count + 1)
            If CurValue > MinV And CurValue <= TMP Then
                pGetGlideTop = (CurValue - MinV) \ (m_Count + 1) + (nH + 3)
            Else
                pGetGlideTop = pGetGlideTop(TMP) + (CurValue - TMP) \ m_Count
            End If
        Else
            pGetGlideTop = (CurValue - MinV) \ m_Count + (nH + 3)
        End If
    End If
    
End Function

'根据滑块的Top值得到对应的Value值(近似)
Private Function pGetValueByGlideTop(ByVal nGlideTop As Long)
    
    Dim TMP As Long
    
    If nGlideTop <= nH + 3 Then
        pGetValueByGlideTop = MinV
        Exit Function
    ElseIf nGlideTop >= ScaleHeight - GlideHeight - nH - 3 Then
        pGetValueByGlideTop = MaxV
        Exit Function
    End If

    If m_IsMoreThan Then
        If m_Mod Then
            TMP = m_Mod * (m_Count + 1) + nH + 3
            If nGlideTop > nH + 3 And nGlideTop <= TMP Then
                pGetValueByGlideTop = MinV + (nGlideTop - nH - 3) \ (m_Count + 1)
            Else
                pGetValueByGlideTop = MinV + m_Mod + (nGlideTop - TMP) \ m_Count
            End If
        Else
            pGetValueByGlideTop = (nGlideTop - nH - 3) \ m_Count + MinV
        End If
    Else
        If m_Mod Then
            TMP = m_Mod * (m_Count + 1)
            If nGlideTop > nH + 3 And nGlideTop <= nH + 3 + m_Mod Then
                pGetValueByGlideTop = MinV + (m_Count + 1) * (nGlideTop - nH - 3)
            Else
                pGetValueByGlideTop = MinV + TMP + m_Count * (nGlideTop - nH - 3 - m_Mod)
            End If
        Else
            pGetValueByGlideTop = (nGlideTop - nH - 3) * m_Count + MinV
        End If
    End If

End Function

'得到m_GlideHeight,m_Mod和m_Count的值,还有m_GlideTop的值
Private Sub pGetGlideInformation()

    Dim MaxCnt As Long
    Dim ActMaxCnt As Long
    MaxCnt = ScaleHeight - nH * 2 - nMinGlideHeight - 6
    
    If m_GlideHeight < nMinGlideHeight Then m_GlideHeight = nMinGlideHeight
    If m_GlideHeight >= MaxCnt Then
        'MsgBox "滑块高度值超过最大值,请适当调高滚动条的高度!", vbOKOnly, "提示"
        Exit Sub
    End If
    
    ActMaxCnt = ScaleHeight - nH * 2 - m_GlideHeight - 6
    
    If (MaxV - MinV) <= ActMaxCnt Then
        m_Mod = ActMaxCnt Mod (MaxV - MinV)
        m_Count = ActMaxCnt \ (MaxV - MinV)
        m_IsMoreThan = True
    Else
        m_Mod = (MaxV - MinV) Mod ActMaxCnt
        m_Count = (MaxV - MinV) \ ActMaxCnt
        m_IsMoreThan = False
    End If
    
    m_GlideTop = pGetGlideTop(m_Value)
    
End Sub

Private Sub UserControl_Initialize()
    m_MinV = 0
    m_MaxV = 20
    m_SmallChange = 1
    m_LargeChange = 10
    m_HoverIndex = -1
    m_GlideTop = nH
    m_srcVScroll = pCreateDCByHandle(LoadResPicture(102, vbResBitmap).handle)
    m_GlideHeight = nMinGlideHeight
End Sub

Private Sub UserControl_Paint()
    SetStretchBltMode hdc, STRETCH_HALFTONE
    Call pDrawScroll(hdc)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  MinV = PropBag.ReadProperty("MinV", m_MinV)
  MaxV = PropBag.ReadProperty("MaxV", m_MaxV)
  Value = PropBag.ReadProperty("Value", m_Value)
  SmallChange = PropBag.ReadProperty("SmallChange", m_SmallChange)
  LargeChange = PropBag.ReadProperty("LargeChange", m_LargeChange)

  If Ambient.UserMode Then
      Call Subclass_Start(hwnd)
      Call Subclass_AddMsg(hwnd, WM_PAINT, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_LBUTTONDOWN, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_RBUTTONDOWN, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_LBUTTONUP, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_RBUTTONUP, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_ERASEBKGND, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_MOUSEMOVE, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_MOUSELEAVE, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_LBUTTONDBLCLK, MSG_BEFORE)
      Call Subclass_AddMsg(hwnd, WM_KILLFOCUS, MSG_BEFORE)
      
      
      '你也可以添加其他消息.
      '当消息发生时 你可以在 zSubclass_Proc 这个Function里获取
  End If
End Sub

Private Sub UserControl_Resize()
    Call pGetGlideInformation
End Sub

Private Sub UserControl_Terminate()
  On Error GoTo Catch
  Call Subclass_StopAll   '停止所有消息
  Call DeleteDC(m_srcVScroll)
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

Public Property Get MinV() As Long
    MinV = m_MinV
End Property

Public Property Let MinV(ByVal va As Long)
    If va <= Value And va < MaxV Then
        m_MinV = va
        Call pGetGlideInformation
        InvalidateRect hwnd, 0, 0
    End If
End Property

Public Property Get MaxV() As Long
    MaxV = m_MaxV
End Property

Public Property Let MaxV(ByVal va As Long)
    If Value <> MinV Then
        If va >= Value Then
            m_MaxV = va
            Call pGetGlideInformation
            InvalidateRect hwnd, 0, 0
        End If
    Else
        If va > Value Then
            m_MaxV = va
            Call pGetGlideInformation
            InvalidateRect hwnd, 0, 0
        End If
    End If
End Property

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal va As Long)
    If (va >= MinV And va <= MaxV) Then
        m_Value = va
        m_OldValue = va
        m_GlideTop = pGetGlideTop(va)
        InvalidateRect hwnd, 0, 0
    End If
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal va As Long)
    If va > 0 Then
        m_SmallChange = va
    End If
End Property

Public Property Get LargeChange() As Long
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal va As Long)
    If va > 0 Then
        m_LargeChange = va
    End If
End Property

Public Property Get GlideHeight() As Long
    GlideHeight = m_GlideHeight
End Property

Public Property Let GlideHeight(ByVal va As Long)
    m_GlideHeight = IIf(va <= nMinGlideHeight, nMinGlideHeight, va)
    Call pGetGlideInformation
End Property

Public Property Get hwnd()
    hwnd = UserControl.hwnd
End Property

Public Property Get OldValue() As Long
    OldValue = m_OldValue
End Property

Public Property Let OldValue(ByVal va As Long)
    m_OldValue = va
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MinV", m_MinV)
    Call PropBag.WriteProperty("MaxV", m_MaxV)
    Call PropBag.WriteProperty("Value", m_Value)
    Call PropBag.WriteProperty("SmallChange", m_SmallChange)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange)
End Sub

Private Sub pFillRect(ByVal hdc As Long, rc As RECT, ByVal clrFill As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(clrFill)
    FillRect hdc, rc, hBrush
    DeleteObject hBrush
End Sub



