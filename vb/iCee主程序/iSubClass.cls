VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Private Type ThisClassSet
    s_srcWndProcAddress     As Long
    s_Hwnd                  As Long
    
    n_heapAlloc             As Long
End Type
Dim LinkProc(29)  As Long
Dim PG                      As ThisClassSet

Event GetWindowMessage(Result As Long, ByVal chwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)

Private Sub MsgHook(Result As Long, ByVal chwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)
    '子类化接口过程
    RaiseEvent GetWindowMessage(Result, chwnd, Message, wParam, lParam)
End Sub

Private Function GetWndProcAddress(ByVal OrgWindowProc As Long, ByVal SinceCount As Long) As Long
'   地址指针 ＝ GetWndProcAddress( 取第 N 个公共函数（属性）  =或= 所有公共函数个数 + 第 N 个私有函数的函数地址）
    Dim mePtr As Long
    Dim jmpAddress As Long
    Dim i As Long
    Dim Protlng As Long
    
    mePtr = ObjPtr(Me)
    CopyMemory jmpAddress, ByVal mePtr, 4
    CopyMemory jmpAddress, ByVal jmpAddress + (SinceCount - 1) * 4 + &H1C, 4

    
    LinkProc(0) = &H83EC8B55
    LinkProc(1) = &H75FFFCC4
    LinkProc(2) = &H1075FF14
    LinkProc(3) = &HFF0C75FF
    LinkProc(4) = &HB80875
    LinkProc(5) = &HB000040
    LinkProc(6) = &HB94575C0
    LinkProc(7) = &H1000&
    LinkProc(8) = &H830C458B
    LinkProc(9) = &H87502F8
    LinkProc(10) = &H1C7&
    LinkProc(11) = &H1BEB0000
    LinkProc(12) = &H863D&
    LinkProc(13) = &H8B077500
    LinkProc(14) = &H1891045
    LinkProc(15) = &H5A3D0DEB
    LinkProc(16) = &H75000010
    LinkProc(17) = &H101C706
    LinkProc(18) = &H83000000
    LinkProc(19) = &H2750139
    LinkProc(20) = &H680EEB
    LinkProc(21) = &HB8000020
    LinkProc(22) = &H3000&
    LinkProc(23) = &H13EBD0FF
    LinkProc(24) = &H50FC458D
    LinkProc(25) = &H500068
    LinkProc(26) = &H6000B800
    LinkProc(27) = &HD0FF0000
    LinkProc(28) = &HC9FC458B
    LinkProc(29) = &H10C2&
        
    i = App.LogMode
    CopyMemory ByVal VarPtr(LinkProc(4)) + 3, i, 4&                                     ' Label Sign: 0400000
    CopyMemory ByVal VarPtr(LinkProc(25)) + 1, mePtr, 4&                                ' Label Sign: 0500000
    CopyMemory ByVal VarPtr(LinkProc(26)) + 2, jmpAddress, 4&                           ' Label Sign: 0600000
    
    If i Then
        i = VarPtr(LinkProc(0))
        Protlng = 120
    Else
        PG.n_heapAlloc = HeapAlloc(GetProcessHeap, &H8, 128&)
        CopyMemory ByVal PG.n_heapAlloc + 120&, 1&, 4
        LinkProc(7) = PG.n_heapAlloc + 120                                                  ' Label Sign: 0100000
        CopyMemory ByVal VarPtr(LinkProc(20)) + 3, OrgWindowProc, 4&                        ' Label Sign: 0200000
        LinkProc(22) = GetProcAddress(GetModuleHandle("user32.dll"), "CallWindowProcA")     ' Label Sign: 0300000
        
        CopyMemory ByVal PG.n_heapAlloc&, LinkProc(0), 120&
        i = PG.n_heapAlloc
        Protlng = 128
    End If
    
    VirtualProtect ByVal i&, Protlng, &H40, mePtr
    GetWndProcAddress = i
End Function

Function CallDefaultWindowProc(ByVal chwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '调用窗口默认处理过程
    CallDefaultWindowProc = CallWindowProc(PG.s_srcWndProcAddress, ByVal chwnd&, ByVal Message&, ByVal wParam&, ByVal lParam&)
End Function

Function SetMsgHook(ByVal chwnd As Long) As Long
    '设置指定窗口的子类化
    PG.s_Hwnd = chwnd
    PG.s_srcWndProcAddress = GetWindowLong(chwnd, ByVal -4&)
    SetWindowLong ByVal chwnd, ByVal -4&, ByVal GetWndProcAddress(PG.s_srcWndProcAddress, 4)
    SetMsgHook = PG.s_srcWndProcAddress
End Function

Sub SetMsgUnHook()
    '取消窗口子类化
    SetWindowLong ByVal PG.s_Hwnd&, ByVal -4&, ByVal PG.s_srcWndProcAddress&
End Sub

