Attribute VB_Name = "SystemHook"
' by CLE


' APIs
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)



Private Const WH_MOUSE_LL = 14
Private Const WH_KEYBOARD_LL = 13
Private Const WH_MOUSE = 7
Private Const WH_KEYBOARD = 2
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_MBUTTONUP = &H208
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105

Private Const VK_SHIFT As Byte = &H10
Private Const VK_CAPITAL As Byte = &H14
Private Const VK_NUMLOCK As Byte = &H90



Public Type Point
    x As Long
    y As Long
End Type

Private Type KeyboardHookStruct
    vkCode As Long
    ScanCode As Long
    Flags As Long
    Time As Long
    DwExtraInfo As Long
End Type

Dim hKeyboardHook As Long

Sub RegHook()
    If hKeyboardHook = 0 Then
        hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardHookProcHookProc, App.hInstance, 0)
        If hKeyboardHook = 0 Then
            ' 这里处理注册错误
            MsgBox "注册系统键盘钩子失败！"
        Else
            
        End If
    End If
End Sub

Sub UnHook()
    If hKeyboardHook <> 0 Then
        Dim retKeyboard As Long
        retKeyboard = UnhookWindowsHookEx(hKeyboardHook)
        hKeyboardHook = 0
        If retKeyboard = 0 Then
            ' 这里处理卸载错误
            MsgBox "卸载系统键盘钩子失败！"
        End If
    End If
End Sub


' 回调函数
Private Function KeyboardHookProcHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If nCode >= 0 Then
        Dim ks As KeyboardHookStruct
        CopyMemory ks, ByVal lParam, 20
        
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
            ' 这里处理键盘按下的事件
            'Form2.key_down ks.vkCode
        End If
        
        If wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            ' 这里处理键盘弹起事件
            'Form2.key_up ks.vkCode
        End If
        
        
        ' 想要截获按键事件，可以直接设置 KeyboardHookProcHookProc = 1
        ' 否则呼叫下一个钩子
        CallNextHookEx hKeyboardHook, nCode, wParam, lParam
    End If
End Function





















