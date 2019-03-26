Attribute VB_Name = "Module1"
Private Declare Function CallNextHookEx Lib "user32" _
                          (ByVal hHook As Long, _
                          ByVal nCode As Long, _
                          ByVal wParam As Long, _
                          lParam As Any) As Long

Private Declare Function SetWindowsHookEx Lib "user32" _
                          Alias "SetWindowsHookExA" _
                          (ByVal idHook As Long, _
                          ByVal lpfn As Long, _
                          ByVal hmod As Long, _
                          ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" _
                          (ByVal hHook As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
                          Alias "RtlMoveMemory" _
                          (Destination As Any, _
                          Source As Any, _
                          ByVal Length As Long)

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Const WM_LBUTTONUP = &H202
Private Const WH_MOUSE_LL = 14

Private hHook As Long

Public Function MouseHook(ByVal nCode As Long, _
                       ByVal wParam As Long, _
                       ByVal lParam As Long) As Long

    Dim mhs As MSLLHOOKSTRUCT, pt As POINTAPI

    If wParam = WM_LBUTTONUP Then
        Call CopyMemory(mhs, ByVal lParam, LenB(mhs))
        pt = mhs.pt
        Call CopyMemory(p, ByVal lParam, Len(p))
        'Debug.Print "×ó¼üµ¥»÷    ×ø±ê:" & pt.x & "  "; pt.y
        Form2.addCoins
        'If Form2.Visible = False Then
            'Call Form2.Command1_Click
        'End If
    End If

    Call CallNextHookEx(hHook, nCode, wParam, lParam)
End Function

Public Sub HooK()
    hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHook, App.hInstance, 0)
End Sub

Public Sub UnHooK()
    Call UnhookWindowsHookEx(hHook)
End Sub

