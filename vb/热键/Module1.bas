Attribute VB_Name = "Module1"
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fskey_Modifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal Scan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const WM_HOTKEY = &H312
Const MOD_ALT = &H1
Const MOD_CONTROL = &H2
Const MOD_SHIFT = &H4
Const GWL_WNDPROC = (-4) '���ں����ĵ�ַ

Dim key_preWinProc As Long '�������洰����Ϣ
Dim key_Modifiers As Long, key_uVirtKey As Long, key_idHotKey As Long
Dim key_IsWinAddress As Boolean '�Ƿ�ȡ�ô�����Ϣ���ж�

Function keyWndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If Msg = WM_HOTKEY Then
Select Case wParam 'wParam ֵ���� key_idHotKey
Case 1 '���� 3 ���ȼ���,3 ���ȼ�����Ӧ�Ĳ���,����������ĳ����У�ֻҪ�޸Ĵ˴��Ϳ�����
    Form1.Top = Form1.Top + 100
Case 2
    Form1.Top = Form1.Top - 100
Case 3
    Form1.Left = Form1.Left + 100
Case 4
    Form1.Left = Form1.Left - 100
Case 5
    Form1.Width = Form1.Width + 100
End Select
End If

' ����Ϣ���͸�ָ���Ĵ���
keyWndproc = CallWindowProc(key_preWinProc, hwnd, Msg, wParam, lParam)

End Function

Function SetHotkey(ByVal KeyId As Long, ByVal KeyAss0 As String, ByVal Action As String)
Dim KeyAss1 As Long
Dim KeyAss2 As String
Dim i As Long

i = InStr(1, KeyAss0, ",")
If i = 0 Then
KeyAss1 = Val(KeyAss0)
KeyAss2 = ""
Else
KeyAss1 = Right(KeyAss0, Len(KeyAss0) - i)
KeyAss2 = Left(KeyAss0, i - 1)
End If

key_idHotKey = 0
key_Modifiers = 0
key_uVirtKey = 0

If key_IsWinAddress = False Then '�ж��Ƿ���Ҫȡ�ô�����Ϣ������ظ�ȡ��,�����ָ�����ʱ��������ɳ�������
' ��¼ԭ����window�����ַ
key_preWinProc = GetWindowLong(Form1.hwnd, GWL_WNDPROC)
' ���Զ���������ԭ����window����
SetWindowLong Form1.hwnd, GWL_WNDPROC, AddressOf keyWndproc
End If

key_idHotKey = KeyId
Select Case Action
Case "Add"
If KeyAss2 = "Ctrl" Then key_Modifiers = MOD_CONTROL
If KeyAss2 = "Alt" Then key_Modifiers = MOD_ALT
If KeyAss2 = "Shift" Then key_Modifiers = MOD_SHIFT
If KeyAss2 = "Ctrl+Alt" Then key_Modifiers = MOD_CONTROL + MOD_ALT
If KeyAss2 = "Ctrl+Shift" Then key_Modifiers = MOD_CONTROL + MOD_SHIFT
If KeyAss2 = "Ctrl+Alt+Shift" Then key_Modifiers = MOD_CONTROL + MOD_ALT + MOD_SHIFT
If KeyAss2 = "Shift+Alt" Then key_Modifiers = MOD_SHIFT + MOD_ALT
key_uVirtKey = Val(KeyAss1)
RegisterHotKey Form1.hwnd, key_idHotKey, key_Modifiers, key_uVirtKey '�򴰿�ע��ϵͳ�ȼ�
key_IsWinAddress = True '����Ҫ��ȡ�ô�����Ϣ

Case "Del"
SetWindowLong Form1.hwnd, GWL_WNDPROC, key_preWinProc '�ָ�������Ϣ
UnregisterHotKey Form1.hwnd, key_uVirtKey 'ȡ��ϵͳ�ȼ�
key_IsWinAddress = False '�����ٴ�ȡ�ô�����Ϣ
End Select
End Function
