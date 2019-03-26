Attribute VB_Name = "Module1"

Option Explicit
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Private MyObj As Object
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const GWL_WNDPROC = -4&
Public OldWindowProc As Long '��������ϵͳĬ�ϵĴ�����Ϣ�������ĵ�ַ
'�Զ������Ϣ������
Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    '    Debug.Print Msg
    If Msg = WM_NCLBUTTONDOWN Then
        NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
        'FormMove�¼�����ʼ��
        MyObj.Form1.Left = "X��" & Form2.Left
        MyObj.Form1.Left = "Y��" & Form2.Top
        'FormMove�¼���������
    Else
        NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
    End If
End Function
Public Sub myhook(ByVal obj As Object)
    Set MyObj = obj
    Dim hwndobject As Long
    hwndobject = obj.hwnd
    OldWindowProc = GetWindowLong(hwndobject, GWL_WNDPROC)
    Call SetWindowLong(hwndobject, GWL_WNDPROC, AddressOf NewWindowProc)
End Sub

