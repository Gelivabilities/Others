Attribute VB_Name = "Module1"
Option Explicit
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2

Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Download by http://www.codefans.net

'*************************************************************************
'**�� �� ����transparence
'**��    �룺ByVal Frm(Form)     -  Ҫ͸���Ĵ�������
'**        ��ByVal alpha(Single) -  ���ô����͸����
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�mrlbb
'*************************************************************************
Public Sub transparence(ByVal Frm As Form, ByVal alpha As Single)
    '**********************************��������Ϊ͸��
    Dim rtn As Long
    rtn = GetWindowLong(Frm.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong Frm.hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes Frm.hwnd, 0, 255 * alpha, LWA_ALPHA 'LWA_COLORKEY
    '***********************************
End Sub


