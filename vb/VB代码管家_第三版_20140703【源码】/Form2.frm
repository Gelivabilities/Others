VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SPY++"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8070
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox PIDֵ 
      Height          =   270
      Left            =   3840
      TabIndex        =   23
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox ������ 
      Height          =   270
      Left            =   960
      TabIndex        =   22
      Top             =   855
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5760
      Top             =   1800
   End
   Begin VB.TextBox ��ɫֵ 
      Height          =   270
      Left            =   6720
      TabIndex        =   10
      Top             =   2655
      Width           =   1215
   End
   Begin VB.TextBox ��ǰ���� 
      Height          =   270
      Left            =   960
      TabIndex        =   9
      Top             =   2655
      Width           =   4815
   End
   Begin VB.TextBox ������� 
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Top             =   2295
      Width           =   4815
   End
   Begin VB.TextBox ����·�� 
      Height          =   270
      Left            =   960
      TabIndex        =   7
      Top             =   1935
      Width           =   4815
   End
   Begin VB.TextBox �������� 
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   1575
      Width           =   4815
   End
   Begin VB.TextBox �Ӿ�� 
      Height          =   270
      Left            =   960
      TabIndex        =   5
      Top             =   495
      Width           =   1935
   End
   Begin VB.TextBox ������� 
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Top             =   1215
      Width           =   4815
   End
   Begin VB.TextBox ����� 
      Height          =   270
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox ��ȡ���� 
      Height          =   270
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox ��ǰ���� 
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   5880
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   840
         Picture         =   "Form2.frx":030A
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Label Label13 
      Caption         =   "P I D ֵ:"
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "�� �� ��:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "R G B ֵ:"
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   2685
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "�϶�ͼ�굽Ŀ��λ������:"
      Height          =   240
      Left            =   5880
      TabIndex        =   20
      Top             =   135
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "��ǰ����:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2685
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "�������:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2325
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "����·��:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1965
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "��������:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1605
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "�������:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1245
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "�� �� ��:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "�� �� ��:"
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "��ȡ����:"
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   165
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "��ǰ����:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   165
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'---------------------�����״
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'---------------------ȡ�����µľ��
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'---------------------�����Ӿ��ȡ����������
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpclassname As String, ByVal nMaxCount As Long) As Long
'---------------------ȡ��������
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'---------------------ȡĳ�������
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'---------------------���ݾ��ȡ����·��
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'---------------------�����ļ�
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
'---------------------ȡ�������

'-------------------------------------------------------------------------------------------------------------------
Dim a     As POINTAPI
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long '��ȡָ�����ڵ��豸����
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long '�ͷ��豸����
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long 'ȡRGBֵ
Private Sub GetRGB(ByVal Col As Long, ByRef r As Long, ByRef g As Long, ByRef B As Long) 'ת��ΪRGB
    r = Col Mod 256
    g = ((Col And &HFF00&) \ 256&) Mod 256&
    B = (Col And &HFF0000) \ 65536
End Sub
'-------------------------------------------------------------------------------------------------------------------ȡ����µ���ɫֵ


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Visible = False
    SetCursor Image1.Picture
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ��ȡ����.Text = ""
    �Ӿ��.Text = ""
    �����.Text = ""
    �������.Text = ""
    ��������.Text = ""
    ����·��.Text = ""
    �������.Text = ""
    ��ǰ����.Text = ""
    ��ɫֵ.Text = ""
    ������.Text = ""
    PIDֵ.Text = ""
    
    Dim ������� As String
    
    Dim a As POINTAPI
    GetCursorPos a
    
    ��ȡ���� = a.x & "," & a.y
    
    �Ӿ�� = WindowFromPoint(a.x, a.y)
    ����� = WindowFromPoint(a.x, a.y)
    
    '***************************************ȡ��ǰ����
    ��ǰ���� = ""
    Dim dq As String
    dq = Space(255) '�˾��൱�ڿո�
    GetWindowText �Ӿ��, dq, 255 'hwndΪ������
    ��ǰ���� = dq
    '***************************************
    
    ������� = ""
    ������� = �����
    Dim leiming As String
    leiming = Space(255)
    GetClassName �����, leiming, 255
    �������� = ""
    �������� = leiming
    
1:
    ������� = GetParent(�����) '�����Ϊ�Ӿ��,�������Ϊ��һ��ؼ��ľ��
    If ������� <> 0 Then '���������������
        ������� = ������� & "\" & �������
        GetClassName �������, leiming, 255
        �������� = �������� & "\" & leiming
        ����� = �������
        ������� = ""
        GoTo 1: '����ת����ǩ1����
    End If
    
    '***************************************ȡ�������
    ������� = ""
    Dim bt As String
    bt = Space(255) '�˾��൱�ڿո�
    GetWindowText �����, bt, 255 'hwndΪ������
    ������� = bt
    '***************************************
    
    '***************************************���ݾ��ȡ����·��
    Dim PID As Long
    GetWindowThreadProcessId �����, PID '���ݾ��ȡ�ý���ID,text1Ϊ���
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process")
    For Each objprocess In colProcesses
        
        If objprocess.processid = PID Then  '��ĳ����ID����ID��ͬʱ
            ����·�� = objprocess.ExecutablePath  'ȡ�ý���·��
            ������ = objprocess.Name
            PIDֵ = PID
        End If
        
    Next
    '***************************************
    '***************************************ȡ����µ���ɫֵ
    Dim hdc     As Long, Col       As Long
    Dim r     As Long, g       As Long, B       As Long
    hdc = GetDC(0)
    GetCursorPos a
    cor = GetPixel(hdc, a.x, a.y) 'a.xΪ�����꣬a.yΪ������
    GetRGB cor, r, g, B
    ReleaseDC Me.hwnd, hdc
    ��ɫֵ = r & "," & g & "," & B
    Frame1.BackColor = RGB(r, g, B) '�ı�Picture1����ɫֵ����Ĳ�������ֻ����Frame1
    Image1.Visible = True
End Sub

Private Sub Timer1_Timer()
    Dim a As POINTAPI
    GetCursorPos a
    ��ǰ���� = a.x & "," & a.y
End Sub
