VERSION 5.00
Begin VB.UserControl ITXT 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox NumBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   120
   End
   Begin VB.PictureBox picSep 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   600
      MouseIcon       =   "ITXT.ctx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   1560
      Width           =   135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "ITXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function apiLockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hWndLock As Long) As Long
Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_USER = &H400
Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETSEL = &HB0
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_UNDO = &HC7
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum BDRstyle
    GTB_NoBorder = 0
    GTB_FixedSingle = 1
End Enum

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private MyTitle As String
Private AVMenabled As Boolean
Private TBchanged As Boolean
Private MyFileName As String
Private SStart As Long
Private OldLine As Long

Dim tbMouseX As Single
Dim tbMouseY As Single

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DBLCLICK() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event ChangeTitle(NewTitle As String)
Event BeforeChangeTitle(Cancel As Boolean)

Const Title_Height = 255

Dim PrintedTopNumber As Long

Public SyntaxColoring As Boolean
Private Sub ResizeControls()
Dim EdgeSize As Integer, NBWid As Long

If UserControl.Width < 1000 Then Exit Sub
If UserControl.Height < 1000 Then Exit Sub

EdgeSize = (Width - ScaleWidth) / 2
UserControl.Font = Text1.Font

If NumBar.Visible Then
    NumBar.Move 0, 0, NumBar.Width, ScaleHeight
    NBWid = NumBar.Width
Else
    NBWid = 0
End If
picSep.Move NBWid, Title_Height, 0, ScaleHeight - Title_Height
Text1.Move NBWid + picSep.Width, 0, ScaleWidth - (NBWid + picSep.Width), ScaleHeight
PrintLineCount
If NBWid <> 0 Then PrintNums
End Sub
Private Sub NumBar_Click()
Text1.SetFocus
End Sub

Private Sub NumBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.SetFocus
End Sub

Private Sub NumBar_Resize()
ResizeControls
End Sub

Private Sub picSep_Click()
    Text1.SetFocus
End Sub

Private Sub Text1_Change()
Dim S1 As Single, S2 As Single
TBchanged = True
If Not Text1.FontSize = Text1.FontSize Or Not Text1.FontName = Text1.FontName Then
    S1 = Text1.SelStart
    S2 = Text1.SelLength
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.FontSize = Text1.FontSize
    Text1.FontName = Text1.FontName
    Text1.SelStart = S1
    Text1.SelLength = S2
End If
If GetTopLineNumber <> PrintedTopNumber Then PrintNums
PrintLineCount
RaiseEvent Change
End Sub

Private Sub Text1_Click()
'PrintLineCount
RaiseEvent Click
'Text1.SelLength = 0
End Sub

Private Sub Text1_DblClick()
Text1.SelLength = 0
RaiseEvent DBLCLICK
End Sub

Private Sub Text1_GotFocus()
    PrintLineCount
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'PrintLineCount
If KeyCode = 9 And Shift = 0 And Text1.SelText = "" Then
    Text1.SelText = vbTab
    KeyCode = 0
End If
If KeyCode = 67 And Shift = 2 Then
    Exit Sub
End If
If GetLineNumber <> OldLine And OldLine <> 0 Then PrintNums
OldLine = GetLineNumber
RaiseEvent KeyDown(KeyCode, Shift)
'If KeyCode = 46 Or KeyCode = 8 Then
'    PrintNums
'End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Dim stg As String
'PrintLineCount
RaiseEvent KeyPress(KeyAscii)
'If KeyAscii = 13 Then
'    stg = GetLineOfText
'    If LCase(Left(stg, 11)) = "private sub" Then
'        If InStr(11, stg, "()") = 0 Then
'
'        End If
'    ElseIf LCase(Left(stg, 16)) = "private function" Then
'
'    End If
'End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim CommentColor As Long
'    Dim StringColor As Long
'    Dim KeysColor As Long
'
'If SyntaxColoring = True Then
'    LockWindowUpdate UserControl.hwnd
'
'    Select Case KeyCode
'
'    Case 13, 32 ', 38, 40, 37, 39
'        'CommentColor = RGB(0, 128, 0)       '// DARK GREEN  //
'        'StringColor = RGB(0, 0, 0)          '// BLACK       //
'        'KeysColor = RGB(0, 0, 128)          '// DARK BLUE   //
'        KeysColor = &H800000
'        StringColor = vbBlack
'        CommentColor = &H8000&
'
'        Colorize Text1, CommentColor, StringColor, KeysColor, KeyCode
'
'            Text1.SelColor = StringColor
'
'    End Select
'
'    LockWindowUpdate 0&
'End If
'###########
'###########
'###########
PrintLineCount
'If GetLineNumber <> OldLine And OldLine <> 0 Then PrintNums
'OldLine = GetLineNumber
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SStart = Text1.SelStart
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MOUSEMOVE(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Text1.SelStart <> SStart Then PrintLineCount: PrintNums
RaiseEvent MOUSEUP(Button, Shift, X, Y)
Text1.SetFocus
Screen.ActiveControl = Text1
If Button = 2 Then UserControl.PopupMenu Frmm.ÎÄ±¾
End Sub
Private Sub Timer1_Timer()
    If NumBar.Visible Then If GetTopLineNumber <> PrintedTopNumber Then PrintNums
End Sub
Private Sub UserControl_Initialize()
Text1.Text = ""
TBchanged = False
UserControl.Refresh
NumBar.Refresh
UserControl_Resize
oldproc = GetWindowLong(Text1.hwnd, GWL_WNDPROC)
SetWindowLong Text1.hwnd, GWL_WNDPROC, AddressOf TextWndProc
End Sub

Private Sub UserControl_Resize()
ResizeControls
PrintLineCount
PrintNums
End Sub
Private Sub PrintLineCount()
  Dim linecount As Long
  Dim LineNumber As Long
    linecount = apiSendMessage(Text1.hwnd, EM_GETLINECOUNT, 0&, 0&)
End Sub

Private Function GetLineNumber() As Long
Dim CaretPos As Long
Dim TXT As String, posa As Long, Posb As Long

    DoEvents
    If InStr(1, Text1.Text, vbCr) = 0 Then GetLineNumber = 1: Exit Function
    CaretPos = Text1.SelStart
    DoEvents
    TXT = Left(Text1.Text, CaretPos)
    If InStr(1, TXT, vbCr) = 0 Then GetLineNumber = 1: Exit Function
    
    posa = 1
    Posb = 1
    
    Do While posa <> 0 And posa <> Len(TXT)
        If posa = 1 Then
            posa = InStr(posa, TXT, vbCr)
            If posa = 1 Then posa = posa + 1
        Else
            posa = InStr(posa + 1, TXT, vbCr)
        End If
        If posa <> 0 Then
            Posb = Posb + 1
        End If
    Loop
    
    GetLineNumber = Posb
End Function

Private Sub PrintNumsold()
Dim TopNumber As Long
Dim Numbers As Integer, CNum As Long

NumBar.Font = Text1.Font

NumBar.FontSize = Text1.Font.Size

TopNumber = apiSendMessage(Text1.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
TopNumber = TopNumber + 1

Numbers = NumBar.ScaleHeight / (NumBar.TextHeight("ABC") + 0.545)

NumBar.Cls

NumBar.Width = 200 + NumBar.TextWidth(CStr(TopNumber + Numbers))

CNum = 0

For I = 1 To Numbers
    NumBar.CurrentX = 40
    NumBar.CurrentY = CNum
    NumBar.Print TopNumber
    TopNumber = TopNumber + 1
    CNum = CNum + NumBar.ScaleHeight / Numbers
Next I

End Sub

Private Sub PrintNums(Optional Override As Boolean)
Dim TopNumber As Long
Dim Numbers As Integer, CNum As Long
Dim oldsize As Currency
Dim txtHgt As Integer

    NumBar.Font = Text1.Font
    NumBar.FontSize = Text1.Font.Size
    
    TopNumber = GetTopLineNumber
    
    PrintedTopNumber = TopNumber
    
    If NumBar.Visible Then
        Numbers = Text1.Height / NumBar.TextHeight("ABC")
        NumBar.Line (0, 0)-(NumBar.ScaleWidth, NumBar.ScaleHeight), NumBar.BackColor, BF
        NumBar.Width = 210 + (NumBar.TextWidth(CStr(TopNumber + Numbers)) * 1.1)
        
        CNum = 0
        txtHgt = NumBar.TextHeight("ABC")
        
        For I = 1 To Numbers
            NumBar.CurrentX = 0
            NumBar.CurrentY = CNum
            NumBar.Print TopNumber
            TopNumber = TopNumber + 1
            CNum = CNum + txtHgt
        Next I
        NumBar.Refresh
    End If

End Sub

Private Function GetTopLineNumber() As Long
    Dim TopNumber As Long
    TopNumber = apiSendMessage(Text1.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    GetTopLineNumber = TopNumber + 1
End Function
Sub LoadFile(File As String)
Open File For Input As #1
Dim Str1 As String
While EOF(1) = False
Input #1, Str1
DoEvents
Text1.Text = Text1.Text & vbCrLf & Str1
Wend
Close #1
End Sub
Sub SaveFile(File As String, Optional UseTitleAsFilename As Boolean = True)
Dim fFile As String
fFile = App.Path & "\COFING\NODE.txt"
Open fFile For Binary As #1
Put #1, LOF(1) + 1, Now & vbCrLf & Text1.Text & vbCrLf & vbCrLf
Close #1
End Sub
Public Property Get SelStart() As Long
    SelStart = Text1.SelStart
End Property
Public Property Let SelStart(ByVal New_strt As Long)
    Text1.SelStart = New_strt
    'PropertyChanged "SelText"
End Property
Public Property Get SelLength() As Long
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_strt As Long)
    Text1.SelLength = New_strt
    'PropertyChanged "SelText"
End Property







Public Property Get Font() As Font
    Set Font = Text1.Font
End Property


Public Property Set Font(ByVal New_Font As Font)
    Dim S1 As Single, S2 As Single
    Set Text1.Font = New_Font
    Set NumBar.Font = New_Font
    
'    LockWindowUpdate Text1.hwnd
'    S1 = Text1.SelStart
'    S2 = Text1.SelLength
'    Text1.SelStart = 0
'    Text1.SelLength = Len(Text1.Text)
'    Text1.SelFontSize = Text1.Font.Size
'    Text1.SelFontName = Text1.Font.Name
'    Text1.SelBold = Text1.Font.Bold
'    Text1.SelItalic = Text1.Font.Italic
'    Text1.SelStrikeThru = Text1.Font.Strikethrough
'    Text1.SelUnderline = Text1.Font.Underline
'    Text1.SelStart = S1
'    Text1.SelLength = S2
'    LockWindowUpdate 0
    
    PrintNums
    PropertyChanged "Font"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Text() As String
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get MouseIcon() As PICTURE
    Set MouseIcon = Text1.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As PICTURE)
    Set Text1.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = Text1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ' Validation is supplied by UserControl.
    Text1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    picSep.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Public Property Get SelTextColor() As OLE_COLOR
'    SelTextColor = Text1.SelColor
'End Property

'Public Property Let SelTextColor(ByVal New_SelTextColor As OLE_COLOR)
'    Text1.SelColor() = New_SelTextColor
'    PropertyChanged "SelTextColor"
'End Property

Public Property Get Title() As String
    Title = MyTitle
End Property

Public Property Let Title(ByVal New_Title As String)
    MyTitle = New_Title
    RaiseEvent ChangeTitle(MyTitle)
    PrintLineCount
    PropertyChanged "Title"
End Property

Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get BorderStyle() As BDRstyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BDRstyle)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get AutoVerbMenu() As Boolean
    AutoVerbMenu = AVMenabled
End Property

Public Property Let AutoVerbMenu(ByVal New_Menu As Boolean)
    'Text1.AutoVerbMenu = New_Menu
    AVMenabled = New_Menu
    PropertyChanged "AutoVerbMenu"
End Property

Public Property Get NumBackColor() As OLE_COLOR
    NumBackColor = NumBar.BackColor
End Property

Public Property Let NumBackColor(ByVal New_NumBackColor As OLE_COLOR)
    NumBar.BackColor = New_NumBackColor
    PrintNums
    PropertyChanged "NumBackColor"
End Property

Public Property Get FOREColor() As OLE_COLOR
    FOREColor = Text1.FOREColor
End Property

Public Property Let FOREColor(ByVal New_FOREColor As OLE_COLOR)
    Text1.FOREColor = New_FOREColor
    PropertyChanged "FOREColor"
End Property

Public Property Get NumForeColor() As OLE_COLOR
    NumForeColor = NumBar.FOREColor
End Property

Public Property Let NumForeColor(ByVal New_NumForeColor As OLE_COLOR)
    NumBar.FOREColor = New_NumForeColor
    PrintNums
    PropertyChanged "NumForeColor"
End Property
Public Property Let TitleForeColor(ByVal New_TitleForeColor As OLE_COLOR)
    TTLBack.FOREColor = New_TitleForeColor
    PrintLineCount
    PropertyChanged "TitleForeColor"
End Property
Public Property Get filename() As String
    filename = MyFileName
End Property

Public Property Let filename(ByVal New_FileName As String)
    MyFileName = New_FileName
    PropertyChanged "FileName"
End Property

Public Property Get Numbar_Visible() As Boolean
    Numbar_Visible = NumBar.Visible
End Property

Public Property Let Numbar_Visible(ByVal NewVal As Boolean)
    NumBar.Visible = NewVal
    ResizeControls
    PropertyChanged "NumbarVisible"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picSep.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    'Text1.SelColor = PropBag.ReadProperty("SelTextColour", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    'm_Alignment = PropBag.ReadProperty("Alignment", 2)
    Text1.Text = PropBag.ReadProperty("Text", "Text")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Text1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    MyTitle = PropBag.ReadProperty("Title", "Untitled")
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    'Text1.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", True)
    AVMenabled = PropBag.ReadProperty("AutoVerbMenu", True)
    NumBar.BackColor = PropBag.ReadProperty("NumBackColor", &H808080)
    NumBar.FOREColor = PropBag.ReadProperty("NumForeColor", &HFFFFFF)
    MyFileName = PropBag.ReadProperty("FileName", "")
    'ColorText = PropBag.ReadProperty("ColorText", False)
    NumBar.Visible = PropBag.ReadProperty("NumbarVisible", True)
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        Timer1.Interval = 10
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Terminate()
SetWindowLong Text1.hwnd, GWL_WNDPROC, oldproc
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColor", picSep.BackColor, &H80000005)
    'Call PropBag.WriteProperty("SelTextColour", Text1.SelColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    'Call PropBag.WriteProperty("Alignment", m_Alignment, 2)
    Call PropBag.WriteProperty("Text", Text1.Text, "Text")
    Call PropBag.WriteProperty("MouseIcon", Text1.MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", Text1.MousePointer, 0)
    Call PropBag.WriteProperty("Title", MyTitle, "Untitled")
    Call PropBag.WriteProperty("Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("AutoverbMenu", AVMenabled, True)
    Call PropBag.WriteProperty("NumBackColor", NumBar.BackColor, &H808080)
    Call PropBag.WriteProperty("NumForeColor", NumBar.FOREColor, &HFFFFFF)
    Call PropBag.WriteProperty("FileName", MyFileName, "")
    'Call PropBag.WriteProperty("ColorText", ColorText, False)
    Call PropBag.WriteProperty("NumbarVisible", NumBar.Visible, True)
End Sub

Function GetLineOfText() As String
Dim SP As Long, EP As Long
SP = Text1.SelStart
On Local Error Resume Next
If SP = 0 Then Exit Function
'If SP = Len(Text1.Text) Then Exit Function
Do Until Mid(Text1.Text, SP - 1, 1) = vbCrLf
    'Debug.Print """" & Mid(Text1.Text, SP - 1, 1) & """"
    If Mid(Text1.Text, SP - 1, 1) = vbCr Then Exit Do
    If Mid(Text1.Text, SP - 1, 1) = vbLf Then Exit Do
    If Mid(Text1.Text, SP - 1, 1) = vbCrLf Then Exit Do
    SP = SP - 1
    If SP = 0 Or SP = 1 Then Exit Do
Loop
EP = SP
Do Until EP = Len(Text1.Text) 'Or Mid(Text1.Text, EP, 1) = vbCr
    EP = EP + 1
Loop
'EP = EP + 2
    If Mid(Text1.Text, SP, EP - (SP - 1)) = vbCr Then Exit Function
    If Mid(Text1.Text, SP, EP - (SP - 1)) = vbLf Then Exit Function
    If Mid(Text1.Text, SP, EP - (SP - 1)) = vbCrLf Then Exit Function

GetLineOfText = """" & Mid(Text1.Text, SP, EP - (SP - 1)) & """"
End Function
