VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'MD5加密
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private i As Integer
Private j As Integer
Private K As Integer
Private a As Byte
Private b As Byte
Dim m As Integer
Private L As Long
Private RC4KEY(255) As Byte
Private ADDTABLE(255, 255) As Byte
Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21


Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte

Property Get RegisterA() As String
    RegisterA = State(1)
End Property

Property Get RegisterB() As String
    RegisterB = State(2)
End Property

Property Get RegisterC() As String
    RegisterC = State(3)
End Property

Property Get RegisterD() As String
    RegisterD = State(4)
End Property

Public Function DigestFileToHexStr(filename As String) As String
    Open filename For Binary Access Read As #1
    MD5Init
    Do While Not EOF(1)
        Get #1, , ByteBuffer
        If Loc(1) < LOF(1) Then
            ByteCounter = ByteCounter + 64
            MD5Transform ByteBuffer
        End If
    Loop
    ByteCounter = ByteCounter + (LOF(1) Mod 64)
    Close #1
    MD5Final
    DigestFileToHexStr = GetValues
End Function


Public Function DigestStrToHexStr(SourceString As String) As String
    MD5Init
    MD5Update Len(SourceString), StringToArray(SourceString)
    MD5Final
    DigestStrToHexStr = GetValues
End Function


Private Function StringToArray(InString As String) As Byte()
    Dim i As Integer
    Dim bytBuffer() As Byte
    ReDim bytBuffer(Len(InString))
    For i = 0 To Len(InString) - 1
        bytBuffer(i) = Asc(Mid(InString, i + 1, 1))
    Next
    StringToArray = bytBuffer
End Function

Public Function GetValues() As String
    GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))
End Function


Private Function LongToString(num As Long) As String
        Dim a As Byte
        Dim b As Byte
        Dim C As Byte
        Dim d As Byte
        
        a = num And &HFF&
        If a < 16 Then
            LongToString = "0" & Hex(a)
        Else
            LongToString = Hex(a)
        End If
               
        b = (num And &HFF00&) \ 256
        If b < 16 Then
            LongToString = LongToString & "0" & Hex(b)
        Else
            LongToString = LongToString & Hex(b)
        End If
        
        C = (num And &HFF0000) \ 65536
        If C < 16 Then
            LongToString = LongToString & "0" & Hex(C)
        Else
            LongToString = LongToString & Hex(C)
        End If
       
        If num < 0 Then
            d = ((num And &H7F000000) \ 16777216) Or &H80&
        Else
            d = (num And &HFF000000) \ 16777216
        End If
        
        If d < 16 Then
            LongToString = LongToString & "0" & Hex(d)
        Else
            LongToString = LongToString & Hex(d)
        End If
    
End Function

Public Sub MD5Init()
    ByteCounter = 0
    State(1) = UnsignedToLong(1732584193#)
    State(2) = UnsignedToLong(4023233417#)
    State(3) = UnsignedToLong(2562383102#)
    State(4) = UnsignedToLong(271733878#)
End Sub

'
' MD5 Final
'
Public Sub MD5Final()
    Dim dblBits As Double
    
    Dim padding(72) As Byte
    Dim lngBytesBuffered As Long
    
    padding(0) = &H80
    
    dblBits = ByteCounter * 8
    
    ' Pad out
    lngBytesBuffered = ByteCounter Mod 64
    If lngBytesBuffered <= 56 Then
        MD5Update 56 - lngBytesBuffered, padding
    Else
        MD5Update 120 - ByteCounter, padding
    End If
    
    
    padding(0) = UnsignedToLong(dblBits) And &HFF&
    padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
    padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
    padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
    padding(4) = 0
    padding(5) = 0
    padding(6) = 0
    padding(7) = 0
    
    MD5Update 8, padding
End Sub

'
' Break up input stream into 64 byte chunks
'
Public Sub MD5Update(InputLen As Long, InputBuffer() As Byte)
    Dim ii As Integer
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim lngBufferedBytes As Long
    Dim lngBufferRemaining As Long
    Dim lngRem As Long
    
    lngBufferedBytes = ByteCounter Mod 64
    lngBufferRemaining = 64 - lngBufferedBytes
    ByteCounter = ByteCounter + InputLen
    ' Use up old buffer results first
    If InputLen >= lngBufferRemaining Then
        For ii = 0 To lngBufferRemaining - 1
            ByteBuffer(lngBufferedBytes + ii) = InputBuffer(ii)
        Next
        MD5Transform ByteBuffer
        
        lngRem = (InputLen) Mod 64
        ' The transfer is a multiple of 64 lets do some transformations
        For i = lngBufferRemaining To InputLen - ii - lngRem Step 64
            For j = 0 To 63
                ByteBuffer(j) = InputBuffer(i + j)
            Next
            MD5Transform ByteBuffer
        Next
        lngBufferedBytes = 0
    Else
      i = 0
    End If
    
    ' Buffer any remaining input
    For K = 0 To InputLen - i - 1
        ByteBuffer(lngBufferedBytes + K) = InputBuffer(i + K)
    Next
    
End Sub

'
' MD5 Transform
'
Private Sub MD5Transform(Buffer() As Byte)
    Dim X(16) As Long
    Dim a As Long
    Dim b As Long
    Dim C As Long
    Dim d As Long
    
    a = State(1)
    b = State(2)
    C = State(3)
    d = State(4)
    
    Decode 64, X, Buffer

    ' Round 1
    FF a, b, C, d, X(0), S11, -680876936
    FF d, a, b, C, X(1), S12, -389564586
    FF C, d, a, b, X(2), S13, 606105819
    FF b, C, d, a, X(3), S14, -1044525330
    FF a, b, C, d, X(4), S11, -176418897
    FF d, a, b, C, X(5), S12, 1200080426
    FF C, d, a, b, X(6), S13, -1473231341
    FF b, C, d, a, X(7), S14, -45705983
    FF a, b, C, d, X(8), S11, 1770035416
    FF d, a, b, C, X(9), S12, -1958414417
    FF C, d, a, b, X(10), S13, -42063
    FF b, C, d, a, X(11), S14, -1990404162
    FF a, b, C, d, X(12), S11, 1804603682
    FF d, a, b, C, X(13), S12, -40341101
    FF C, d, a, b, X(14), S13, -1502002290
    FF b, C, d, a, X(15), S14, 1236535329
    
    ' Round 2
    gg a, b, C, d, X(1), S21, -165796510
    gg d, a, b, C, X(6), S22, -1069501632
    gg C, d, a, b, X(11), S23, 643717713
    gg b, C, d, a, X(0), S24, -373897302
    gg a, b, C, d, X(5), S21, -701558691
    gg d, a, b, C, X(10), S22, 38016083
    gg C, d, a, b, X(15), S23, -660478335
    gg b, C, d, a, X(4), S24, -405537848
    gg a, b, C, d, X(9), S21, 568446438
    gg d, a, b, C, X(14), S22, -1019803690
    gg C, d, a, b, X(3), S23, -187363961
    gg b, C, d, a, X(8), S24, 1163531501
    gg a, b, C, d, X(13), S21, -1444681467
    gg d, a, b, C, X(2), S22, -51403784
    gg C, d, a, b, X(7), S23, 1735328473
    gg b, C, d, a, X(12), S24, -1926607734
    
    ' Round 3
    hh a, b, C, d, X(5), S31, -378558
    hh d, a, b, C, X(8), S32, -2022574463
    hh C, d, a, b, X(11), S33, 1839030562
    hh b, C, d, a, X(14), S34, -35309556
    hh a, b, C, d, X(1), S31, -1530992060
    hh d, a, b, C, X(4), S32, 1272893353
    hh C, d, a, b, X(7), S33, -155497632
    hh b, C, d, a, X(10), S34, -1094730640
    hh a, b, C, d, X(13), S31, 681279174
    hh d, a, b, C, X(0), S32, -358537222
    hh C, d, a, b, X(3), S33, -722521979
    hh b, C, d, a, X(6), S34, 76029189
    hh a, b, C, d, X(9), S31, -640364487
    hh d, a, b, C, X(12), S32, -421815835
    hh C, d, a, b, X(15), S33, 530742520
    hh b, C, d, a, X(2), S34, -995338651
    
    ' Round 4
    ii a, b, C, d, X(0), S41, -198630844
    ii d, a, b, C, X(7), S42, 1126891415
    ii C, d, a, b, X(14), S43, -1416354905
    ii b, C, d, a, X(5), S44, -57434055
    ii a, b, C, d, X(12), S41, 1700485571
    ii d, a, b, C, X(3), S42, -1894986606
    ii C, d, a, b, X(10), S43, -1051523
    ii b, C, d, a, X(1), S44, -2054922799
    ii a, b, C, d, X(8), S41, 1873313359
    ii d, a, b, C, X(15), S42, -30611744
    ii C, d, a, b, X(6), S43, -1560198380
    ii b, C, d, a, X(13), S44, 1309151649
    ii a, b, C, d, X(4), S41, -145523070
    ii d, a, b, C, X(11), S42, -1120210379
    ii C, d, a, b, X(2), S43, 718787259
    ii b, C, d, a, X(9), S44, -343485551
    
    
    State(1) = LongOverflowAdd(State(1), a)
    State(2) = LongOverflowAdd(State(2), b)
    State(3) = LongOverflowAdd(State(3), C)
    State(4) = LongOverflowAdd(State(4), d)

'  /* Zeroize sensitive information.
'*/
'  MD5_memset ((POINTER)x, 0, sizeof (x));
    
End Sub

Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double
    
    intDblIndex = 0
    For intByteIndex = 0 To Length - 1 Step 4
        dblSum = InputBuffer(intByteIndex) + _
                                    InputBuffer(intByteIndex + 1) * 256# + _
                                    InputBuffer(intByteIndex + 2) * 65536# + _
                                    InputBuffer(intByteIndex + 3) * 16777216#
        OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
        intDblIndex = intDblIndex + 1
    Next
End Sub

'
' FF, GG, HH, and II transformations for rounds 1, 2, 3, and 4.
' Rotation is separate from addition to prevent recomputation.
'
Private Function FF(a As Long, _
                    b As Long, _
                    C As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, (b And C) Or (Not (b) And d), X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function gg(a As Long, _
                    b As Long, _
                    C As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, (b And d) Or (C And Not (d)), X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function hh(a As Long, _
                    b As Long, _
                    C As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, b Xor C Xor d, X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function ii(a As Long, _
                    b As Long, _
                    C As Long, _
                    d As Long, _
                    X As Long, _
                    s As Long, _
                    ac As Long) As Long
    a = LongOverflowAdd4(a, C Xor (b Or Not (d)), X, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Function LongLeftRotate(Value As Long, Bits As Long) As Long
    Dim lngSign As Long
    Dim lngI As Long
    Bits = Bits Mod 32
    If Bits = 0 Then LongLeftRotate = Value: Exit Function
    For lngI = 1 To Bits
        lngSign = Value And &HC0000000
        Value = (Value And &H3FFFFFFF) * 2
        Value = Value Or ((lngSign < 0) And 1) Or (CBool(lngSign And _
                &H40000000) And &H80000000)
    Next
    LongLeftRotate = Value
End Function

Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + _
                   ((Val2 And &HFFFF0000) \ 65536) + _
                   ((val3 And &HFFFF0000) \ 65536) + _
                   ((val4 And &HFFFF0000) \ 65536) + _
                   lngOverflow) And &HFFFF&
    LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

Private Function UnsignedToLong(Value As Double) As Long
        If Value < 0 Or Value >= OFFSET_4 Then Error 6
        If Value <= MAXINT_4 Then
          UnsignedToLong = Value
        Else
          UnsignedToLong = Value - OFFSET_4
        End If
      End Function

Private Function LongToUnsigned(Value As Long) As Double
        If Value < 0 Then
          LongToUnsigned = Value + OFFSET_4
        Else
          LongToUnsigned = Value
        End If
End Function


'Base64编码函数:Base64Encode
'Instr1    编码前字符串
'Outstr1    编码后字符串
Public Function Base64Encode(InStr1 As String) As String
    Dim mInByte(3) As Byte, mOutByte(4) As Byte
    Dim myByte As Byte
    Dim i As Integer, LenArray As Integer, j As Integer
    Dim myBArray() As Byte
    Dim OutStr1 As String
    
    myBArray() = StrConv(InStr1, vbFromUnicode)
    LenArray = UBound(myBArray) + 1
    For i = 0 To LenArray Step 3
        If LenArray - i = 0 Then
            Exit For
        End If
        If LenArray - i = 2 Then
            mInByte(0) = myBArray(i)
            mInByte(1) = myBArray(i + 1)
            Base64EncodeByte mInByte, mOutByte, 2
        ElseIf LenArray - i = 1 Then
            mInByte(0) = myBArray(i)
            Base64EncodeByte mInByte, mOutByte, 1
        Else
            mInByte(0) = myBArray(i)
            mInByte(1) = myBArray(i + 1)
            mInByte(2) = myBArray(i + 2)
            Base64EncodeByte mInByte, mOutByte, 3
        End If
        For j = 0 To 3
            OutStr1 = OutStr1 & Chr(mOutByte(j))
        Next j
    Next i
    Base64Encode = OutStr1
End Function

Private Sub Base64EncodeByte(mInByte() As Byte, mOutByte() As Byte, num As Integer)
   Dim tByte As Byte
   Dim i As Integer

   If num = 1 Then
       mInByte(1) = 0
       mInByte(2) = 0
   ElseIf num = 2 Then
       mInByte(2) = 0
   End If
   tByte = mInByte(0) And &HFC
   mOutByte(0) = tByte / 4
   tByte = ((mInByte(0) And &H3) * 16) + (mInByte(1) And &HF0) / 16
   mOutByte(1) = tByte
   tByte = ((mInByte(1) And &HF) * 4) + ((mInByte(2) And &HC0) / 64)
   mOutByte(2) = tByte
   tByte = (mInByte(2) And &H3F)
   mOutByte(3) = tByte
   For i = 0 To 3
       If mOutByte(i) >= 0 And mOutByte(i) <= 25 Then
           mOutByte(i) = mOutByte(i) + Asc("A")
       ElseIf mOutByte(i) >= 26 And mOutByte(i) <= 51 Then
           mOutByte(i) = mOutByte(i) - 26 + Asc("a")
       ElseIf mOutByte(i) >= 52 And mOutByte(i) <= 61 Then
           mOutByte(i) = mOutByte(i) - 52 + Asc("0")
       ElseIf mOutByte(i) = 62 Then
           mOutByte(i) = Asc("+")
       Else
           mOutByte(i) = Asc("/")
       End If
   Next i
   If num = 1 Then
       mOutByte(2) = Asc("=")
       mOutByte(3) = Asc("=")
   ElseIf num = 2 Then
       mOutByte(3) = Asc("=")
   End If
End Sub

Public Function Base64Decode(InStr1 As String) As String
   Dim mInByte(4) As Byte, mOutByte(3) As Byte
   Dim i As Integer, LenArray As Integer, j As Integer
   Dim myBArray() As Byte
   Dim OutStr1 As String
   Dim tmpArray() As Byte

   myBArray() = StrConv(InStr1, vbFromUnicode)
   LenArray = UBound(myBArray)
   ReDim tmpArray(((LenArray + 1) / 4) * 3)
    j = 0
    
    For i = 0 To LenArray Step 4
       If LenArray - i = 0 Then
           Exit For
       Else
            mInByte(0) = myBArray(i)
            mInByte(1) = myBArray(i + 1)
            mInByte(2) = myBArray(i + 2)
            mInByte(3) = myBArray(i + 3)
            Base64DecodeByte mInByte, mOutByte, 4
        End If
        tmpArray(j * 3) = mOutByte(0)
        tmpArray(j * 3 + 1) = mOutByte(1)
        tmpArray(j * 3 + 2) = mOutByte(2)
        j = j + 1
    Next i
    Base64Decode = BinaryToString(tmpArray)
End Function

Private Sub Base64DecodeByte(mInByte() As Byte, mOutByte() As Byte, ByteNum As Integer)
    Dim tByte As Byte
    Dim i As Integer
    ByteNum = 0
    For i = 0 To 3
        If mInByte(i) >= Asc("A") And mInByte(i) <= Asc("Z") Then
            mInByte(i) = mInByte(i) - Asc("A")
        ElseIf mInByte(i) >= Asc("a") And mInByte(i) <= Asc("z") Then
            mInByte(i) = mInByte(i) - Asc("a") + 26
        ElseIf mInByte(i) >= Asc("0") And mInByte(i) <= Asc("9") Then
            mInByte(i) = mInByte(i) - Asc("0") + 52
        ElseIf mInByte(i) = Asc("+") Then
            mInByte(i) = 62
        ElseIf mInByte(i) = Asc("/") Then
            mInByte(i) = 63
        Else '"="
            ByteNum = ByteNum + 1
            mInByte(i) = 0
        End If
    Next i
    '取前六位
    tByte = (mInByte(0) And &H3F) * 4 + (mInByte(1) And &H30) / 16
    '0的六位和1的前两位
    mOutByte(0) = tByte
    tByte = (mInByte(1) And &HF) * 16 + (mInByte(2) And &H3C) / 4
    '1的后四位和2的前四位
    mOutByte(1) = tByte
    tByte = (mInByte(2) And &H3) * 64 + (mInByte(3) And &H3F)
    mOutByte(2) = tByte
    '2的后两位和3的六位
ERR:
End Sub
Private Function BinaryToString(ByVal BinaryStr As Variant) As String '二进制转换为字符串
  Dim lnglen As Long
  Dim tmpBin As Variant
  Dim strC As String
  Dim skipflag As Long
  Dim i As Long
  skipflag = 0
  strC = ""
  
  If Not IsNull(BinaryStr) Then
      lnglen = LenB(BinaryStr)
      For i = 1 To lnglen
          If skipflag = 0 Then
            tmpBin = MidB(BinaryStr, i, 1)
            If AscB(tmpBin) > 127 Then
                strC = strC & Chr(AscW(MidB(BinaryStr, i + 1, 1) & tmpBin))
                skipflag = 1
            Else
                strC = strC & Chr(AscB(tmpBin))
            End If
          Else
            skipflag = 0
          End If
      Next
    End If
    BinaryToString = strC
End Function

Private Function StringToBinary(ByVal VarString As String) As Variant '字符串转成二进制
  Dim strBin As Variant
  Dim varchar As Variant
  Dim varasc As Long
  Dim varlow, varhigh
  Dim i As Long
  strBin = ""
  
  For i = 1 To Len(VarString)
      varchar = Mid(VarString, i, 1)
      varasc = Asc(varchar)
      If varasc < 0 Then
          varasc = varasc + 65535
      End If
      If varasc > 255 Then
          varlow = Left(Hex(Asc(varchar)), 2)
          varhigh = Right(Hex(Asc(varchar)), 2)
          
          
          strBin = strBin & ChrB("&H" & varlow) & ChrB("&H" & varhigh)
      Else
          strBin = strBin & ChrB(AscB(varchar))
      End If
  Next
  StringToBinary = strBin
End Function
