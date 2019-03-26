VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ЖўЮЌТы
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private m_objField As New clsFiniteField
Private m_objRS As New clsReedSolomon
Private m_nAlignmentPatternStep(1 To 40) As Byte
Private m_nECBlockCount(1 To 4, 1 To 40) As Byte
Private m_nECCodewordPerBlock(1 To 4, 1 To 40) As Byte
Private Type typeDynamicProgrammingNode
 nPrevMode As Long
 nCost As Long
 nCount As Long 'only when mode=1,2
End Type
Private m_nPowerOfTwo(17) As Long

Private Sub Class_Initialize()
Dim i As Long, j As Long ', k As Long
Dim t(19) As Currency
'///init field
m_objField.Init 256, 285, 2
'///init alignment pattern
For i = 1 To 6
 m_nAlignmentPatternStep(i) = 255
Next i
For i = 7 To 13
 m_nAlignmentPatternStep(i) = i * 2 + 2
Next i
For i = 14 To 20
 m_nAlignmentPatternStep(i) = ((i * 2 + 4) \ 3) * 2
Next i
For i = 21 To 27
 m_nAlignmentPatternStep(i) = (i + 2) And &HFFFFFFFE
Next i
For i = 28 To 34
 m_nAlignmentPatternStep(i) = (i \ 3) * 2 + 6
Next i
For i = 35 To 40
 m_nAlignmentPatternStep(i) = (i \ 3) * 2 + 2
Next i
'///
t(0) = 7234017283807.6673@
t(1) = 28879552945704.1665@
t(2) = 28936068276237.1585@
t(3) = 43403882080855.757@
t(4) = 57871806698986.0098@
t(5) = 79545709764228.0196@
t(6) = 115743501728796.0836@
t(7) = 115771760039978.855@
t(8) = 151829034411806.183@
t(9) = 180708697744900.4551@
t(10) = 245645085252721.4856@
t(11) = 231346486727206.9641@
t(12) = 267572646950765.0828@
t(13) = 303629921758806.6572@
t(14) = 347005546390815.4382@
t(15) = 390381280975671.272@
t(16) = 433785163058198.197@
t(17) = 477160787262394.5235@
t(18) = 534976298105354.1141@
t(19) = 585585940022854.8376@
CopyMemory m_nECBlockCount(1, 1), t(0), 160&
'///
t(0) = 202382271797140.1223@
t(1) = 116025973148025.4991@
t(2) = 202438570236877.8266@
t(3) = 187971418724617.8836@
t(4) = 202439669751874.1022@
t(5) = 202495527519598.1332@
t(6) = 173503827394324.6362@
t(7) = 216851411038831.2086@
t(8) = 202552264903282.5884@
t(9) = 202608559036262.6588@
t(10) = 173785741330141.0332@
t(11) = 217020298611904.4126@
t(12) = 216964002757568.8218@
t(13) = 217020298611904.4126@
t(14) = 217020298611904.4126@
t(15) = 217020298611904.4126@
t(16) = 217020298611904.4126@
t(17) = 217020298611904.4126@
t(18) = 217020298611904.4126@
t(19) = 217020298611904.4126@
CopyMemory m_nECCodewordPerBlock(1, 1), t(0), 160&
'///
j = 1
For i = 0 To 17
 m_nPowerOfTwo(i) = j
 j = j + j
Next i
End Sub

'internal function
Friend Function pDataModuleCount(ByVal nVersion As Long) As Long
Dim m As Long, i As Long
m = pModuleSize(nVersion)
Select Case nVersion
Case -4 To -1
 pDataModuleCount = m * m - 81 + 4 * nVersion
Case 1 To 40
 m = m * (m - 2) - 191
 i = pAlignmentPatternSize(nVersion)
 If i >= 2 Then
  m = m + i * (10 - i * 25) + 55
  If i > 2 Then m = m - 36
 End If
 pDataModuleCount = m
End Select
End Function

'internal function
Friend Function pDataCodewordCount(ByVal nVersion As Long) As Long
Select Case nVersion
Case -4 To -1
 pDataCodewordCount = (pDataModuleCount(nVersion) + 4) \ 8
Case 1 To 40
 pDataCodewordCount = pDataModuleCount(nVersion) \ 8
End Select
End Function

'internal function
Friend Function pModuleSize(ByVal nVersion As Long) As Long
Select Case nVersion
Case -4 To -1
 pModuleSize = 9 - 2 * nVersion
Case 1 To 40
 pModuleSize = 17 + 4 * nVersion
End Select
End Function

'internal function
Friend Function pAlignmentPatternSize(ByVal nVersion As Long) As Long
If nVersion < 2 Then pAlignmentPatternSize = 1 _
Else pAlignmentPatternSize = 2 + nVersion \ 7
End Function

'internal function
Friend Function pAlignmentPatternCount(ByVal nVersion As Long) As Long
Dim i As Long
i = pAlignmentPatternSize(nVersion)
If i < 2 Then pAlignmentPatternCount = 0 _
Else pAlignmentPatternCount = i * i - 3
End Function

'bInput: 0-based byte data
'nSize: input data size
'nVersion: -1 to -4 (currently unsupported), 1 to 40, or 0 for automatic
'          another negative number for automatic (MicroQR allowed) (currently unsupported)
'nECLevel: 1 to 4
'nMaskType: 0 to 7, -1 for automatic (currently unsupported)
Friend Function Encode(ByRef bInput() As Byte, ByVal nSize As Long, Optional ByVal nVersion As Long, Optional ByVal nECLevel As Long = 2, Optional ByVal nMaskType As Long = -1) As StdPicture
Dim b() As Byte, b2() As Byte
Dim bEncodedBit() As Byte
Dim bInterleavedBit() As Byte
Dim nPolynomial(255) As Long
'///
Dim nEncodedBitCount As Long
Dim nDataCodewordCount As Long
Dim nAvaliableDataCodewordCount As Long
Dim nECBlockCount As Long
Dim nSmallBlockCount As Long
Dim nDataCodewordPerBlock As Long
Dim nECCodewordPerBlock As Long
Dim nModuleSize As Long
'///
Dim i As Long, ii As Long, j As Long, K As Long, kk As Long
Dim LP As Long, lp2 As Long
Dim bb As Byte
'///
If nECLevel <= 0 Or nECLevel > 4 Then nECLevel = 2
'///
If nVersion < 0 Then
 'TODO:
Else
 ReDim bEncodedBit(32767)
 '///find optimal size
 If nVersion = 0 Or nVersion > 40 Then
  Do
   nEncodedBitCount = pEncodeToBitArray(bEncodedBit, bInput, nSize, 1, True)
   For nVersion = 1 To 9
    nAvaliableDataCodewordCount = pDataCodewordCount(nVersion) - CLng(m_nECBlockCount(nECLevel, nVersion)) _
    * CLng(m_nECCodewordPerBlock(nECLevel, nVersion))
    If nEncodedBitCount - 4 <= nAvaliableDataCodewordCount * 8& Then Exit Do
   Next nVersion
   nEncodedBitCount = pEncodeToBitArray(bEncodedBit, bInput, nSize, 10, True)
   For nVersion = 10 To 26
    nAvaliableDataCodewordCount = pDataCodewordCount(nVersion) - CLng(m_nECBlockCount(nECLevel, nVersion)) _
    * CLng(m_nECCodewordPerBlock(nECLevel, nVersion))
    If nEncodedBitCount - 4 <= nAvaliableDataCodewordCount * 8& Then Exit Do
   Next nVersion
   nEncodedBitCount = pEncodeToBitArray(bEncodedBit, bInput, nSize, 27, True)
   For nVersion = 27 To 40
    nAvaliableDataCodewordCount = pDataCodewordCount(nVersion) - CLng(m_nECBlockCount(nECLevel, nVersion)) _
    * CLng(m_nECCodewordPerBlock(nECLevel, nVersion))
    If nEncodedBitCount - 4 <= nAvaliableDataCodewordCount * 8& Then Exit Do
   Next nVersion
   'too many data
   Exit Function
  Loop
 End If
 '///check size
 nEncodedBitCount = pEncodeToBitArray(bEncodedBit, bInput, nSize, nVersion)
 nECBlockCount = m_nECBlockCount(nECLevel, nVersion)
 nECCodewordPerBlock = m_nECCodewordPerBlock(nECLevel, nVersion)
 nDataCodewordCount = pDataCodewordCount(nVersion)
 nAvaliableDataCodewordCount = nDataCodewordCount - nECBlockCount * nECCodewordPerBlock
 If nEncodedBitCount - 4 > nAvaliableDataCodewordCount * 8& Then Exit Function
 nDataCodewordPerBlock = (nAvaliableDataCodewordCount + nECBlockCount - 1) \ nECBlockCount
 nSmallBlockCount = nAvaliableDataCodewordCount Mod nECBlockCount
 If nSmallBlockCount > 0 Then nSmallBlockCount = nECBlockCount - nSmallBlockCount
 '///padding
 For i = 0 To nAvaliableDataCodewordCount * 8& - nEncodedBitCount
  j = i And &HF&
  bEncodedBit(nEncodedBitCount + i) = (j <= 2 Or j = 4 Or j = 5 Or j = 11 Or j = 15) And 1&
 Next i
 '///calc error correction code and interleave
 ReDim bInterleavedBit(32767)
 m_objRS.Init nDataCodewordPerBlock, nECCodewordPerBlock, m_objField
 LP = 0
 For i = 0 To nECBlockCount - 1
  nPolynomial(nDataCodewordPerBlock + nECCodewordPerBlock - 1) = 0
  lp2 = i
  For j = 0 To nDataCodewordPerBlock - 1
   If i < nSmallBlockCount And j = nDataCodewordPerBlock - 1 Then Exit For
   Debug.Assert LP < nAvaliableDataCodewordCount
   Debug.Assert lp2 < nAvaliableDataCodewordCount
   '///
   kk = 0
   For K = 0 To 7
    bb = bEncodedBit(LP * 8 + K)
    bInterleavedBit(lp2 * 8 + K) = bb
    kk = kk Or (m_nPowerOfTwo(7 - K) And bb <> 0)
   Next K
   nPolynomial(nDataCodewordPerBlock + nECCodewordPerBlock + (i < nSmallBlockCount) - j - 1) = kk
   '///
   LP = LP + 1
   lp2 = lp2 + nECBlockCount
   If i >= nSmallBlockCount And j = nDataCodewordPerBlock - 2 Then lp2 = lp2 - nSmallBlockCount
  Next j
  '///calc EC
  m_objRS.Encode nPolynomial
  '///
  lp2 = nAvaliableDataCodewordCount + i
  For j = 0 To nECCodewordPerBlock - 1
   kk = nPolynomial(nECCodewordPerBlock - 1 - j)
   For K = 0 To 7
    bInterleavedBit(lp2 * 8 + K) = ((kk And m_nPowerOfTwo(7 - K)) <> 0) And 1&
   Next K
   lp2 = lp2 + nECBlockCount
  Next j
 Next i
 Debug.Assert LP = nAvaliableDataCodewordCount
 '///create bitmap
 nModuleSize = pModuleSize(nVersion)
 ReDim b(nModuleSize + 2, nModuleSize + 2)
 '///fill function pattern
 pFillFunctionPattern b, nVersion
 '///put data to temp buffer
 ReDim b2(nModuleSize - 1, nModuleSize - 1)
 LP = 0
 i = nModuleSize
 Do
  '///upward
  ii = i - 2
  If ii < 6 Then ii = ii - 1
  For j = nModuleSize - 1 To 0 Step -1
   If b(ii + 1, j) = 0 Then
    b2(ii + 1, j) = bInterleavedBit(LP)
    LP = LP + 1
    If LP >= nDataCodewordCount * 8& Then Exit Do
   End If
   If b(ii, j) = 0 Then
    b2(ii, j) = bInterleavedBit(LP)
    LP = LP + 1
    If LP >= nDataCodewordCount * 8& Then Exit Do
   End If
  Next j
  '///downward
  i = i - 4
  ii = i
  If ii < 6 Then ii = ii - 1
  For j = 0 To nModuleSize - 1
   If b(ii + 1, j) = 0 Then
    b2(ii + 1, j) = bInterleavedBit(LP)
    LP = LP + 1
    If LP >= nDataCodewordCount * 8& Then Exit Do
   End If
   If b(ii, j) = 0 Then
    b2(ii, j) = bInterleavedBit(LP)
    LP = LP + 1
    If LP >= nDataCodewordCount * 8& Then Exit Do
   End If
  Next j
  '///
 Loop While i >= 5
 Debug.Assert LP >= nDataCodewordCount * 8&
 '///mask data
 nMaskType = pMaskData(b2, nModuleSize, nMaskType)
 '///get back
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   bb = b(i, j)
   If bb = 0 Then bb = b2(i, j)
   b(i, j) = bb And 1&
  Next i
 Next j
 '///fill version info
 If nVersion >= 7 Then
  kk = (nVersion * 4096&) Or pGF2PolynomialDivide(nVersion, 7973, 4096, 12)
  For i = 0 To 5
   For j = 0 To 2
    bb = (kk And m_nPowerOfTwo(i * 3 + j)) <> 0 And 1&
    b(i, nModuleSize - 11 + j) = bb
    b(nModuleSize - 11 + j, i) = bb
   Next j
  Next i
 End If
 '///fill format info
 nECLevel = (nECLevel - 1) Xor 1
 K = (nECLevel * 8&) Or nMaskType
 kk = pGF2PolynomialDivide(K, 1335, 1024, 10)
 'XOR mask      101010000010010
 bb = (nECLevel And 2&) = 0 And 1&
 b(0, 8) = bb: b(8, nModuleSize - 1) = bb
 bb = nECLevel And 1&
 b(1, 8) = bb: b(8, nModuleSize - 2) = bb
 bb = (nMaskType And 4&) = 0 And 1&
 b(2, 8) = bb: b(8, nModuleSize - 3) = bb
 bb = (nMaskType And 2&) <> 0 And 1&
 b(3, 8) = bb: b(8, nModuleSize - 4) = bb
 bb = (nMaskType And 1&) = 0 And 1&
 b(4, 8) = bb: b(8, nModuleSize - 5) = bb
 bb = (kk And 512) <> 0 And 1&
 b(5, 8) = bb: b(8, nModuleSize - 6) = bb
 bb = (kk And 256) <> 0 And 1&
 b(7, 8) = bb: b(8, nModuleSize - 7) = bb
 bb = (kk And 128) <> 0 And 1&
 b(8, 8) = bb: b(8, nModuleSize - 8) = bb
 bb = (kk And 64) <> 0 And 1&
 b(8, 7) = bb: b(nModuleSize - 7, 8) = bb
 bb = (kk And 32) <> 0 And 1&
 b(8, 5) = bb: b(nModuleSize - 6, 8) = bb
 bb = (kk And 16) = 0 And 1&
 b(8, 4) = bb: b(nModuleSize - 5, 8) = bb
 bb = (kk And 8) <> 0 And 1&
 b(8, 3) = bb: b(nModuleSize - 4, 8) = bb
 bb = (kk And 4) <> 0 And 1&
 b(8, 2) = bb: b(nModuleSize - 3, 8) = bb
 bb = (kk And 2) = 0 And 1&
 b(8, 1) = bb: b(nModuleSize - 2, 8) = bb
 bb = kk And 1
 b(8, 0) = bb: b(nModuleSize - 1, 8) = bb
 '///over
 Set Encode = ByteArrayToPicture(VarPtr(b(0, 0)), nModuleSize + 3, nModuleSize + 3, 4, 4, 1, 1)
End If
'///
End Function

'internal function
Friend Function pGF2PolynomialDivide(ByVal f As Long, ByVal G As Long, ByVal nOrder As Long, ByVal nTimes As Long) As Long
Do While nTimes > 0
 f = f + f
 If f And nOrder Then f = (f Xor G) And (nOrder - 1)
 nTimes = nTimes - 1
Loop
pGF2PolynomialDivide = f
End Function

'internal function
Friend Function pMaskData(ByRef b() As Byte, ByVal nModuleSize As Long, ByVal nMaskType As Long) As Long
Dim i As Long, j As Long
'///
If nMaskType < 0 Or nMaskType > 7 Then
 'TODO: find optimal mask
 nMaskType = 0
End If
'///
Select Case nMaskType
Case 0
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor ((i Xor j Xor 1) And 1&)
  Next i
 Next j
Case 1
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor ((j Xor 1) And 1&)
  Next i
 Next j
Case 2
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor ((i Mod 3 = 0) And 1&)
  Next i
 Next j
Case 3
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor (((i + j) Mod 3 = 0) And 1&)
  Next i
 Next j
Case 4
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor (((j \ 2 + i \ 3) Xor 1) And 1&)
  Next i
 Next j
Case 5
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor (((i And j And 1) + ((i * j) Mod 3)) = 0 And 1&)
  Next i
 Next j
Case 6
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor (((i And j And 1) + ((i * j) Mod 3) + 1) And 1&)
  Next i
 Next j
Case 7
 For j = 0 To nModuleSize - 1
  For i = 0 To nModuleSize - 1
   b(i, j) = b(i, j) Xor ((((i Xor j) And 1) + ((i * j) Mod 3) + 1) And 1&)
  Next i
 Next j
End Select
'///
pMaskData = nMaskType
End Function

'internal function
Friend Sub pFillFunctionPattern(ByRef b() As Byte, ByVal nVersion As Long)
Dim m As Long
Dim i As Long, j As Long, K As Long, L As Long
Dim s As Long
'///
Select Case nVersion
Case -4 To -1
 'TODO:
Case 1 To 40
 m = pModuleSize(nVersion)
 '///finder pattern
 For j = 0 To 7
  For i = 0 To 7
   If ((i = 1 Or i = 5) And j >= 1 And j <= 5) _
   Or ((j = 1 Or j = 5) And i >= 1 And i <= 5) Or i = 7 Or j = 7 Then
    b(i, j) = 2
    b(m - 1 - i, j) = 2
    b(i, m - 1 - j) = 2
   Else
    b(i, j) = 1
    b(m - 1 - i, j) = 1
    b(i, m - 1 - j) = 1
   End If
  Next i
 Next j
 '///format information
 For i = 0 To 7
  b(i, 8) = 2
  b(8, i) = 2
  b(m - 1 - i, 8) = 2
  b(8, m - 1 - i) = 2
 Next i
 b(8, 8) = 2
 b(8, m - 8) = 1
 '///version information
 If nVersion >= 7 Then
  For j = m - 11 To m - 9
   For i = 0 To 6
    b(i, j) = 2
    b(j, i) = 2
   Next i
  Next j
 End If
 '///timing pattern
 For i = 8 To m - 9
  b(i, 6) = 1 + (i And 1&)
  b(6, i) = 1 + (i And 1&)
 Next i
 '///alignment pattern
 If nVersion > 1 Then
  s = m_nAlignmentPatternStep(nVersion)
  For j = m - 7 To 18 Step -s
   For i = m - 7 To 18 Step -s
    For L = -2 To 2
     For K = -2 To 2
      b(i + K, j + L) = 2 + (K = 2 Or K = -2 Or L = 2 Or L = -2 Or (K Or L) = 0)
     Next K
    Next L
   Next i
  Next j
  For i = m - 7 - s To 18 Step -s
   For L = -2 To 2
    For K = -2 To 2
     j = 2 + (K = 2 Or K = -2 Or L = 2 Or L = -2 Or (K Or L) = 0)
     b(i + K, 6 + L) = j
     b(6 + L, i + K) = j
    Next K
   Next L
  Next i
 End If
 '///over
End Select
End Sub

Friend Function pAlphaNumericToNumber(ByVal b As Byte) As Long
Select Case b
Case &H30& To &H39&
 pAlphaNumericToNumber = b - 48
Case &H41& To &H5A&
 pAlphaNumericToNumber = b - 55
Case &H20
 pAlphaNumericToNumber = 36
Case &H24
 pAlphaNumericToNumber = 37
Case &H25
 pAlphaNumericToNumber = 38
Case &H2A
 pAlphaNumericToNumber = 39
Case &H2B
 pAlphaNumericToNumber = 40
Case &H2D
 pAlphaNumericToNumber = 41
Case &H2E
 pAlphaNumericToNumber = 42
Case &H2F
 pAlphaNumericToNumber = 43
Case &H3A
 pAlphaNumericToNumber = 44
End Select
End Function

'internal function
'0=byte
'1=number
'2=alphanumeric
Friend Function pCheckExclusiveSubset(ByVal b As Byte) As Long
Select Case b
Case &H30& To &H39&
 pCheckExclusiveSubset = 1
Case &H41& To &H5A&, &H20, &H24, &H25, &H2A, &H2B, &H2D, &H2E&, &H2F, &H3A
 pCheckExclusiveSubset = 2
End Select
End Function

'internal function
Friend Function pCheckKanji(ByVal b1 As Byte, ByVal b2 As Byte) As Boolean
If b2 >= &H40& Then
 Select Case b1
 Case &H81& To &H9F&, &HE0& To &HEA&
  pCheckKanji = True
 Case &HEB&
  pCheckKanji = b2 <= &HBF&
 End Select
End If
End Function

'internal function
'bOutput: 0-based, should be big enough, say 32768
'bInput: 0-based byte data
'nSize: input data size
'nVersion: symbol version, should be 1 to 40
'return value: bit size
Friend Function pEncodeToBitArray(ByRef bOutput() As Byte, ByRef bInput() As Byte, ByVal nSize As Long, ByVal nVersion As Long, Optional ByVal bCheckSizeOnly As Boolean) As Long
'0=byte
'1=number
'2=alphanumeric
'3=kanji
Dim nEncodingMode() As Byte
Dim tNode() As typeDynamicProgrammingNode
Dim nMaxSize(3) As Long
Dim i As Long, j As Long, K As Long, m As Long
Dim nCost As Long, nCount As Long
Dim LP As Long, lp2 As Long
'///a trivial case
If nSize <= 0 Then
 bOutput(0) = 0
 bOutput(1) = 0
 bOutput(2) = 0
 bOutput(3) = 0
 pEncodeToBitArray = 4
 Exit Function
End If
'///
Select Case nVersion
Case Is < 10
 nVersion = 2
 nMaxSize(0) = 255
 nMaxSize(1) = 1023
 nMaxSize(2) = 511
 nMaxSize(3) = 255
Case Is < 27
 nVersion = 4
 nMaxSize(0) = 65535
 nMaxSize(1) = 4095
 nMaxSize(2) = 2047
 nMaxSize(3) = 1023
Case Else
 nVersion = 6
 nMaxSize(0) = 65535
 nMaxSize(1) = 16383
 nMaxSize(2) = 8191
 nMaxSize(3) = 4095
End Select
'///use dynamic programming to determing best encoding mode (.)
ReDim tNode(3, nSize - 1)
For LP = 0 To nSize - 1
 '///check byte mode
 If LP > 0 Then
  nCost = &H7FFFFFFF
  For i = 0 To 3
   K = tNode(i, LP - 1).nCost + 8
   If i <> 0 Or tNode(i, LP - 1).nCount >= nMaxSize(0) Then K = K + 12 + (nVersion And 4) * 2
   If K < nCost Then
    j = i
    nCost = K
    nCount = tNode(i, LP - 1).nCount And i = 0 And tNode(i, LP - 1).nCount < nMaxSize(0)
   End If
  Next i
 Else
  j = -1
  nCost = 20 + (nVersion And 4) * 2
  nCount = 0
 End If
 tNode(0, LP).nPrevMode = j
 tNode(0, LP).nCost = nCost
 tNode(0, LP).nCount = nCount + 1
 '///check number mode
 lp2 = pCheckExclusiveSubset(bInput(LP))
 If lp2 = 1 Then
  If LP > 0 Then
   nCost = &H7FFFFFFF
   For i = 0 To 3
    K = tNode(i, LP - 1).nCost
    If i <> 1 Or tNode(i, LP - 1).nCount >= nMaxSize(1) Then
     K = K + 16 + nVersion
    Else
     K = K + 3
     If (tNode(i, LP - 1).nCount Mod 3) = 0 Then K = K + 1
    End If
    If K < nCost Then
     j = i
     nCost = K
     nCount = tNode(i, LP - 1).nCount And i = 1 And tNode(i, LP - 1).nCount < nMaxSize(1)
    End If
   Next i
  Else
   j = -1
   nCost = 16 + nVersion
   nCount = 0
  End If
  tNode(1, LP).nPrevMode = j
  tNode(1, LP).nCost = nCost
  tNode(1, LP).nCount = nCount + 1
 Else
  tNode(1, LP).nCost = &H70000000
 End If
 '///check alphanumeric mode
 If lp2 > 0 Then
  If LP > 0 Then
   nCost = &H7FFFFFFF
   For i = 0 To 3
    K = tNode(i, LP - 1).nCost
    If i <> 2 Or tNode(i, LP - 1).nCount >= nMaxSize(2) Then
     K = K + 17 + nVersion
    Else
     K = K + 5
     If (tNode(i, LP - 1).nCount And 1) = 0 Then K = K + 1
    End If
    If K < nCost Then
     j = i
     nCost = K
     nCount = tNode(i, LP - 1).nCount And i = 2 And tNode(i, LP - 1).nCount < nMaxSize(2)
    End If
   Next i
  Else
   j = -1
   nCost = 17 + nVersion
   nCount = 0
  End If
  tNode(2, LP).nPrevMode = j
  tNode(2, LP).nCost = nCost
  tNode(2, LP).nCount = nCount + 1
 Else
  tNode(2, LP).nCost = &H70000000
 End If
 '///check kanji mode
 If LP = 0 Then
  tNode(3, LP).nCost = &H70000000
 ElseIf Not pCheckKanji(bInput(LP - 1), bInput(LP)) Then
  tNode(3, LP).nCost = &H70000000
 Else
  If LP > 1 Then
   nCost = &H7FFFFFFF
   For i = 0 To 3
    K = tNode(i, LP - 2).nCost + 13
    If i <> 3 Or tNode(i, LP - 2).nCount >= nMaxSize(3) Then K = K + 10 + nVersion
    If K < nCost Then
     j = i
     nCost = K
     nCount = tNode(i, LP - 2).nCount And i = 3 And tNode(i, LP - 2).nCount < nMaxSize(3)
    End If
   Next i
  Else
   j = -1
   nCost = 23 + nVersion
   nCount = 0
  End If
  tNode(3, LP).nPrevMode = j
  tNode(3, LP).nCost = nCost
  tNode(3, LP).nCount = nCount + 1
 End If
Next LP
'///get encoding mode
nCost = &H7FFFFFFF
For i = 0 To 3
 K = tNode(i, nSize - 1).nCost
 If K < nCost Then
  j = i
  nCost = K
 End If
Next i
'Debug.Print nCost 'debug
If bCheckSizeOnly Then
 pEncodeToBitArray = nCost + 4
 Exit Function
End If
'///
ReDim nEncodingMode(nSize - 1)
LP = nSize - 1
Do
 K = tNode(j, LP).nPrevMode
 nEncodingMode(LP) = j
 LP = LP - 1
 If j = 3 Then
  If LP < 0 Then
   'something goes wrong
   Debug.Assert False
   Exit Function
  End If
  nEncodingMode(LP) = j
  LP = LP - 1
 End If
 j = K
Loop While LP >= 0
Erase tNode
'///
nMaxSize(3) = nMaxSize(3) * 2
LP = 0
lp2 = 0
Do
 j = nEncodingMode(LP)
 For nCount = 1 To nSize - 1 - LP
  If nEncodingMode(LP + nCount) <> j Then Exit For
  If nCount >= nMaxSize(j) Then Exit For
 Next nCount
 '///
 'Debug.Print j, nCount 'debug
 '///
 Select Case j
 Case 0 'byte
  bOutput(lp2) = 0
  bOutput(lp2 + 1) = 1
  bOutput(lp2 + 2) = 0
  bOutput(lp2 + 3) = 0
  lp2 = lp2 + 4
  '///
  m = 8 + (nVersion And 4) * 2
  For i = 0 To m - 1
   bOutput(lp2 + i) = (nCount And m_nPowerOfTwo(m - 1 - i)) <> 0 And 1
  Next i
  lp2 = lp2 + m
  '///
  For j = 0 To nCount - 1
   K = bInput(LP + j)
   For i = 0 To 7
    bOutput(lp2 + i) = (K And m_nPowerOfTwo(7 - i)) <> 0 And 1
   Next i
   lp2 = lp2 + 8
  Next j
 Case 1 'number
  bOutput(lp2) = 0
  bOutput(lp2 + 1) = 0
  bOutput(lp2 + 2) = 0
  bOutput(lp2 + 3) = 1
  lp2 = lp2 + 4
  '///
  m = 8 + nVersion
  For i = 0 To m - 1
   bOutput(lp2 + i) = (nCount And m_nPowerOfTwo(m - 1 - i)) <> 0 And 1
  Next i
  lp2 = lp2 + m
  '///
  For j = 0 To nCount - 3 Step 3
   K = (bInput(LP + j) And &HF&) * 100& + (bInput(LP + j + 1) And &HF&) * 10& + (bInput(LP + j + 2) And &HF&)
   For i = 0 To 9
    bOutput(lp2 + i) = (K And m_nPowerOfTwo(9 - i)) <> 0 And 1
   Next i
   lp2 = lp2 + 10
  Next j
  '///
  Select Case nCount Mod 3
  Case 1
   K = (bInput(LP + nCount - 1) And &HF&)
   For i = 0 To 3
    bOutput(lp2 + i) = (K And m_nPowerOfTwo(3 - i)) <> 0 And 1
   Next i
   lp2 = lp2 + 4
  Case 2
   K = (bInput(LP + nCount - 2) And &HF&) * 10& + (bInput(LP + nCount - 1) And &HF&)
   For i = 0 To 6
    bOutput(lp2 + i) = (K And m_nPowerOfTwo(6 - i)) <> 0 And 1
   Next i
   lp2 = lp2 + 7
  End Select
 Case 2 'alphanumeric
  bOutput(lp2) = 0
  bOutput(lp2 + 1) = 0
  bOutput(lp2 + 2) = 1
  bOutput(lp2 + 3) = 0
  lp2 = lp2 + 4
  '///
  m = 7 + nVersion
  For i = 0 To m - 1
   bOutput(lp2 + i) = (nCount And m_nPowerOfTwo(m - 1 - i)) <> 0 And 1
  Next i
  lp2 = lp2 + m
  '///
  For j = 0 To nCount - 2 Step 2
   K = pAlphaNumericToNumber(bInput(LP + j)) * 45& + pAlphaNumericToNumber(bInput(LP + j + 1))
   For i = 0 To 10
    bOutput(lp2 + i) = (K And m_nPowerOfTwo(10 - i)) <> 0 And 1
   Next i
   lp2 = lp2 + 11
  Next j
  '///
  If nCount And 1& Then
   K = pAlphaNumericToNumber(bInput(LP + nCount - 1))
   For i = 0 To 5
    bOutput(lp2 + i) = (K And m_nPowerOfTwo(5 - i)) <> 0 And 1
   Next i
   lp2 = lp2 + 6
  End If
 Case 3 'kanji
  bOutput(lp2) = 1
  bOutput(lp2 + 1) = 0
  bOutput(lp2 + 2) = 0
  bOutput(lp2 + 3) = 0
  lp2 = lp2 + 4
  '///
  Debug.Assert (nCount And 1&) = 0
  '///
  m = 6 + nVersion
  For i = 0 To m - 1
   bOutput(lp2 + i) = (nCount And m_nPowerOfTwo(m - i)) <> 0 And 1
  Next i
  lp2 = lp2 + m
  '///
  For j = 0 To nCount - 2 Step 2
   i = bInput(LP + j)
   Select Case i
   Case &H81& To &H9F&
    i = i - &H81&
   Case &HE0& To &HEB&
    i = i - &HC1&
   Case Else
    Debug.Assert False
    Exit Function
   End Select
   K = bInput(LP + j + 1) - &H40&
   Debug.Assert K >= 0
   K = K + i * &HC0&
   Debug.Assert K < &H2000&
   For i = 0 To 12
    bOutput(lp2 + i) = (K And m_nPowerOfTwo(12 - i)) <> 0 And 1
   Next i
   lp2 = lp2 + 13
  Next j
 End Select
 '///
 LP = LP + nCount
Loop While LP < nSize
'///over
'Debug.Print lp2 'debug
bOutput(lp2) = 0
bOutput(lp2 + 1) = 0
bOutput(lp2 + 2) = 0
bOutput(lp2 + 3) = 0
pEncodeToBitArray = lp2 + 4
End Function


