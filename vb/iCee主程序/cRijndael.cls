VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRijndael"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'

Option Explicit
'Îå×ÓÆå
#Const SUPPORT_LEVEL = 0     'Default=0
#Const COMPILE_CONSTANTS = 0 'Default=0
Private Te0(255)      As Long
Private Te1(255)      As Long
Private Te2(255)      As Long
Private Te3(255)      As Long
Private Te4(255)      As Long

Private Td0(255)      As Long
Private Td1(255)      As Long
Private Td2(255)      As Long
Private Td3(255)      As Long
Private Td4(255)      As Long

#If SUPPORT_LEVEL Then
Private rco(28)       As Long
#Else
Private rco(9)        As Long
#End If

'Key schedule arrays
Private Nr            As Long 'Number of rounds [For 128 bit block, Nr = {10, 12, 14} for 128, 192, 256 bit cipher key]
#If SUPPORT_LEVEL Then
Private fkey(119)     As Long 'Nb*(Nr + 1)
Private rkey(119)     As Long 'Nb*(Nr + 1)
#Else
Private fkey(59)      As Long 'Nb*(Nr + 1)
Private rkey(59)      As Long 'Nb*(Nr + 1)
#End If

'For file encryption, this is the maximum amount of memory (in bytes) allowed for file data
Private Const MaxFileChunkSize As Long = 4000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


'Decryption Key Scheduler.  Calculate rkey() decryption keys based on fkey() and Nr.
'Nb is the number of 32 bit units in the block size.
Private Sub CreateDecryptionKeys(Nb As Long)
    Dim i    As Long
    Dim j    As Long
    Dim K    As Long
    Dim s(3) As Byte

    'Invert the order of the round keys
    i = 0
    j = Nb * Nr
    For K = 0 To Nr
        CopyMemory rkey(i), fkey(j), Nb * 4&
        i = i + Nb
        j = j - Nb
    Next K

    'Apply the inverse MixColumn transform to all round keys except the first and the last
    For i = Nb To Nb * Nr - 1
        CopyMemory s(0), rkey(i), 4&
        rkey(i) = Td0(Te4(s(0)) And &HFF&) Xor _
                  Td1(Te4(s(1)) And &HFF&) Xor _
                  Td2(Te4(s(2)) And &HFF&) Xor _
                  Td3(Te4(s(3)) And &HFF&)
    Next i
End Sub


'Key Scheduler. Expand the cipher key into the encryption key schedule.
'pass(0 ... n-1) contains the cipher key, where n = {16, 20, 24, 28, 32} , depending on KeyBits.
'If SUPPORT_LEVEL is non-zero, you must specify BlockBits for the block size you will be using.
'Returns zero on success.
#If SUPPORT_LEVEL Then
Private Function SetCipherKey(pass() As Byte, KeyBits As Long, BlockBits As Long) As Long
#Else
Private Function SetCipherKey(pass() As Byte, KeyBits As Long) As Long
#End If
    Dim i    As Long
    Dim j    As Long
    Dim s(3) As Byte

#If SUPPORT_LEVEL Then
    Select Case BlockBits
    Case 128
#End If

    '128 bit block size
    Select Case KeyBits
    Case 128
        i = 4
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 9
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 4) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 3) Xor fkey(i)
            fkey(i + 2) = fkey(i - 2) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 1) Xor fkey(i + 2)
            i = i + 4
        Next j
        Nr = 10
#If SUPPORT_LEVEL = 2 Then
    Case 160
        i = 5
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 5) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 4) Xor fkey(i)
            fkey(i + 2) = fkey(i - 3) Xor fkey(i + 1)
            If j = 8 Then Exit Do
            fkey(i + 3) = fkey(i - 2) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 1) Xor fkey(i + 3)
            i = i + 5
            j = j + 1
        Loop
        Nr = 11
'Debug.Assert i + 3 = (Nr + 1) * 4
#End If
    Case 192
        i = 6
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 6) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 5) Xor fkey(i)
            fkey(i + 2) = fkey(i - 4) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 3) Xor fkey(i + 2)
            If j = 7 Then Exit Do
            fkey(i + 4) = fkey(i - 2) Xor fkey(i + 3)
            fkey(i + 5) = fkey(i - 1) Xor fkey(i + 4)
            i = i + 6
            j = j + 1
        Loop
        Nr = 12
'Debug.Assert i + 4 = (Nr + 1) * 4
#If SUPPORT_LEVEL = 2 Then
    Case 224
        i = 7
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 6
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 7) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 6) Xor fkey(i)
            fkey(i + 2) = fkey(i - 5) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 4) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 3) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 2) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 1) Xor fkey(i + 5)
            i = i + 7
        Next j
        Nr = 13
'Debug.Assert i = (Nr + 1) * 4
#End If
    Case 256
        i = 8
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 8) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 7) Xor fkey(i)
            fkey(i + 2) = fkey(i - 6) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 5) Xor fkey(i + 2)
            If j = 6 Then Exit Do
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 4) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 3) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 2) Xor fkey(i + 5)
            fkey(i + 7) = fkey(i - 1) Xor fkey(i + 6)
            i = i + 8
            j = j + 1
        Loop
        Nr = 14
'Debug.Assert i + 4 = (Nr + 1) * 4
    Case Else
        ERR.Raise 1, , "cRijndael.SetCipherKey - Illegal KeyBits Value"
        SetCipherKey = 1
        Exit Function
    End Select
    CreateDecryptionKeys 4

#If SUPPORT_LEVEL Then
#If SUPPORT_LEVEL = 2 Then
    Case 160

    '160 bit block size
    Select Case KeyBits
    Case 128
        i = 4
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 13
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 4) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 3) Xor fkey(i)
            fkey(i + 2) = fkey(i - 2) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 1) Xor fkey(i + 2)
            i = i + 4
        Next j
        Nr = 11
'Debug.Assert i = (Nr + 1) * 5
    Case 160
        i = 5
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 10
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 5) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 4) Xor fkey(i)
            fkey(i + 2) = fkey(i - 3) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 2) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 1) Xor fkey(i + 3)
            i = i + 5
        Next j
        Nr = 11
'Debug.Assert i = (Nr + 1) * 5
    Case 192
        i = 6
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 6) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 5) Xor fkey(i)
            fkey(i + 2) = fkey(i - 4) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 3) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 2) Xor fkey(i + 3)
            If j = 9 Then Exit Do
            fkey(i + 5) = fkey(i - 1) Xor fkey(i + 4)
            i = i + 6
            j = j + 1
        Loop
        Nr = 12
'Debug.Assert i + 5 = (Nr + 1) * 5
    Case 224
        i = 7
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 8
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 7) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 6) Xor fkey(i)
            fkey(i + 2) = fkey(i - 5) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 4) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 3) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 2) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 1) Xor fkey(i + 5)
            i = i + 7
        Next j
        Nr = 13
'Debug.Assert i = (Nr + 1) * 5
    Case 256
        i = 8
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 8) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 7) Xor fkey(i)
            fkey(i + 2) = fkey(i - 6) Xor fkey(i + 1)
            If j = 8 Then Exit Do
            fkey(i + 3) = fkey(i - 5) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 4) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 3) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 2) Xor fkey(i + 5)
            fkey(i + 7) = fkey(i - 1) Xor fkey(i + 6)
            i = i + 8
            j = j + 1
        Loop
        Nr = 14
'Debug.Assert i + 3 = (Nr + 1) * 5
    Case Else
        ERR.Raise 1, , "cRijndael.SetCipherKey - Illegal KeyBits Value"
        SetCipherKey = 1
        Exit Function
    End Select
    CreateDecryptionKeys 5
#End If
    Case 192

    '192 bit block size
    Select Case KeyBits
    Case 128
        i = 4
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 4) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 3) Xor fkey(i)
            If j = 18 Then Exit Do
            fkey(i + 2) = fkey(i - 2) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 1) Xor fkey(i + 2)
            i = i + 4
            j = j + 1
        Loop
        Nr = 12
'Debug.Assert i + 2 = (Nr + 1) * 6
#If SUPPORT_LEVEL = 2 Then
    Case 160
        i = 5
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 5) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 4) Xor fkey(i)
            fkey(i + 2) = fkey(i - 3) Xor fkey(i + 1)
            If j = 14 Then Exit Do
            fkey(i + 3) = fkey(i - 2) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 1) Xor fkey(i + 3)
            i = i + 5
            j = j + 1
        Loop
        Nr = 12
'Debug.Assert i + 3 = (Nr + 1) * 6
#End If
    Case 192
        i = 6
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 11
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 6) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 5) Xor fkey(i)
            fkey(i + 2) = fkey(i - 4) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 3) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 2) Xor fkey(i + 3)
            fkey(i + 5) = fkey(i - 1) Xor fkey(i + 4)
            i = i + 6
        Next j
        Nr = 12
'Debug.Assert i = (Nr + 1) * 6
#If SUPPORT_LEVEL = 2 Then
    Case 224 '(Nr+1)*Nb/Nk  (when to exit)
        i = 7
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 10
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 7) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 6) Xor fkey(i)
            fkey(i + 2) = fkey(i - 5) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 4) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 3) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 2) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 1) Xor fkey(i + 5)
            i = i + 7
        Next j
        Nr = 13
'Debug.Assert i = (Nr + 1) * 6
#End If
    Case 256
        i = 8
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 8) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 7) Xor fkey(i)
            If j = 10 Then Exit Do
            fkey(i + 2) = fkey(i - 6) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 5) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 4) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 3) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 2) Xor fkey(i + 5)
            fkey(i + 7) = fkey(i - 1) Xor fkey(i + 6)
            i = i + 8
            j = j + 1
        Loop
        Nr = 14
'Debug.Assert i + 2 = (Nr + 1) * 6
    Case Else
        ERR.Raise 1, , "cRijndael.SetCipherKey - Illegal KeyBits Value"
        SetCipherKey = 1
        Exit Function
    End Select
    CreateDecryptionKeys 6

#If SUPPORT_LEVEL = 2 Then
    Case 224

    '224 bit block size
    Select Case KeyBits
    Case 128
        i = 4
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 4) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 3) Xor fkey(i)
            If j = 23 Then Exit Do
            fkey(i + 2) = fkey(i - 2) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 1) Xor fkey(i + 2)
            i = i + 4
            j = j + 1
        Loop
        Nr = 13
'Debug.Assert i + 2 = (Nr + 1) * 7
    Case 160
        i = 5
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 5) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 4) Xor fkey(i)
            fkey(i + 2) = fkey(i - 3) Xor fkey(i + 1)
            If j = 18 Then Exit Do
            fkey(i + 3) = fkey(i - 2) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 1) Xor fkey(i + 3)
            i = i + 5
            j = j + 1
        Loop
        Nr = 13
'Debug.Assert i + 3 = (Nr + 1) * 7
    Case 192
        i = 6
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 6) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 5) Xor fkey(i)
            If j = 15 Then Exit Do
            fkey(i + 2) = fkey(i - 4) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 3) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 2) Xor fkey(i + 3)
            fkey(i + 5) = fkey(i - 1) Xor fkey(i + 4)
            i = i + 6
            j = j + 1
        Loop
        Nr = 13
'Debug.Assert i + 2 = (Nr + 1) * 7
    Case 224
        i = 7
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 12
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 7) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 6) Xor fkey(i)
            fkey(i + 2) = fkey(i - 5) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 4) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 3) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 2) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 1) Xor fkey(i + 5)
            i = i + 7
        Next j
        Nr = 13
'Debug.Assert i = (Nr + 1) * 7
    Case 256
        i = 8
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 8) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            If j = 12 Then Exit Do
            fkey(i + 1) = fkey(i - 7) Xor fkey(i)
            fkey(i + 2) = fkey(i - 6) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 5) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 4) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 3) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 2) Xor fkey(i + 5)
            fkey(i + 7) = fkey(i - 1) Xor fkey(i + 6)
            i = i + 8
            j = j + 1
        Loop
        Nr = 14
'Debug.Assert i + 1 = (Nr + 1) * 7
    Case Else
        ERR.Raise 1, , "cRijndael.SetCipherKey - Illegal KeyBits Value"
        SetCipherKey = 1
        Exit Function
    End Select
    CreateDecryptionKeys 7
#End If

    Case 256

    '256 bit block size
    Select Case KeyBits
    Case 128
        i = 4
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 28
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 4) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 3) Xor fkey(i)
            fkey(i + 2) = fkey(i - 2) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 1) Xor fkey(i + 2)
            i = i + 4
        Next j
        Nr = 14
'Debug.Assert i = (Nr + 1) * 8
#If SUPPORT_LEVEL = 2 Then
    Case 160
        i = 5
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 22
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 5) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 4) Xor fkey(i)
            fkey(i + 2) = fkey(i - 3) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 2) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 1) Xor fkey(i + 3)
            i = i + 5
        Next j
        Nr = 14
'Debug.Assert i = (Nr + 1) * 8
#End If
    Case 192
        i = 6
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 18
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 6) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 5) Xor fkey(i)
            fkey(i + 2) = fkey(i - 4) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 3) Xor fkey(i + 2)
            fkey(i + 4) = fkey(i - 2) Xor fkey(i + 3)
            fkey(i + 5) = fkey(i - 1) Xor fkey(i + 4)
            i = i + 6
        Next j
        Nr = 14
'Debug.Assert i = (Nr + 1) * 8
#If SUPPORT_LEVEL = 2 Then
    Case 224
        i = 7
        j = 0
        CopyMemory fkey(0), pass(0), 4& * i
        Do
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 7) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            If j = 16 Then Exit Do
            fkey(i + 1) = fkey(i - 6) Xor fkey(i)
            fkey(i + 2) = fkey(i - 5) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 4) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 3) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 2) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 1) Xor fkey(i + 5)
            i = i + 7
            j = j + 1
        Loop
        Nr = 14
'Debug.Assert i + 1 = (Nr + 1) * 8
#End If
    Case 256
        i = 8
        CopyMemory fkey(0), pass(0), 4& * i
        For j = 0 To 13
            CopyMemory s(0), fkey(i - 1), 4&
            fkey(i) = fkey(i - 8) Xor (Te4(s(0)) And &HFF000000) _
                                  Xor (Te4(s(3)) And &HFF0000) _
                                  Xor (Te4(s(2)) And &HFF00&) _
                                  Xor (Te4(s(1)) And &HFF&) _
                                  Xor rco(j)
            fkey(i + 1) = fkey(i - 7) Xor fkey(i)
            fkey(i + 2) = fkey(i - 6) Xor fkey(i + 1)
            fkey(i + 3) = fkey(i - 5) Xor fkey(i + 2)
            CopyMemory s(0), fkey(i + 3), 4&
            fkey(i + 4) = fkey(i - 4) Xor (Te4(s(3)) And &HFF000000) _
                                      Xor (Te4(s(2)) And &HFF0000) _
                                      Xor (Te4(s(1)) And &HFF00&) _
                                      Xor (Te4(s(0)) And &HFF&)
            fkey(i + 5) = fkey(i - 3) Xor fkey(i + 4)
            fkey(i + 6) = fkey(i - 2) Xor fkey(i + 5)
            fkey(i + 7) = fkey(i - 1) Xor fkey(i + 6)
            i = i + 8
        Next j
        Nr = 14
'Debug.Assert i = (Nr + 1) * 8
    Case Else
        ERR.Raise 1, , "cRijndael.SetCipherKey - Illegal KeyBits Value"
        SetCipherKey = 1
        Exit Function
    End Select
    CreateDecryptionKeys 8

    Case Else
        ERR.Raise 1, , "cRijndael.SetCipherKey - Illegal BlockBits Value"
        SetCipherKey = 1
        Exit Function
    End Select
#End If
End Function

#If SUPPORT_LEVEL Then
Private Function SetCipherKeyString(PassPhrase As String, KeyBits As Long, BlockBits As Long) As Long
    Dim pass() As Byte

    pass = StrConv(PassPhrase, vbFromUnicode)
    ReDim Preserve pass(31)
    SetCipherKeyString = SetCipherKey(pass, KeyBits, BlockBits)
End Function
#Else
Private Function SetCipherKeyString(PassPhrase As String, KeyBits As Long) As Long
    Dim pass() As Byte

    pass = StrConv(PassPhrase, vbFromUnicode)
    ReDim Preserve pass(31)
    SetCipherKeyString = SetCipherKey(pass, KeyBits)
End Function
#End If


'Encrypt a 128 bit block.  plaintext(p ... p+15) is input, ciphertext(q ... q+15) is output.
'plaintext and ciphertext can be the same array.  Will crash if ciphertext(q ... q+15) is not allocated.
Private Sub BlockEncrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim s(15) As Byte

    CopyMemory t0, plaintext(p + 0), 4&
    CopyMemory t1, plaintext(p + 4), 4&
    CopyMemory t2, plaintext(p + 8), 4&
    CopyMemory t3, plaintext(p + 12), 4&
    t0 = t0 Xor fkey(0)
    t1 = t1 Xor fkey(1)
    t2 = t2 Xor fkey(2)
    t3 = t3 Xor fkey(3)
    K = 4

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        t0 = Te0(s(0)) Xor Te1(s(5)) Xor Te2(s(10)) Xor Te3(s(15)) Xor fkey(K + 0)
        t1 = Te0(s(4)) Xor Te1(s(9)) Xor Te2(s(14)) Xor Te3(s(3)) Xor fkey(K + 1)
        t2 = Te0(s(8)) Xor Te1(s(13)) Xor Te2(s(2)) Xor Te3(s(7)) Xor fkey(K + 2)
        t3 = Te0(s(12)) Xor Te1(s(1)) Xor Te2(s(6)) Xor Te3(s(11)) Xor fkey(K + 3)
        K = K + 4
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    t0 = (Te4(s(0)) And &HFF&) Xor (Te4(s(5)) And &HFF00&) Xor (Te4(s(10)) And &HFF0000) Xor (Te4(s(15)) And &HFF000000) Xor fkey(K + 0)
    t1 = (Te4(s(4)) And &HFF&) Xor (Te4(s(9)) And &HFF00&) Xor (Te4(s(14)) And &HFF0000) Xor (Te4(s(3)) And &HFF000000) Xor fkey(K + 1)
    t2 = (Te4(s(8)) And &HFF&) Xor (Te4(s(13)) And &HFF00&) Xor (Te4(s(2)) And &HFF0000) Xor (Te4(s(7)) And &HFF000000) Xor fkey(K + 2)
    t3 = (Te4(s(12)) And &HFF&) Xor (Te4(s(1)) And &HFF00&) Xor (Te4(s(6)) And &HFF0000) Xor (Te4(s(11)) And &HFF000000) Xor fkey(K + 3)
    CopyMemory ciphertext(Q + 0), t0, 4&
    CopyMemory ciphertext(Q + 4), t1, 4&
    CopyMemory ciphertext(Q + 8), t2, 4&
    CopyMemory ciphertext(Q + 12), t3, 4&
End Sub

'Decrypt a 128 bit block.  ciphertext(q ... q+15) is input, plaintext(p ... p+15) is output.
'plaintext and ciphertext can be the same array.  Will crash if plaintext(p ... p+15) is not allocated.
Private Sub BlockDecrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim s(15) As Byte

    CopyMemory t0, ciphertext(Q + 0), 4&
    CopyMemory t1, ciphertext(Q + 4), 4&
    CopyMemory t2, ciphertext(Q + 8), 4&
    CopyMemory t3, ciphertext(Q + 12), 4&
    t0 = t0 Xor rkey(0)
    t1 = t1 Xor rkey(1)
    t2 = t2 Xor rkey(2)
    t3 = t3 Xor rkey(3)
    K = 4

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        t0 = Td0(s(0)) Xor Td1(s(13)) Xor Td2(s(10)) Xor Td3(s(7)) Xor rkey(K + 0)
        t1 = Td0(s(4)) Xor Td1(s(1)) Xor Td2(s(14)) Xor Td3(s(11)) Xor rkey(K + 1)
        t2 = Td0(s(8)) Xor Td1(s(5)) Xor Td2(s(2)) Xor Td3(s(15)) Xor rkey(K + 2)
        t3 = Td0(s(12)) Xor Td1(s(9)) Xor Td2(s(6)) Xor Td3(s(3)) Xor rkey(K + 3)
        K = K + 4
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    t0 = (Td4(s(0)) And &HFF&) Xor (Td4(s(13)) And &HFF00&) Xor (Td4(s(10)) And &HFF0000) Xor (Td4(s(7)) And &HFF000000) Xor rkey(K + 0)
    t1 = (Td4(s(4)) And &HFF&) Xor (Td4(s(1)) And &HFF00&) Xor (Td4(s(14)) And &HFF0000) Xor (Td4(s(11)) And &HFF000000) Xor rkey(K + 1)
    t2 = (Td4(s(8)) And &HFF&) Xor (Td4(s(5)) And &HFF00&) Xor (Td4(s(2)) And &HFF0000) Xor (Td4(s(15)) And &HFF000000) Xor rkey(K + 2)
    t3 = (Td4(s(12)) And &HFF&) Xor (Td4(s(9)) And &HFF00&) Xor (Td4(s(6)) And &HFF0000) Xor (Td4(s(3)) And &HFF000000) Xor rkey(K + 3)
    CopyMemory plaintext(p + 0), t0, 4&
    CopyMemory plaintext(p + 4), t1, 4&
    CopyMemory plaintext(p + 8), t2, 4&
    CopyMemory plaintext(p + 12), t3, 4&
End Sub

#If SUPPORT_LEVEL Then
#If SUPPORT_LEVEL = 2 Then
'Encrypt a 160 bit block.  plaintext(p ... p+19) is input, ciphertext(q ... q+19) is output.
'plaintext and ciphertext can be the same array.  Will crash if ciphertext(q ... q+19) is not allocated.
Private Sub Block160Encrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim s(19) As Byte

    CopyMemory t0, plaintext(p + 0), 4&
    CopyMemory t1, plaintext(p + 4), 4&
    CopyMemory t2, plaintext(p + 8), 4&
    CopyMemory t3, plaintext(p + 12), 4&
    CopyMemory t4, plaintext(p + 16), 4&
    t0 = t0 Xor fkey(0)
    t1 = t1 Xor fkey(1)
    t2 = t2 Xor fkey(2)
    t3 = t3 Xor fkey(3)
    t4 = t4 Xor fkey(4)
    K = 5

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        t0 = Te0(s(0)) Xor Te1(s(5)) Xor Te2(s(10)) Xor Te3(s(15)) Xor fkey(K + 0)
        t1 = Te0(s(4)) Xor Te1(s(9)) Xor Te2(s(14)) Xor Te3(s(19)) Xor fkey(K + 1)
        t2 = Te0(s(8)) Xor Te1(s(13)) Xor Te2(s(18)) Xor Te3(s(3)) Xor fkey(K + 2)
        t3 = Te0(s(12)) Xor Te1(s(17)) Xor Te2(s(2)) Xor Te3(s(7)) Xor fkey(K + 3)
        t4 = Te0(s(16)) Xor Te1(s(1)) Xor Te2(s(6)) Xor Te3(s(11)) Xor fkey(K + 4)
        K = K + 5
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    t0 = (Te4(s(0)) And &HFF&) Xor (Te4(s(5)) And &HFF00&) Xor (Te4(s(10)) And &HFF0000) Xor (Te4(s(15)) And &HFF000000) Xor fkey(K + 0)
    t1 = (Te4(s(4)) And &HFF&) Xor (Te4(s(9)) And &HFF00&) Xor (Te4(s(14)) And &HFF0000) Xor (Te4(s(19)) And &HFF000000) Xor fkey(K + 1)
    t2 = (Te4(s(8)) And &HFF&) Xor (Te4(s(13)) And &HFF00&) Xor (Te4(s(18)) And &HFF0000) Xor (Te4(s(3)) And &HFF000000) Xor fkey(K + 2)
    t3 = (Te4(s(12)) And &HFF&) Xor (Te4(s(17)) And &HFF00&) Xor (Te4(s(2)) And &HFF0000) Xor (Te4(s(7)) And &HFF000000) Xor fkey(K + 3)
    t4 = (Te4(s(16)) And &HFF&) Xor (Te4(s(1)) And &HFF00&) Xor (Te4(s(6)) And &HFF0000) Xor (Te4(s(11)) And &HFF000000) Xor fkey(K + 4)
    CopyMemory ciphertext(Q + 0), t0, 4&
    CopyMemory ciphertext(Q + 4), t1, 4&
    CopyMemory ciphertext(Q + 8), t2, 4&
    CopyMemory ciphertext(Q + 12), t3, 4&
    CopyMemory ciphertext(Q + 16), t4, 4&
End Sub

'Decrypt a 160 bit block.  ciphertext(q ... q+19) is input, plaintext(p ... p+19) is output.
'plaintext and ciphertext can be the same array.  Will crash if plaintext(p ... p+19) is not allocated.
Private Sub Block160Decrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim s(19) As Byte

    CopyMemory t0, ciphertext(Q + 0), 4&
    CopyMemory t1, ciphertext(Q + 4), 4&
    CopyMemory t2, ciphertext(Q + 8), 4&
    CopyMemory t3, ciphertext(Q + 12), 4&
    CopyMemory t4, ciphertext(Q + 16), 4&
    t0 = t0 Xor rkey(0)
    t1 = t1 Xor rkey(1)
    t2 = t2 Xor rkey(2)
    t3 = t3 Xor rkey(3)
    t4 = t4 Xor rkey(4)
    K = 5

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        t0 = Td0(s(0)) Xor Td1(s(17)) Xor Td2(s(14)) Xor Td3(s(11)) Xor rkey(K + 0)
        t1 = Td0(s(4)) Xor Td1(s(1)) Xor Td2(s(18)) Xor Td3(s(15)) Xor rkey(K + 1)
        t2 = Td0(s(8)) Xor Td1(s(5)) Xor Td2(s(2)) Xor Td3(s(19)) Xor rkey(K + 2)
        t3 = Td0(s(12)) Xor Td1(s(9)) Xor Td2(s(6)) Xor Td3(s(3)) Xor rkey(K + 3)
        t4 = Td0(s(16)) Xor Td1(s(13)) Xor Td2(s(10)) Xor Td3(s(7)) Xor rkey(K + 4)
        K = K + 5
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    t0 = (Td4(s(0)) And &HFF&) Xor (Td4(s(17)) And &HFF00&) Xor (Td4(s(14)) And &HFF0000) Xor (Td4(s(11)) And &HFF000000) Xor rkey(K + 0)
    t1 = (Td4(s(4)) And &HFF&) Xor (Td4(s(1)) And &HFF00&) Xor (Td4(s(18)) And &HFF0000) Xor (Td4(s(15)) And &HFF000000) Xor rkey(K + 1)
    t2 = (Td4(s(8)) And &HFF&) Xor (Td4(s(5)) And &HFF00&) Xor (Td4(s(2)) And &HFF0000) Xor (Td4(s(19)) And &HFF000000) Xor rkey(K + 2)
    t3 = (Td4(s(12)) And &HFF&) Xor (Td4(s(9)) And &HFF00&) Xor (Td4(s(6)) And &HFF0000) Xor (Td4(s(3)) And &HFF000000) Xor rkey(K + 3)
    t4 = (Td4(s(16)) And &HFF&) Xor (Td4(s(13)) And &HFF00&) Xor (Td4(s(10)) And &HFF0000) Xor (Td4(s(7)) And &HFF000000) Xor rkey(K + 4)
    CopyMemory plaintext(p + 0), t0, 4&
    CopyMemory plaintext(p + 4), t1, 4&
    CopyMemory plaintext(p + 8), t2, 4&
    CopyMemory plaintext(p + 12), t3, 4&
    CopyMemory plaintext(p + 16), t4, 4&
End Sub
#End If

'Encrypt a 192 bit block.  plaintext(p ... p+23) is input, ciphertext(q ... q+23) is output.
'plaintext and ciphertext can be the same array.  Will crash if ciphertext(q ... q+23) is not allocated.
Private Sub Block192Encrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim t5    As Long
    Dim s(23) As Byte

    CopyMemory t0, plaintext(p + 0), 4&
    CopyMemory t1, plaintext(p + 4), 4&
    CopyMemory t2, plaintext(p + 8), 4&
    CopyMemory t3, plaintext(p + 12), 4&
    CopyMemory t4, plaintext(p + 16), 4&
    CopyMemory t5, plaintext(p + 20), 4&
    t0 = t0 Xor fkey(0)
    t1 = t1 Xor fkey(1)
    t2 = t2 Xor fkey(2)
    t3 = t3 Xor fkey(3)
    t4 = t4 Xor fkey(4)
    t5 = t5 Xor fkey(5)
    K = 6

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        CopyMemory s(20), t5, 4&
        t0 = Te0(s(0)) Xor Te1(s(5)) Xor Te2(s(10)) Xor Te3(s(15)) Xor fkey(K + 0)
        t1 = Te0(s(4)) Xor Te1(s(9)) Xor Te2(s(14)) Xor Te3(s(19)) Xor fkey(K + 1)
        t2 = Te0(s(8)) Xor Te1(s(13)) Xor Te2(s(18)) Xor Te3(s(23)) Xor fkey(K + 2)
        t3 = Te0(s(12)) Xor Te1(s(17)) Xor Te2(s(22)) Xor Te3(s(3)) Xor fkey(K + 3)
        t4 = Te0(s(16)) Xor Te1(s(21)) Xor Te2(s(2)) Xor Te3(s(7)) Xor fkey(K + 4)
        t5 = Te0(s(20)) Xor Te1(s(1)) Xor Te2(s(6)) Xor Te3(s(11)) Xor fkey(K + 5)
        K = K + 6
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    CopyMemory s(20), t5, 4&
    t0 = (Te4(s(0)) And &HFF&) Xor (Te4(s(5)) And &HFF00&) Xor (Te4(s(10)) And &HFF0000) Xor (Te4(s(15)) And &HFF000000) Xor fkey(K + 0)
    t1 = (Te4(s(4)) And &HFF&) Xor (Te4(s(9)) And &HFF00&) Xor (Te4(s(14)) And &HFF0000) Xor (Te4(s(19)) And &HFF000000) Xor fkey(K + 1)
    t2 = (Te4(s(8)) And &HFF&) Xor (Te4(s(13)) And &HFF00&) Xor (Te4(s(18)) And &HFF0000) Xor (Te4(s(23)) And &HFF000000) Xor fkey(K + 2)
    t3 = (Te4(s(12)) And &HFF&) Xor (Te4(s(17)) And &HFF00&) Xor (Te4(s(22)) And &HFF0000) Xor (Te4(s(3)) And &HFF000000) Xor fkey(K + 3)
    t4 = (Te4(s(16)) And &HFF&) Xor (Te4(s(21)) And &HFF00&) Xor (Te4(s(2)) And &HFF0000) Xor (Te4(s(7)) And &HFF000000) Xor fkey(K + 4)
    t5 = (Te4(s(20)) And &HFF&) Xor (Te4(s(1)) And &HFF00&) Xor (Te4(s(6)) And &HFF0000) Xor (Te4(s(11)) And &HFF000000) Xor fkey(K + 5)

    CopyMemory ciphertext(Q + 0), t0, 4&
    CopyMemory ciphertext(Q + 4), t1, 4&
    CopyMemory ciphertext(Q + 8), t2, 4&
    CopyMemory ciphertext(Q + 12), t3, 4&
    CopyMemory ciphertext(Q + 16), t4, 4&
    CopyMemory ciphertext(Q + 20), t5, 4&
End Sub

'Decrypt a 192 bit block.  ciphertext(q ... q+23) is input, plaintext(p ... p+23) is output.
'plaintext and ciphertext can be the same array.  Will crash if plaintext(p ... p+23) is not allocated.
Private Sub Block192Decrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim t5    As Long
    Dim s(23) As Byte

    CopyMemory t0, ciphertext(Q + 0), 4&
    CopyMemory t1, ciphertext(Q + 4), 4&
    CopyMemory t2, ciphertext(Q + 8), 4&
    CopyMemory t3, ciphertext(Q + 12), 4&
    CopyMemory t4, ciphertext(Q + 16), 4&
    CopyMemory t5, ciphertext(Q + 20), 4&
    t0 = t0 Xor rkey(0)
    t1 = t1 Xor rkey(1)
    t2 = t2 Xor rkey(2)
    t3 = t3 Xor rkey(3)
    t4 = t4 Xor rkey(4)
    t5 = t5 Xor rkey(5)
    K = 6

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        CopyMemory s(20), t5, 4&
        t0 = Td0(s(0)) Xor Td1(s(21)) Xor Td2(s(18)) Xor Td3(s(15)) Xor rkey(K + 0)
        t1 = Td0(s(4)) Xor Td1(s(1)) Xor Td2(s(22)) Xor Td3(s(19)) Xor rkey(K + 1)
        t2 = Td0(s(8)) Xor Td1(s(5)) Xor Td2(s(2)) Xor Td3(s(23)) Xor rkey(K + 2)
        t3 = Td0(s(12)) Xor Td1(s(9)) Xor Td2(s(6)) Xor Td3(s(3)) Xor rkey(K + 3)
        t4 = Td0(s(16)) Xor Td1(s(13)) Xor Td2(s(10)) Xor Td3(s(7)) Xor rkey(K + 4)
        t5 = Td0(s(20)) Xor Td1(s(17)) Xor Td2(s(14)) Xor Td3(s(11)) Xor rkey(K + 5)
        K = K + 6
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    CopyMemory s(20), t5, 4&
    t0 = (Td4(s(0)) And &HFF&) Xor (Td4(s(21)) And &HFF00&) Xor (Td4(s(18)) And &HFF0000) Xor (Td4(s(15)) And &HFF000000) Xor rkey(K + 0)
    t1 = (Td4(s(4)) And &HFF&) Xor (Td4(s(1)) And &HFF00&) Xor (Td4(s(22)) And &HFF0000) Xor (Td4(s(19)) And &HFF000000) Xor rkey(K + 1)
    t2 = (Td4(s(8)) And &HFF&) Xor (Td4(s(5)) And &HFF00&) Xor (Td4(s(2)) And &HFF0000) Xor (Td4(s(23)) And &HFF000000) Xor rkey(K + 2)
    t3 = (Td4(s(12)) And &HFF&) Xor (Td4(s(9)) And &HFF00&) Xor (Td4(s(6)) And &HFF0000) Xor (Td4(s(3)) And &HFF000000) Xor rkey(K + 3)
    t4 = (Td4(s(16)) And &HFF&) Xor (Td4(s(13)) And &HFF00&) Xor (Td4(s(10)) And &HFF0000) Xor (Td4(s(7)) And &HFF000000) Xor rkey(K + 4)
    t5 = (Td4(s(20)) And &HFF&) Xor (Td4(s(17)) And &HFF00&) Xor (Td4(s(14)) And &HFF0000) Xor (Td4(s(11)) And &HFF000000) Xor rkey(K + 5)
    CopyMemory plaintext(p + 0), t0, 4&
    CopyMemory plaintext(p + 4), t1, 4&
    CopyMemory plaintext(p + 8), t2, 4&
    CopyMemory plaintext(p + 12), t3, 4&
    CopyMemory plaintext(p + 16), t4, 4&
    CopyMemory plaintext(p + 20), t5, 4&
End Sub

#If SUPPORT_LEVEL = 2 Then
'Encrypt a 224 bit block.  plaintext(p ... p+27) is input, ciphertext(q ... q+27) is output.
'plaintext and ciphertext can be the same array.  Will crash if ciphertext(q ... q+27) is not allocated.
Private Sub Block224Encrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim t5    As Long
    Dim t6    As Long
    Dim s(27) As Byte

    CopyMemory t0, plaintext(p + 0), 4&
    CopyMemory t1, plaintext(p + 4), 4&
    CopyMemory t2, plaintext(p + 8), 4&
    CopyMemory t3, plaintext(p + 12), 4&
    CopyMemory t4, plaintext(p + 16), 4&
    CopyMemory t5, plaintext(p + 20), 4&
    CopyMemory t6, plaintext(p + 24), 4&
    t0 = t0 Xor fkey(0)
    t1 = t1 Xor fkey(1)
    t2 = t2 Xor fkey(2)
    t3 = t3 Xor fkey(3)
    t4 = t4 Xor fkey(4)
    t5 = t5 Xor fkey(5)
    t6 = t6 Xor fkey(6)
    K = 7

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        CopyMemory s(20), t5, 4&
        CopyMemory s(24), t6, 4&
        t0 = Te0(s(0)) Xor Te1(s(5)) Xor Te2(s(10)) Xor Te3(s(19)) Xor fkey(K + 0)
        t1 = Te0(s(4)) Xor Te1(s(9)) Xor Te2(s(14)) Xor Te3(s(23)) Xor fkey(K + 1)
        t2 = Te0(s(8)) Xor Te1(s(13)) Xor Te2(s(18)) Xor Te3(s(27)) Xor fkey(K + 2)
        t3 = Te0(s(12)) Xor Te1(s(17)) Xor Te2(s(22)) Xor Te3(s(3)) Xor fkey(K + 3)
        t4 = Te0(s(16)) Xor Te1(s(21)) Xor Te2(s(26)) Xor Te3(s(7)) Xor fkey(K + 4)
        t5 = Te0(s(20)) Xor Te1(s(25)) Xor Te2(s(2)) Xor Te3(s(11)) Xor fkey(K + 5)
        t6 = Te0(s(24)) Xor Te1(s(1)) Xor Te2(s(6)) Xor Te3(s(15)) Xor fkey(K + 6)
        K = K + 7
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    CopyMemory s(20), t5, 4&
    CopyMemory s(24), t6, 4&
    t0 = (Te4(s(0)) And &HFF&) Xor (Te4(s(5)) And &HFF00&) Xor (Te4(s(10)) And &HFF0000) Xor (Te4(s(19)) And &HFF000000) Xor fkey(K + 0)
    t1 = (Te4(s(4)) And &HFF&) Xor (Te4(s(9)) And &HFF00&) Xor (Te4(s(14)) And &HFF0000) Xor (Te4(s(23)) And &HFF000000) Xor fkey(K + 1)
    t2 = (Te4(s(8)) And &HFF&) Xor (Te4(s(13)) And &HFF00&) Xor (Te4(s(18)) And &HFF0000) Xor (Te4(s(27)) And &HFF000000) Xor fkey(K + 2)
    t3 = (Te4(s(12)) And &HFF&) Xor (Te4(s(17)) And &HFF00&) Xor (Te4(s(22)) And &HFF0000) Xor (Te4(s(3)) And &HFF000000) Xor fkey(K + 3)
    t4 = (Te4(s(16)) And &HFF&) Xor (Te4(s(21)) And &HFF00&) Xor (Te4(s(26)) And &HFF0000) Xor (Te4(s(7)) And &HFF000000) Xor fkey(K + 4)
    t5 = (Te4(s(20)) And &HFF&) Xor (Te4(s(25)) And &HFF00&) Xor (Te4(s(2)) And &HFF0000) Xor (Te4(s(11)) And &HFF000000) Xor fkey(K + 5)
    t6 = (Te4(s(24)) And &HFF&) Xor (Te4(s(1)) And &HFF00&) Xor (Te4(s(6)) And &HFF0000) Xor (Te4(s(15)) And &HFF000000) Xor fkey(K + 6)
    CopyMemory ciphertext(Q + 0), t0, 4&
    CopyMemory ciphertext(Q + 4), t1, 4&
    CopyMemory ciphertext(Q + 8), t2, 4&
    CopyMemory ciphertext(Q + 12), t3, 4&
    CopyMemory ciphertext(Q + 16), t4, 4&
    CopyMemory ciphertext(Q + 20), t5, 4&
    CopyMemory ciphertext(Q + 24), t6, 4&
End Sub

'Decrypt a 224 bit block.  ciphertext(q ... q+27) is input, plaintext(p ... p+27) is output.
'plaintext and ciphertext can be the same array.  Will crash if plaintext(p ... p+27) is not allocated.
Private Sub Block224Decrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim t5    As Long
    Dim t6    As Long
    Dim s(27) As Byte

    CopyMemory t0, ciphertext(Q + 0), 4&
    CopyMemory t1, ciphertext(Q + 4), 4&
    CopyMemory t2, ciphertext(Q + 8), 4&
    CopyMemory t3, ciphertext(Q + 12), 4&
    CopyMemory t4, ciphertext(Q + 16), 4&
    CopyMemory t5, ciphertext(Q + 20), 4&
    CopyMemory t6, ciphertext(Q + 24), 4&
    t0 = t0 Xor rkey(0)
    t1 = t1 Xor rkey(1)
    t2 = t2 Xor rkey(2)
    t3 = t3 Xor rkey(3)
    t4 = t4 Xor rkey(4)
    t5 = t5 Xor rkey(5)
    t6 = t6 Xor rkey(6)
    K = 7

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        CopyMemory s(20), t5, 4&
        CopyMemory s(24), t6, 4&
        t0 = Td0(s(0)) Xor Td1(s(25)) Xor Td2(s(22)) Xor Td3(s(15)) Xor rkey(K + 0)
        t1 = Td0(s(4)) Xor Td1(s(1)) Xor Td2(s(26)) Xor Td3(s(19)) Xor rkey(K + 1)
        t2 = Td0(s(8)) Xor Td1(s(5)) Xor Td2(s(2)) Xor Td3(s(23)) Xor rkey(K + 2)
        t3 = Td0(s(12)) Xor Td1(s(9)) Xor Td2(s(6)) Xor Td3(s(27)) Xor rkey(K + 3)
        t4 = Td0(s(16)) Xor Td1(s(13)) Xor Td2(s(10)) Xor Td3(s(3)) Xor rkey(K + 4)
        t5 = Td0(s(20)) Xor Td1(s(17)) Xor Td2(s(14)) Xor Td3(s(7)) Xor rkey(K + 5)
        t6 = Td0(s(24)) Xor Td1(s(21)) Xor Td2(s(18)) Xor Td3(s(11)) Xor rkey(K + 6)
        K = K + 7
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    CopyMemory s(20), t5, 4&
    CopyMemory s(24), t6, 4&
    t0 = (Td4(s(0)) And &HFF&) Xor (Td4(s(25)) And &HFF00&) Xor (Td4(s(22)) And &HFF0000) Xor (Td4(s(15)) And &HFF000000) Xor rkey(K + 0)
    t1 = (Td4(s(4)) And &HFF&) Xor (Td4(s(1)) And &HFF00&) Xor (Td4(s(26)) And &HFF0000) Xor (Td4(s(19)) And &HFF000000) Xor rkey(K + 1)
    t2 = (Td4(s(8)) And &HFF&) Xor (Td4(s(5)) And &HFF00&) Xor (Td4(s(2)) And &HFF0000) Xor (Td4(s(23)) And &HFF000000) Xor rkey(K + 2)
    t3 = (Td4(s(12)) And &HFF&) Xor (Td4(s(9)) And &HFF00&) Xor (Td4(s(6)) And &HFF0000) Xor (Td4(s(27)) And &HFF000000) Xor rkey(K + 3)
    t4 = (Td4(s(16)) And &HFF&) Xor (Td4(s(13)) And &HFF00&) Xor (Td4(s(10)) And &HFF0000) Xor (Td4(s(3)) And &HFF000000) Xor rkey(K + 4)
    t5 = (Td4(s(20)) And &HFF&) Xor (Td4(s(17)) And &HFF00&) Xor (Td4(s(14)) And &HFF0000) Xor (Td4(s(7)) And &HFF000000) Xor rkey(K + 5)
    t6 = (Td4(s(24)) And &HFF&) Xor (Td4(s(21)) And &HFF00&) Xor (Td4(s(18)) And &HFF0000) Xor (Td4(s(11)) And &HFF000000) Xor rkey(K + 6)
    CopyMemory plaintext(p + 0), t0, 4&
    CopyMemory plaintext(p + 4), t1, 4&
    CopyMemory plaintext(p + 8), t2, 4&
    CopyMemory plaintext(p + 12), t3, 4&
    CopyMemory plaintext(p + 16), t4, 4&
    CopyMemory plaintext(p + 20), t5, 4&
    CopyMemory plaintext(p + 24), t6, 4&
End Sub
#End If

'Encrypt a 256 bit block.  plaintext(p ... p+31) is input, ciphertext(q ... q+31) is output.
'plaintext and ciphertext can be the same array.  Will crash if ciphertext(q ... q+31) is not allocated.
Private Sub Block256Encrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim t5    As Long
    Dim t6    As Long
    Dim t7    As Long
    Dim s(31) As Byte

    CopyMemory t0, plaintext(p + 0), 4&
    CopyMemory t1, plaintext(p + 4), 4&
    CopyMemory t2, plaintext(p + 8), 4&
    CopyMemory t3, plaintext(p + 12), 4&
    CopyMemory t4, plaintext(p + 16), 4&
    CopyMemory t5, plaintext(p + 20), 4&
    CopyMemory t6, plaintext(p + 24), 4&
    CopyMemory t7, plaintext(p + 28), 4&
    t0 = t0 Xor fkey(0)
    t1 = t1 Xor fkey(1)
    t2 = t2 Xor fkey(2)
    t3 = t3 Xor fkey(3)
    t4 = t4 Xor fkey(4)
    t5 = t5 Xor fkey(5)
    t6 = t6 Xor fkey(6)
    t7 = t7 Xor fkey(7)
    K = 8

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        CopyMemory s(20), t5, 4&
        CopyMemory s(24), t6, 4&
        CopyMemory s(28), t7, 4&
        t0 = Te0(s(0)) Xor Te1(s(5)) Xor Te2(s(14)) Xor Te3(s(19)) Xor fkey(K + 0)
        t1 = Te0(s(4)) Xor Te1(s(9)) Xor Te2(s(18)) Xor Te3(s(23)) Xor fkey(K + 1)
        t2 = Te0(s(8)) Xor Te1(s(13)) Xor Te2(s(22)) Xor Te3(s(27)) Xor fkey(K + 2)
        t3 = Te0(s(12)) Xor Te1(s(17)) Xor Te2(s(26)) Xor Te3(s(31)) Xor fkey(K + 3)
        t4 = Te0(s(16)) Xor Te1(s(21)) Xor Te2(s(30)) Xor Te3(s(3)) Xor fkey(K + 4)
        t5 = Te0(s(20)) Xor Te1(s(25)) Xor Te2(s(2)) Xor Te3(s(7)) Xor fkey(K + 5)
        t6 = Te0(s(24)) Xor Te1(s(29)) Xor Te2(s(6)) Xor Te3(s(11)) Xor fkey(K + 6)
        t7 = Te0(s(28)) Xor Te1(s(1)) Xor Te2(s(10)) Xor Te3(s(15)) Xor fkey(K + 7)
        K = K + 8
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    CopyMemory s(20), t5, 4&
    CopyMemory s(24), t6, 4&
    CopyMemory s(28), t7, 4&
    t0 = (Te4(s(0)) And &HFF&) Xor (Te4(s(5)) And &HFF00&) Xor (Te4(s(14)) And &HFF0000) Xor (Te4(s(19)) And &HFF000000) Xor fkey(K + 0)
    t1 = (Te4(s(4)) And &HFF&) Xor (Te4(s(9)) And &HFF00&) Xor (Te4(s(18)) And &HFF0000) Xor (Te4(s(23)) And &HFF000000) Xor fkey(K + 1)
    t2 = (Te4(s(8)) And &HFF&) Xor (Te4(s(13)) And &HFF00&) Xor (Te4(s(22)) And &HFF0000) Xor (Te4(s(27)) And &HFF000000) Xor fkey(K + 2)
    t3 = (Te4(s(12)) And &HFF&) Xor (Te4(s(17)) And &HFF00&) Xor (Te4(s(26)) And &HFF0000) Xor (Te4(s(31)) And &HFF000000) Xor fkey(K + 3)
    t4 = (Te4(s(16)) And &HFF&) Xor (Te4(s(21)) And &HFF00&) Xor (Te4(s(30)) And &HFF0000) Xor (Te4(s(3)) And &HFF000000) Xor fkey(K + 4)
    t5 = (Te4(s(20)) And &HFF&) Xor (Te4(s(25)) And &HFF00&) Xor (Te4(s(2)) And &HFF0000) Xor (Te4(s(7)) And &HFF000000) Xor fkey(K + 5)
    t6 = (Te4(s(24)) And &HFF&) Xor (Te4(s(29)) And &HFF00&) Xor (Te4(s(6)) And &HFF0000) Xor (Te4(s(11)) And &HFF000000) Xor fkey(K + 6)
    t7 = (Te4(s(28)) And &HFF&) Xor (Te4(s(1)) And &HFF00&) Xor (Te4(s(10)) And &HFF0000) Xor (Te4(s(15)) And &HFF000000) Xor fkey(K + 7)
    CopyMemory ciphertext(Q + 0), t0, 4&
    CopyMemory ciphertext(Q + 4), t1, 4&
    CopyMemory ciphertext(Q + 8), t2, 4&
    CopyMemory ciphertext(Q + 12), t3, 4&
    CopyMemory ciphertext(Q + 16), t4, 4&
    CopyMemory ciphertext(Q + 20), t5, 4&
    CopyMemory ciphertext(Q + 24), t6, 4&
    CopyMemory ciphertext(Q + 28), t7, 4&
End Sub

'Decrypt a 256 bit block.  ciphertext(q ... q+31) is input, plaintext(p ... p+31) is output.
'plaintext and ciphertext can be the same array.  Will crash if plaintext(p ... p+31) is not allocated.
Private Sub Block256Decrypt(plaintext() As Byte, ciphertext() As Byte, p As Long, Q As Long)
    Dim i     As Long
    Dim K     As Long
    Dim t0    As Long
    Dim t1    As Long
    Dim t2    As Long
    Dim t3    As Long
    Dim t4    As Long
    Dim t5    As Long
    Dim t6    As Long
    Dim t7    As Long
    Dim s(31) As Byte

    CopyMemory t0, ciphertext(Q + 0), 4&
    CopyMemory t1, ciphertext(Q + 4), 4&
    CopyMemory t2, ciphertext(Q + 8), 4&
    CopyMemory t3, ciphertext(Q + 12), 4&
    CopyMemory t4, ciphertext(Q + 16), 4&
    CopyMemory t5, ciphertext(Q + 20), 4&
    CopyMemory t6, ciphertext(Q + 24), 4&
    CopyMemory t7, ciphertext(Q + 28), 4&
    t0 = t0 Xor rkey(0)
    t1 = t1 Xor rkey(1)
    t2 = t2 Xor rkey(2)
    t3 = t3 Xor rkey(3)
    t4 = t4 Xor rkey(4)
    t5 = t5 Xor rkey(5)
    t6 = t6 Xor rkey(6)
    t7 = t7 Xor rkey(7)
    K = 8

    For i = 1 To Nr - 1 'Nr is number of rounds
        CopyMemory s(0), t0, 4&
        CopyMemory s(4), t1, 4&
        CopyMemory s(8), t2, 4&
        CopyMemory s(12), t3, 4&
        CopyMemory s(16), t4, 4&
        CopyMemory s(20), t5, 4&
        CopyMemory s(24), t6, 4&
        CopyMemory s(28), t7, 4&
        t0 = Td0(s(0)) Xor Td1(s(29)) Xor Td2(s(22)) Xor Td3(s(19)) Xor rkey(K + 0)
        t1 = Td0(s(4)) Xor Td1(s(1)) Xor Td2(s(26)) Xor Td3(s(23)) Xor rkey(K + 1)
        t2 = Td0(s(8)) Xor Td1(s(5)) Xor Td2(s(30)) Xor Td3(s(27)) Xor rkey(K + 2)
        t3 = Td0(s(12)) Xor Td1(s(9)) Xor Td2(s(2)) Xor Td3(s(31)) Xor rkey(K + 3)
        t4 = Td0(s(16)) Xor Td1(s(13)) Xor Td2(s(6)) Xor Td3(s(3)) Xor rkey(K + 4)
        t5 = Td0(s(20)) Xor Td1(s(17)) Xor Td2(s(10)) Xor Td3(s(7)) Xor rkey(K + 5)
        t6 = Td0(s(24)) Xor Td1(s(21)) Xor Td2(s(14)) Xor Td3(s(11)) Xor rkey(K + 6)
        t7 = Td0(s(28)) Xor Td1(s(25)) Xor Td2(s(18)) Xor Td3(s(15)) Xor rkey(K + 7)
        K = K + 8
    Next i

    'Final round
    CopyMemory s(0), t0, 4&
    CopyMemory s(4), t1, 4&
    CopyMemory s(8), t2, 4&
    CopyMemory s(12), t3, 4&
    CopyMemory s(16), t4, 4&
    CopyMemory s(20), t5, 4&
    CopyMemory s(24), t6, 4&
    CopyMemory s(28), t7, 4&
    t0 = (Td4(s(0)) And &HFF&) Xor (Td4(s(29)) And &HFF00&) Xor (Td4(s(22)) And &HFF0000) Xor (Td4(s(19)) And &HFF000000) Xor rkey(K + 0)
    t1 = (Td4(s(4)) And &HFF&) Xor (Td4(s(1)) And &HFF00&) Xor (Td4(s(26)) And &HFF0000) Xor (Td4(s(23)) And &HFF000000) Xor rkey(K + 1)
    t2 = (Td4(s(8)) And &HFF&) Xor (Td4(s(5)) And &HFF00&) Xor (Td4(s(30)) And &HFF0000) Xor (Td4(s(27)) And &HFF000000) Xor rkey(K + 2)
    t3 = (Td4(s(12)) And &HFF&) Xor (Td4(s(9)) And &HFF00&) Xor (Td4(s(2)) And &HFF0000) Xor (Td4(s(31)) And &HFF000000) Xor rkey(K + 3)
    t4 = (Td4(s(16)) And &HFF&) Xor (Td4(s(13)) And &HFF00&) Xor (Td4(s(6)) And &HFF0000) Xor (Td4(s(3)) And &HFF000000) Xor rkey(K + 4)
    t5 = (Td4(s(20)) And &HFF&) Xor (Td4(s(17)) And &HFF00&) Xor (Td4(s(10)) And &HFF0000) Xor (Td4(s(7)) And &HFF000000) Xor rkey(K + 5)
    t6 = (Td4(s(24)) And &HFF&) Xor (Td4(s(21)) And &HFF00&) Xor (Td4(s(14)) And &HFF0000) Xor (Td4(s(11)) And &HFF000000) Xor rkey(K + 6)
    t7 = (Td4(s(28)) And &HFF&) Xor (Td4(s(25)) And &HFF00&) Xor (Td4(s(18)) And &HFF0000) Xor (Td4(s(15)) And &HFF000000) Xor rkey(K + 7)
    CopyMemory plaintext(p + 0), t0, 4&
    CopyMemory plaintext(p + 4), t1, 4&
    CopyMemory plaintext(p + 8), t2, 4&
    CopyMemory plaintext(p + 12), t3, 4&
    CopyMemory plaintext(p + 16), t4, 4&
    CopyMemory plaintext(p + 20), t5, 4&
    CopyMemory plaintext(p + 24), t6, 4&
    CopyMemory plaintext(p + 28), t7, 4&
End Sub
#End If


'Encrypt an arbritrary size array in blocks.  plaintext(0 ... n-1) is input,
'ciphertext(0 ... m-1) is output, where m is n padded to a multiple of (BlockBits\8) bytes.
'plaintext and ciphertext must be distinct, and ciphertext must be a redimensionable array.
'If appendsize is non-zero, a one byte length field is appended to the plaintext
'before encrypting so the original length can be retrieved after decryption.
#If SUPPORT_LEVEL Then
Private Function ArrayEncrypt(plaintext() As Byte, ciphertext() As Byte, appendsize As Long, BlockBits As Long) As Long
#Else
Private Function ArrayEncrypt(plaintext() As Byte, ciphertext() As Byte, appendsize As Long) As Long
#End If
    Dim i           As Long
    Dim m           As Long
    Dim n           As Long
#If SUPPORT_LEVEL = 0 Then
    Const BlockSize As Long = 16 'bytes
#Else
    Dim BlockSize   As Long

    Select Case BlockBits
    Case 128:  BlockSize = 16
    Case 192:  BlockSize = 24
    Case 256:  BlockSize = 32
#If SUPPORT_LEVEL = 2 Then
    Case 160:  BlockSize = 20
    Case 224:  BlockSize = 28
#End If
    Case Else: ERR.Raise 1, , "cRijndael.ArrayEncrypt - Illegal BlockBits value"
    End Select
#End If

    If LBound(plaintext) <> 0 Then ERR.Raise 1, , "cRijndael.ArrayEncrypt - plaintext must be zero based array"

    n = UBound(plaintext) + 1
    If appendsize = 0 Then
#If SUPPORT_LEVEL Then
        m = ((n + BlockSize - 1) \ BlockSize) * BlockSize
#Else
        m = (n + BlockSize - 1) And &HFFFFFFF0  'BlockSize=16 specific
#End If
        ReDim ciphertext(m - 1)
    Else
#If SUPPORT_LEVEL Then
        m = ((n + BlockSize) \ BlockSize) * BlockSize
#Else
        m = (n + BlockSize) And &HFFFFFFF0      'BlockSize=16 specific
#End If
        ReDim ciphertext(m - 1)
        ciphertext(m - 1) = n Mod BlockSize
    End If

#If SUPPORT_LEVEL Then
    Select Case BlockBits
    Case 128
#End If

    For i = 0 To n - BlockSize Step BlockSize
        BlockEncrypt plaintext, ciphertext, i, i
    Next i

#If SUPPORT_LEVEL Then
    Case 192
    For i = 0 To n - BlockSize Step BlockSize
        Block192Encrypt plaintext, ciphertext, i, i
    Next i
    Case 256
    For i = 0 To n - BlockSize Step BlockSize
        Block256Encrypt plaintext, ciphertext, i, i
    Next i
#If SUPPORT_LEVEL = 2 Then
    Case 160
    For i = 0 To n - BlockSize Step BlockSize
        Block160Encrypt plaintext, ciphertext, i, i
    Next i
    Case 224
    For i = 0 To n - BlockSize Step BlockSize
        Block224Encrypt plaintext, ciphertext, i, i
    Next i
#End If
    End Select
#End If

    If (n Mod BlockSize) <> 0 Then CopyMemory ciphertext(i), plaintext(i), n Mod BlockSize

#If SUPPORT_LEVEL Then
    Select Case BlockBits
    Case 128
#End If

    If i <> m Then BlockEncrypt ciphertext, ciphertext, i, i

#If SUPPORT_LEVEL Then
    Case 192
    If i <> m Then Block192Encrypt ciphertext, ciphertext, i, i
    Case 256
    If i <> m Then Block256Encrypt ciphertext, ciphertext, i, i
#If SUPPORT_LEVEL = 2 Then
    Case 160
    If i <> m Then Block160Encrypt ciphertext, ciphertext, i, i
    Case 224
    If i <> m Then Block224Encrypt ciphertext, ciphertext, i, i
#End If
    End Select
#End If

End Function

'Decrypts an array encrypted with the ArrayEncrypt function
#If SUPPORT_LEVEL Then
Private Function ArrayDecrypt(plaintext() As Byte, ciphertext() As Byte, appendsize As Long, BlockBits As Long) As Long
#Else
Private Function ArrayDecrypt(plaintext() As Byte, ciphertext() As Byte, appendsize As Long) As Long
#End If
    Dim i           As Long
    Dim m           As Long
    Dim n           As Long
#If SUPPORT_LEVEL = 0 Then
    Const BlockSize As Long = 16 'bytes
#Else
    Dim BlockSize   As Long

    Select Case BlockBits
    Case 128:  BlockSize = 16
    Case 192:  BlockSize = 24
    Case 256:  BlockSize = 32
#If SUPPORT_LEVEL = 2 Then
    Case 160:  BlockSize = 20
    Case 224:  BlockSize = 28
#End If
    Case Else: ERR.Raise 1, , "cRijndael.ArrayDecrypt - Illegal BlockBits value"
    End Select
#End If

    If LBound(ciphertext) <> 0 Then ERR.Raise 1, , "cRijndael.ArrayDecrypt - ciphertext must be zero based array"

    n = UBound(ciphertext) + 1
    If ((n Mod BlockSize) = 0) Then
        ReDim plaintext(n - 1)

#If SUPPORT_LEVEL Then
        Select Case BlockBits
        Case 128
#End If

        For i = 0 To n - BlockSize Step BlockSize
            BlockDecrypt plaintext, ciphertext, i, i
        Next i

#If SUPPORT_LEVEL Then
        Case 192
        For i = 0 To n - BlockSize Step BlockSize
            Block192Decrypt plaintext, ciphertext, i, i
        Next i
        Case 256
        For i = 0 To n - BlockSize Step BlockSize
            Block256Decrypt plaintext, ciphertext, i, i
        Next i
#If SUPPORT_LEVEL = 2 Then
        Case 160
        For i = 0 To n - BlockSize Step BlockSize
            Block160Decrypt plaintext, ciphertext, i, i
        Next i
        Case 224
        For i = 0 To n - BlockSize Step BlockSize
            Block224Decrypt plaintext, ciphertext, i, i
        Next i
#End If
        End Select
#End If

        If appendsize Then
            If plaintext(n - 1) < BlockSize Then
                n = n - BlockSize + plaintext(n - 1)
                If n > 0 Then ReDim Preserve plaintext(n - 1)
            Else
                Call SHOWWRONG("warning - incorrect length field", 0)
                ArrayDecrypt = 1
            End If
        End If
    Else
         Call SHOWWRONG("ciphertext size not a multiple of block size", 0)
        ArrayDecrypt = -1
    End If

End Function




'File encryption.  The ciphertext file will be padded to a multiple of (BlockBits\8) bytes.
'One byte is appended to the ciphertext to indicate the length of the final block.
'FileDecrypt recovers this length field and restores the original plaintext file length.
'PlaintextFileName and CiphertextFileName must be distinct, PlaintextFileName must exist,
'CiphertextFileName must not exist, etc.  Be sure to do file checking before calling this.
#If SUPPORT_LEVEL Then
Private Function FileEncrypt(PlaintextFileName As String, CiphertextFileName As String, BlockBits As Long) As Long
#Else
Private Function FileEncrypt(PlaintextFileName As String, CiphertextFileName As String) As Long
#End If
    Dim FileNum     As Integer
    Dim FileNum2    As Integer
    Dim i           As Long
    Dim m           As Long 'ciphertext file size
    Dim n           As Long 'plaintext file size
    Dim Data()      As Byte
#If SUPPORT_LEVEL = 0 Then
    Const BlockSize As Long = 16 'bytes
    Const MaxBlocks As Long = MaxFileChunkSize \ BlockSize
#Else
    Dim BlockSize   As Long
    Dim MaxBlocks   As Long

    Select Case BlockBits
    Case 128:  BlockSize = 16
    Case 192:  BlockSize = 24
    Case 256:  BlockSize = 32
#If SUPPORT_LEVEL = 2 Then
    Case 160:  BlockSize = 20
    Case 224:  BlockSize = 28
#End If
    Case Else: ERR.Raise 1, , "cRijndael.FileEncrypt - Illegal BlockBits value"
    End Select
    MaxBlocks = MaxFileChunkSize \ BlockSize
#End If

    n = FileLen(PlaintextFileName)
#If SUPPORT_LEVEL Then
    m = ((n + BlockSize) \ BlockSize) * BlockSize
#Else
    m = (n + BlockSize) And (-BlockSize) 'BlockSize=16 specific
#End If

    FileNum = FreeFile
    Open PlaintextFileName For Binary Access Read As FileNum
    FileNum2 = FreeFile
    Open CiphertextFileName For Binary Access Write As FileNum2

    'For large files, encrypt in pieces no larger than MaxFileChunkSize
    If m > MaxBlocks * BlockSize Then
        ReDim Data(MaxBlocks * BlockSize - 1)
        Do
            Get #FileNum, , Data
#If SUPPORT_LEVEL Then
            Select Case BlockBits
            Case 128
#End If
            For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                BlockEncrypt Data, Data, i, i
            Next i
#If SUPPORT_LEVEL Then
            Case 192
            For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                Block192Encrypt Data, Data, i, i
            Next i
            Case 256
            For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                Block256Encrypt Data, Data, i, i
            Next i
#If SUPPORT_LEVEL = 2 Then
            Case 160
            For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                Block160Encrypt Data, Data, i, i
            Next i
            Case 224
            For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                Block224Encrypt Data, Data, i, i
            Next i
#End If
            End Select
#End If
            Put #FileNum2, , Data
            m = m - MaxBlocks * BlockSize
        Loop While m > MaxBlocks * BlockSize
    End If

    'Encrypt the last piece of the file
    ReDim Data(m - 1)
    Get #FileNum, , Data
    Data(m - 1) = n Mod BlockSize
#If SUPPORT_LEVEL Then
    Select Case BlockBits
    Case 128
#End If
    For i = 0 To m - BlockSize Step BlockSize
        BlockEncrypt Data, Data, i, i
    Next i
#If SUPPORT_LEVEL Then
    Case 192
    For i = 0 To m - BlockSize Step BlockSize
        Block192Encrypt Data, Data, i, i
    Next i
    Case 256
    For i = 0 To m - BlockSize Step BlockSize
        Block256Encrypt Data, Data, i, i
    Next i
#If SUPPORT_LEVEL = 2 Then
    Case 160
    For i = 0 To m - BlockSize Step BlockSize
        Block160Encrypt Data, Data, i, i
    Next i
    Case 224
    For i = 0 To m - BlockSize Step BlockSize
        Block224Encrypt Data, Data, i, i
    Next i
#End If
    End Select
#End If
    Put FileNum2, , Data

    Close FileNum
    Close FileNum2
End Function

'File decryption.
#If SUPPORT_LEVEL Then
Private Function FileDecrypt(PlaintextFileName As String, CiphertextFileName As String, BlockBits As Long) As Long
#Else
Private Function FileDecrypt(PlaintextFileName As String, CiphertextFileName As String) As Long
#End If
    Dim FileNum     As Integer
    Dim FileNum2    As Integer
    Dim i           As Long
    Dim m           As Long 'ciphertext file size
    Dim n           As Long 'plaintext file size
    Dim Data()      As Byte
#If SUPPORT_LEVEL = 0 Then
    Const BlockSize As Long = 16 'bytes
    Const MaxBlocks As Long = MaxFileChunkSize \ BlockSize
#Else
    Dim BlockSize   As Long
    Dim MaxBlocks   As Long

    Select Case BlockBits
    Case 128:  BlockSize = 16
    Case 192:  BlockSize = 24
    Case 256:  BlockSize = 32
#If SUPPORT_LEVEL = 2 Then
    Case 160:  BlockSize = 20
    Case 224:  BlockSize = 28
#End If
    Case Else: ERR.Raise 1, , "cRijndael.FileDecrypt - Illegal BlockBits value"
    End Select
    MaxBlocks = MaxFileChunkSize \ BlockSize
#End If

    m = FileLen(CiphertextFileName)
#If SUPPORT_LEVEL Then
    If (m = 0) Or ((m Mod BlockSize) <> 0) Then
#Else
    If (m = 0) Or ((m And (BlockSize - 1)) <> 0) Then 'BlockSize=16 specific
#End If
        Call SHOWWRONG("File Size Error - ciphertext file not a multiple of block size", 2)
        FileDecrypt = 1
    Else
        FileNum = FreeFile
        Open CiphertextFileName For Binary Access Read As FileNum
        FileNum2 = FreeFile
        Open PlaintextFileName For Binary Access Write As FileNum2

        'For large files, decrypt in pieces no larger than MaxFileChunkSize
        If m > MaxBlocks * BlockSize Then
            ReDim Data(MaxBlocks * BlockSize - 1)
            Do
                Get #FileNum, , Data
#If SUPPORT_LEVEL Then
                Select Case BlockBits
                Case 128
#End If
                For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                    BlockDecrypt Data, Data, i, i
                Next i
#If SUPPORT_LEVEL Then
                Case 192
                For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                    Block192Decrypt Data, Data, i, i
                Next i
                Case 256
                For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                    Block256Decrypt Data, Data, i, i
                Next i
#If SUPPORT_LEVEL = 2 Then
                Case 160
                For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                    Block160Decrypt Data, Data, i, i
                Next i
                Case 224
                For i = 0 To (MaxBlocks - 1) * BlockSize Step BlockSize
                    Block224Decrypt Data, Data, i, i
                Next i
#End If
                End Select
#End If
                Put #FileNum2, , Data
                m = m - MaxBlocks * BlockSize
            Loop While m > MaxBlocks * BlockSize
        End If

        'Decrypt the last piece of the file
        ReDim Data(m - 1)
        Get #FileNum, , Data
#If SUPPORT_LEVEL Then
        Select Case BlockBits
        Case 128
#End If
        For i = 0 To m - BlockSize Step BlockSize
            BlockDecrypt Data, Data, i, i
        Next i
#If SUPPORT_LEVEL Then
        Case 192
        For i = 0 To m - BlockSize Step BlockSize
            Block192Decrypt Data, Data, i, i
        Next i
        Case 256
        For i = 0 To m - BlockSize Step BlockSize
            Block256Decrypt Data, Data, i, i
        Next i
#If SUPPORT_LEVEL = 2 Then
        Case 160
        For i = 0 To m - BlockSize Step BlockSize
            Block160Decrypt Data, Data, i, i
        Next i
        Case 224
        For i = 0 To m - BlockSize Step BlockSize
            Block224Decrypt Data, Data, i, i
        Next i
#End If
        End Select
#End If
        'Recover length field
        If Data(m - 1) < BlockSize Then
            n = m - BlockSize + CLng(Data(m - 1))
        Else
#If SUPPORT_LEVEL Then
            Call SHOWWRONG("warning - incorrect length field in decrypted file." & vbCrLf & "Wrong key, keysize, or blocksize.", 2)
#Else
           Call SHOWWRONG("warning - incorrect length field in decrypted file." & vbCrLf & "Wrong key or keysize.", 2)
#End If
            n = m
        End If
        If n > 0 Then
            ReDim Preserve Data(n - 1)
            Put FileNum2, , Data
        End If

        Close FileNum
        Close FileNum2
    End If
End Function


Private Sub Class_Initialize()
#If COMPILE_CONSTANTS = 0 Then
    Dim i          As Long
    Dim Y          As Byte
    Dim s(7)       As Byte
    Dim IB         As Byte
    Dim ptab(255)  As Byte
    Dim ltab(255)  As Byte

    'use 3 as primitive root to generate power and log tables
    ltab(0) = 0
    ltab(1) = 0
    ltab(3) = 1
    ptab(0) = 1
    ptab(1) = 3
    For i = 2 To 255 'ptab(i) = ptab(i - 1) Xor Xtime(ptab(i - 1))
        If (ptab(i - 1) And &H80) Then
            ptab(i) = ptab(i - 1) Xor ((ptab(i - 1) And 127) * 2) Xor &H1B
        Else
            ptab(i) = ptab(i - 1) Xor (ptab(i - 1) * 2)
        End If
        ltab(ptab(i)) = i
    Next i

    'affine transformation:- each bit is xored with itself shifted one bit
    Te4(0) = &H63636363
    Td4(&H63) = 0
    For i = 1 To 255
        Y = ptab(255 - ltab(i)) 'multiplicative inverse
        IB = Y
        If IB And &H80 Then IB = (IB And 127) * 2 Or 1 Else IB = IB * 2 'RotateLeftByte
        Y = Y Xor IB
        If IB And &H80 Then IB = (IB And 127) * 2 Or 1 Else IB = IB * 2
        Y = Y Xor IB
        If IB And &H80 Then IB = (IB And 127) * 2 Or 1 Else IB = IB * 2
        Y = Y Xor IB
        If IB And &H80 Then IB = (IB And 127) * 2 Or 1 Else IB = IB * 2
        Y = Y Xor IB Xor &H63

        s(0) = Y
        s(1) = s(0)
        s(2) = s(0)
        s(3) = s(0)
        CopyMemory Te4(i), s(0), 4&

        s(0) = i
        s(1) = s(0)
        s(2) = s(0)
        s(3) = s(0)
        CopyMemory Td4(Y), s(0), 4&
    Next i

    Y = 1
    For i = 0 To UBound(rco)
        rco(i) = Y
        If (Y And &H80) Then 'y = Xtime(y)
            Y = ((Y And 127) * 2) Xor &H1B
        Else
            Y = Y * 2
        End If
    Next i

    'calculate forward and reverse tables
    For i = 0 To 255
        Y = Te4(i) And &HFF&
        's(3) = y Xor Xtime(y)
'        s(0) = Xtime(y)
        If (Y And &H80) Then
            s(0) = ((Y And 127) * 2) Xor &H1B
            s(3) = Y Xor s(0)
        Else
            s(0) = Y * 2
            s(3) = Y Xor s(0)
        End If
        s(2) = Y
        s(1) = Y
        CopyMemory s(4), s(0), 4&
        CopyMemory Te0(i), s(0), 4&
        CopyMemory Te1(i), s(3), 4&
        CopyMemory Te2(i), s(2), 4&
        CopyMemory Te3(i), s(1), 4&
        Y = Td4(i) And &HFF&
        If Y = 0 Then 'x.y= AntiLog(Log(x) + Log(y))
            s(3) = 0
            s(2) = 0
            s(1) = 0
            s(0) = 0
        Else
            s(3) = ptab((CLng(ltab(&HB)) + CLng(ltab(Y))) Mod 255)
            s(2) = ptab((CLng(ltab(&HD)) + CLng(ltab(Y))) Mod 255)
            s(1) = ptab((CLng(ltab(&H9)) + CLng(ltab(Y))) Mod 255)
            s(0) = ptab((CLng(ltab(&HE)) + CLng(ltab(Y))) Mod 255)
        End If
        CopyMemory s(4), s(0), 4&
        CopyMemory Td0(i), s(0), 4&
        CopyMemory Td1(i), s(3), 4&
        CopyMemory Td2(i), s(2), 4&
        CopyMemory Td3(i), s(1), 4&
    Next i
#Else
    Te0(0) = &HA56363C6:   Te0(1) = &H847C7CF8:   Te0(2) = &H997777EE:   Te0(3) = &H8D7B7BF6:   Te0(4) = &HDF2F2FF:    Te0(5) = &HBD6B6BD6:   Te0(6) = &HB16F6FDE:   Te0(7) = &H54C5C591
    Te0(8) = &H50303060:   Te0(9) = &H3010102:    Te0(10) = &HA96767CE:  Te0(11) = &H7D2B2B56:  Te0(12) = &H19FEFEE7:  Te0(13) = &H62D7D7B5:  Te0(14) = &HE6ABAB4D:  Te0(15) = &H9A7676EC
    Te0(16) = &H45CACA8F:  Te0(17) = &H9D82821F:  Te0(18) = &H40C9C989:  Te0(19) = &H877D7DFA:  Te0(20) = &H15FAFAEF:  Te0(21) = &HEB5959B2:  Te0(22) = &HC947478E:  Te0(23) = &HBF0F0FB
    Te0(24) = &HECADAD41:  Te0(25) = &H67D4D4B3:  Te0(26) = &HFDA2A25F:  Te0(27) = &HEAAFAF45:  Te0(28) = &HBF9C9C23:  Te0(29) = &HF7A4A453:  Te0(30) = &H967272E4:  Te0(31) = &H5BC0C09B
    Te0(32) = &HC2B7B775:  Te0(33) = &H1CFDFDE1:  Te0(34) = &HAE93933D:  Te0(35) = &H6A26264C:  Te0(36) = &H5A36366C:  Te0(37) = &H413F3F7E:  Te0(38) = &H2F7F7F5:   Te0(39) = &H4FCCCC83
    Te0(40) = &H5C343468:  Te0(41) = &HF4A5A551:  Te0(42) = &H34E5E5D1:  Te0(43) = &H8F1F1F9:   Te0(44) = &H937171E2:  Te0(45) = &H73D8D8AB:  Te0(46) = &H53313162:  Te0(47) = &H3F15152A
    Te0(48) = &HC040408:   Te0(49) = &H52C7C795:  Te0(50) = &H65232346:  Te0(51) = &H5EC3C39D:  Te0(52) = &H28181830:  Te0(53) = &HA1969637:  Te0(54) = &HF05050A:   Te0(55) = &HB59A9A2F
    Te0(56) = &H907070E:   Te0(57) = &H36121224:  Te0(58) = &H9B80801B:  Te0(59) = &H3DE2E2DF:  Te0(60) = &H26EBEBCD:  Te0(61) = &H6927274E:  Te0(62) = &HCDB2B27F:  Te0(63) = &H9F7575EA
    Te0(64) = &H1B090912:  Te0(65) = &H9E83831D:  Te0(66) = &H742C2C58:  Te0(67) = &H2E1A1A34:  Te0(68) = &H2D1B1B36:  Te0(69) = &HB26E6EDC:  Te0(70) = &HEE5A5AB4:  Te0(71) = &HFBA0A05B
    Te0(72) = &HF65252A4:  Te0(73) = &H4D3B3B76:  Te0(74) = &H61D6D6B7:  Te0(75) = &HCEB3B37D:  Te0(76) = &H7B292952:  Te0(77) = &H3EE3E3DD:  Te0(78) = &H712F2F5E:  Te0(79) = &H97848413
    Te0(80) = &HF55353A6:  Te0(81) = &H68D1D1B9:  Te0(82) = &H0&:        Te0(83) = &H2CEDEDC1:  Te0(84) = &H60202040:  Te0(85) = &H1FFCFCE3:  Te0(86) = &HC8B1B179:  Te0(87) = &HED5B5BB6
    Te0(88) = &HBE6A6AD4:  Te0(89) = &H46CBCB8D:  Te0(90) = &HD9BEBE67:  Te0(91) = &H4B393972:  Te0(92) = &HDE4A4A94:  Te0(93) = &HD44C4C98:  Te0(94) = &HE85858B0:  Te0(95) = &H4ACFCF85
    Te0(96) = &H6BD0D0BB:  Te0(97) = &H2AEFEFC5:  Te0(98) = &HE5AAAA4F:  Te0(99) = &H16FBFBED:  Te0(100) = &HC5434386: Te0(101) = &HD74D4D9A: Te0(102) = &H55333366: Te0(103) = &H94858511
    Te0(104) = &HCF45458A: Te0(105) = &H10F9F9E9: Te0(106) = &H6020204:  Te0(107) = &H817F7FFE: Te0(108) = &HF05050A0: Te0(109) = &H443C3C78: Te0(110) = &HBA9F9F25: Te0(111) = &HE3A8A84B
    Te0(112) = &HF35151A2: Te0(113) = &HFEA3A35D: Te0(114) = &HC0404080: Te0(115) = &H8A8F8F05: Te0(116) = &HAD92923F: Te0(117) = &HBC9D9D21: Te0(118) = &H48383870: Te0(119) = &H4F5F5F1
    Te0(120) = &HDFBCBC63: Te0(121) = &HC1B6B677: Te0(122) = &H75DADAAF: Te0(123) = &H63212142: Te0(124) = &H30101020: Te0(125) = &H1AFFFFE5: Te0(126) = &HEF3F3FD:  Te0(127) = &H6DD2D2BF
    Te0(128) = &H4CCDCD81: Te0(129) = &H140C0C18: Te0(130) = &H35131326: Te0(131) = &H2FECECC3: Te0(132) = &HE15F5FBE: Te0(133) = &HA2979735: Te0(134) = &HCC444488: Te0(135) = &H3917172E
    Te0(136) = &H57C4C493: Te0(137) = &HF2A7A755: Te0(138) = &H827E7EFC: Te0(139) = &H473D3D7A: Te0(140) = &HAC6464C8: Te0(141) = &HE75D5DBA: Te0(142) = &H2B191932: Te0(143) = &H957373E6
    Te0(144) = &HA06060C0: Te0(145) = &H98818119: Te0(146) = &HD14F4F9E: Te0(147) = &H7FDCDCA3: Te0(148) = &H66222244: Te0(149) = &H7E2A2A54: Te0(150) = &HAB90903B: Te0(151) = &H8388880B
    Te0(152) = &HCA46468C: Te0(153) = &H29EEEEC7: Te0(154) = &HD3B8B86B: Te0(155) = &H3C141428: Te0(156) = &H79DEDEA7: Te0(157) = &HE25E5EBC: Te0(158) = &H1D0B0B16: Te0(159) = &H76DBDBAD
    Te0(160) = &H3BE0E0DB: Te0(161) = &H56323264: Te0(162) = &H4E3A3A74: Te0(163) = &H1E0A0A14: Te0(164) = &HDB494992: Te0(165) = &HA06060C:  Te0(166) = &H6C242448: Te0(167) = &HE45C5CB8
    Te0(168) = &H5DC2C29F: Te0(169) = &H6ED3D3BD: Te0(170) = &HEFACAC43: Te0(171) = &HA66262C4: Te0(172) = &HA8919139: Te0(173) = &HA4959531: Te0(174) = &H37E4E4D3: Te0(175) = &H8B7979F2
    Te0(176) = &H32E7E7D5: Te0(177) = &H43C8C88B: Te0(178) = &H5937376E: Te0(179) = &HB76D6DDA: Te0(180) = &H8C8D8D01: Te0(181) = &H64D5D5B1: Te0(182) = &HD24E4E9C: Te0(183) = &HE0A9A949
    Te0(184) = &HB46C6CD8: Te0(185) = &HFA5656AC: Te0(186) = &H7F4F4F3:  Te0(187) = &H25EAEACF: Te0(188) = &HAF6565CA: Te0(189) = &H8E7A7AF4: Te0(190) = &HE9AEAE47: Te0(191) = &H18080810
    Te0(192) = &HD5BABA6F: Te0(193) = &H887878F0: Te0(194) = &H6F25254A: Te0(195) = &H722E2E5C: Te0(196) = &H241C1C38: Te0(197) = &HF1A6A657: Te0(198) = &HC7B4B473: Te0(199) = &H51C6C697
    Te0(200) = &H23E8E8CB: Te0(201) = &H7CDDDDA1: Te0(202) = &H9C7474E8: Te0(203) = &H211F1F3E: Te0(204) = &HDD4B4B96: Te0(205) = &HDCBDBD61: Te0(206) = &H868B8B0D: Te0(207) = &H858A8A0F
    Te0(208) = &H907070E0: Te0(209) = &H423E3E7C: Te0(210) = &HC4B5B571: Te0(211) = &HAA6666CC: Te0(212) = &HD8484890: Te0(213) = &H5030306:  Te0(214) = &H1F6F6F7:  Te0(215) = &H120E0E1C
    Te0(216) = &HA36161C2: Te0(217) = &H5F35356A: Te0(218) = &HF95757AE: Te0(219) = &HD0B9B969: Te0(220) = &H91868617: Te0(221) = &H58C1C199: Te0(222) = &H271D1D3A: Te0(223) = &HB99E9E27
    Te0(224) = &H38E1E1D9: Te0(225) = &H13F8F8EB: Te0(226) = &HB398982B: Te0(227) = &H33111122: Te0(228) = &HBB6969D2: Te0(229) = &H70D9D9A9: Te0(230) = &H898E8E07: Te0(231) = &HA7949433
    Te0(232) = &HB69B9B2D: Te0(233) = &H221E1E3C: Te0(234) = &H92878715: Te0(235) = &H20E9E9C9: Te0(236) = &H49CECE87: Te0(237) = &HFF5555AA: Te0(238) = &H78282850: Te0(239) = &H7ADFDFA5
    Te0(240) = &H8F8C8C03: Te0(241) = &HF8A1A159: Te0(242) = &H80898909: Te0(243) = &H170D0D1A: Te0(244) = &HDABFBF65: Te0(245) = &H31E6E6D7: Te0(246) = &HC6424284: Te0(247) = &HB86868D0
    Te0(248) = &HC3414182: Te0(249) = &HB0999929: Te0(250) = &H772D2D5A: Te0(251) = &H110F0F1E: Te0(252) = &HCBB0B07B: Te0(253) = &HFC5454A8: Te0(254) = &HD6BBBB6D: Te0(255) = &H3A16162C

    Te1(0) = &H6363C6A5:   Te1(1) = &H7C7CF884:   Te1(2) = &H7777EE99:   Te1(3) = &H7B7BF68D:   Te1(4) = &HF2F2FF0D:   Te1(5) = &H6B6BD6BD:   Te1(6) = &H6F6FDEB1:   Te1(7) = &HC5C59154
    Te1(8) = &H30306050:   Te1(9) = &H1010203:    Te1(10) = &H6767CEA9:  Te1(11) = &H2B2B567D:  Te1(12) = &HFEFEE719:  Te1(13) = &HD7D7B562:  Te1(14) = &HABAB4DE6:  Te1(15) = &H7676EC9A
    Te1(16) = &HCACA8F45:  Te1(17) = &H82821F9D:  Te1(18) = &HC9C98940:  Te1(19) = &H7D7DFA87:  Te1(20) = &HFAFAEF15:  Te1(21) = &H5959B2EB:  Te1(22) = &H47478EC9:  Te1(23) = &HF0F0FB0B
    Te1(24) = &HADAD41EC:  Te1(25) = &HD4D4B367:  Te1(26) = &HA2A25FFD:  Te1(27) = &HAFAF45EA:  Te1(28) = &H9C9C23BF:  Te1(29) = &HA4A453F7:  Te1(30) = &H7272E496:  Te1(31) = &HC0C09B5B
    Te1(32) = &HB7B775C2:  Te1(33) = &HFDFDE11C:  Te1(34) = &H93933DAE:  Te1(35) = &H26264C6A:  Te1(36) = &H36366C5A:  Te1(37) = &H3F3F7E41:  Te1(38) = &HF7F7F502:  Te1(39) = &HCCCC834F
    Te1(40) = &H3434685C:  Te1(41) = &HA5A551F4:  Te1(42) = &HE5E5D134:  Te1(43) = &HF1F1F908:  Te1(44) = &H7171E293:  Te1(45) = &HD8D8AB73:  Te1(46) = &H31316253:  Te1(47) = &H15152A3F
    Te1(48) = &H404080C:   Te1(49) = &HC7C79552:  Te1(50) = &H23234665:  Te1(51) = &HC3C39D5E:  Te1(52) = &H18183028:  Te1(53) = &H969637A1:  Te1(54) = &H5050A0F:   Te1(55) = &H9A9A2FB5
    Te1(56) = &H7070E09:   Te1(57) = &H12122436:  Te1(58) = &H80801B9B:  Te1(59) = &HE2E2DF3D:  Te1(60) = &HEBEBCD26:  Te1(61) = &H27274E69:  Te1(62) = &HB2B27FCD:  Te1(63) = &H7575EA9F
    Te1(64) = &H909121B:   Te1(65) = &H83831D9E:  Te1(66) = &H2C2C5874:  Te1(67) = &H1A1A342E:  Te1(68) = &H1B1B362D:  Te1(69) = &H6E6EDCB2:  Te1(70) = &H5A5AB4EE:  Te1(71) = &HA0A05BFB
    Te1(72) = &H5252A4F6:  Te1(73) = &H3B3B764D:  Te1(74) = &HD6D6B761:  Te1(75) = &HB3B37DCE:  Te1(76) = &H2929527B:  Te1(77) = &HE3E3DD3E:  Te1(78) = &H2F2F5E71:  Te1(79) = &H84841397
    Te1(80) = &H5353A6F5:  Te1(81) = &HD1D1B968:  Te1(82) = &H0&:        Te1(83) = &HEDEDC12C:  Te1(84) = &H20204060:  Te1(85) = &HFCFCE31F:  Te1(86) = &HB1B179C8:  Te1(87) = &H5B5BB6ED
    Te1(88) = &H6A6AD4BE:  Te1(89) = &HCBCB8D46:  Te1(90) = &HBEBE67D9:  Te1(91) = &H3939724B:  Te1(92) = &H4A4A94DE:  Te1(93) = &H4C4C98D4:  Te1(94) = &H5858B0E8:  Te1(95) = &HCFCF854A
    Te1(96) = &HD0D0BB6B:  Te1(97) = &HEFEFC52A:  Te1(98) = &HAAAA4FE5:  Te1(99) = &HFBFBED16:  Te1(100) = &H434386C5: Te1(101) = &H4D4D9AD7: Te1(102) = &H33336655: Te1(103) = &H85851194
    Te1(104) = &H45458ACF: Te1(105) = &HF9F9E910: Te1(106) = &H2020406:  Te1(107) = &H7F7FFE81: Te1(108) = &H5050A0F0: Te1(109) = &H3C3C7844: Te1(110) = &H9F9F25BA: Te1(111) = &HA8A84BE3
    Te1(112) = &H5151A2F3: Te1(113) = &HA3A35DFE: Te1(114) = &H404080C0: Te1(115) = &H8F8F058A: Te1(116) = &H92923FAD: Te1(117) = &H9D9D21BC: Te1(118) = &H38387048: Te1(119) = &HF5F5F104
    Te1(120) = &HBCBC63DF: Te1(121) = &HB6B677C1: Te1(122) = &HDADAAF75: Te1(123) = &H21214263: Te1(124) = &H10102030: Te1(125) = &HFFFFE51A: Te1(126) = &HF3F3FD0E: Te1(127) = &HD2D2BF6D
    Te1(128) = &HCDCD814C: Te1(129) = &HC0C1814:  Te1(130) = &H13132635: Te1(131) = &HECECC32F: Te1(132) = &H5F5FBEE1: Te1(133) = &H979735A2: Te1(134) = &H444488CC: Te1(135) = &H17172E39
    Te1(136) = &HC4C49357: Te1(137) = &HA7A755F2: Te1(138) = &H7E7EFC82: Te1(139) = &H3D3D7A47: Te1(140) = &H6464C8AC: Te1(141) = &H5D5DBAE7: Te1(142) = &H1919322B: Te1(143) = &H7373E695
    Te1(144) = &H6060C0A0: Te1(145) = &H81811998: Te1(146) = &H4F4F9ED1: Te1(147) = &HDCDCA37F: Te1(148) = &H22224466: Te1(149) = &H2A2A547E: Te1(150) = &H90903BAB: Te1(151) = &H88880B83
    Te1(152) = &H46468CCA: Te1(153) = &HEEEEC729: Te1(154) = &HB8B86BD3: Te1(155) = &H1414283C: Te1(156) = &HDEDEA779: Te1(157) = &H5E5EBCE2: Te1(158) = &HB0B161D:  Te1(159) = &HDBDBAD76
    Te1(160) = &HE0E0DB3B: Te1(161) = &H32326456: Te1(162) = &H3A3A744E: Te1(163) = &HA0A141E:  Te1(164) = &H494992DB: Te1(165) = &H6060C0A:  Te1(166) = &H2424486C: Te1(167) = &H5C5CB8E4
    Te1(168) = &HC2C29F5D: Te1(169) = &HD3D3BD6E: Te1(170) = &HACAC43EF: Te1(171) = &H6262C4A6: Te1(172) = &H919139A8: Te1(173) = &H959531A4: Te1(174) = &HE4E4D337: Te1(175) = &H7979F28B
    Te1(176) = &HE7E7D532: Te1(177) = &HC8C88B43: Te1(178) = &H37376E59: Te1(179) = &H6D6DDAB7: Te1(180) = &H8D8D018C: Te1(181) = &HD5D5B164: Te1(182) = &H4E4E9CD2: Te1(183) = &HA9A949E0
    Te1(184) = &H6C6CD8B4: Te1(185) = &H5656ACFA: Te1(186) = &HF4F4F307: Te1(187) = &HEAEACF25: Te1(188) = &H6565CAAF: Te1(189) = &H7A7AF48E: Te1(190) = &HAEAE47E9: Te1(191) = &H8081018
    Te1(192) = &HBABA6FD5: Te1(193) = &H7878F088: Te1(194) = &H25254A6F: Te1(195) = &H2E2E5C72: Te1(196) = &H1C1C3824: Te1(197) = &HA6A657F1: Te1(198) = &HB4B473C7: Te1(199) = &HC6C69751
    Te1(200) = &HE8E8CB23: Te1(201) = &HDDDDA17C: Te1(202) = &H7474E89C: Te1(203) = &H1F1F3E21: Te1(204) = &H4B4B96DD: Te1(205) = &HBDBD61DC: Te1(206) = &H8B8B0D86: Te1(207) = &H8A8A0F85
    Te1(208) = &H7070E090: Te1(209) = &H3E3E7C42: Te1(210) = &HB5B571C4: Te1(211) = &H6666CCAA: Te1(212) = &H484890D8: Te1(213) = &H3030605:  Te1(214) = &HF6F6F701: Te1(215) = &HE0E1C12
    Te1(216) = &H6161C2A3: Te1(217) = &H35356A5F: Te1(218) = &H5757AEF9: Te1(219) = &HB9B969D0: Te1(220) = &H86861791: Te1(221) = &HC1C19958: Te1(222) = &H1D1D3A27: Te1(223) = &H9E9E27B9
    Te1(224) = &HE1E1D938: Te1(225) = &HF8F8EB13: Te1(226) = &H98982BB3: Te1(227) = &H11112233: Te1(228) = &H6969D2BB: Te1(229) = &HD9D9A970: Te1(230) = &H8E8E0789: Te1(231) = &H949433A7
    Te1(232) = &H9B9B2DB6: Te1(233) = &H1E1E3C22: Te1(234) = &H87871592: Te1(235) = &HE9E9C920: Te1(236) = &HCECE8749: Te1(237) = &H5555AAFF: Te1(238) = &H28285078: Te1(239) = &HDFDFA57A
    Te1(240) = &H8C8C038F: Te1(241) = &HA1A159F8: Te1(242) = &H89890980: Te1(243) = &HD0D1A17:  Te1(244) = &HBFBF65DA: Te1(245) = &HE6E6D731: Te1(246) = &H424284C6: Te1(247) = &H6868D0B8
    Te1(248) = &H414182C3: Te1(249) = &H999929B0: Te1(250) = &H2D2D5A77: Te1(251) = &HF0F1E11:  Te1(252) = &HB0B07BCB: Te1(253) = &H5454A8FC: Te1(254) = &HBBBB6DD6: Te1(255) = &H16162C3A

    Te2(0) = &H63C6A563:   Te2(1) = &H7CF8847C:   Te2(2) = &H77EE9977:   Te2(3) = &H7BF68D7B:   Te2(4) = &HF2FF0DF2:   Te2(5) = &H6BD6BD6B:   Te2(6) = &H6FDEB16F:   Te2(7) = &HC59154C5
    Te2(8) = &H30605030:   Te2(9) = &H1020301:    Te2(10) = &H67CEA967:  Te2(11) = &H2B567D2B:  Te2(12) = &HFEE719FE:  Te2(13) = &HD7B562D7:  Te2(14) = &HAB4DE6AB:  Te2(15) = &H76EC9A76
    Te2(16) = &HCA8F45CA:  Te2(17) = &H821F9D82:  Te2(18) = &HC98940C9:  Te2(19) = &H7DFA877D:  Te2(20) = &HFAEF15FA:  Te2(21) = &H59B2EB59:  Te2(22) = &H478EC947:  Te2(23) = &HF0FB0BF0
    Te2(24) = &HAD41ECAD:  Te2(25) = &HD4B367D4:  Te2(26) = &HA25FFDA2:  Te2(27) = &HAF45EAAF:  Te2(28) = &H9C23BF9C:  Te2(29) = &HA453F7A4:  Te2(30) = &H72E49672:  Te2(31) = &HC09B5BC0
    Te2(32) = &HB775C2B7:  Te2(33) = &HFDE11CFD:  Te2(34) = &H933DAE93:  Te2(35) = &H264C6A26:  Te2(36) = &H366C5A36:  Te2(37) = &H3F7E413F:  Te2(38) = &HF7F502F7:  Te2(39) = &HCC834FCC
    Te2(40) = &H34685C34:  Te2(41) = &HA551F4A5:  Te2(42) = &HE5D134E5:  Te2(43) = &HF1F908F1:  Te2(44) = &H71E29371:  Te2(45) = &HD8AB73D8:  Te2(46) = &H31625331:  Te2(47) = &H152A3F15
    Te2(48) = &H4080C04:   Te2(49) = &HC79552C7:  Te2(50) = &H23466523:  Te2(51) = &HC39D5EC3:  Te2(52) = &H18302818:  Te2(53) = &H9637A196:  Te2(54) = &H50A0F05:   Te2(55) = &H9A2FB59A
    Te2(56) = &H70E0907:   Te2(57) = &H12243612:  Te2(58) = &H801B9B80:  Te2(59) = &HE2DF3DE2:  Te2(60) = &HEBCD26EB:  Te2(61) = &H274E6927:  Te2(62) = &HB27FCDB2:  Te2(63) = &H75EA9F75
    Te2(64) = &H9121B09:   Te2(65) = &H831D9E83:  Te2(66) = &H2C58742C:  Te2(67) = &H1A342E1A:  Te2(68) = &H1B362D1B:  Te2(69) = &H6EDCB26E:  Te2(70) = &H5AB4EE5A:  Te2(71) = &HA05BFBA0
    Te2(72) = &H52A4F652:  Te2(73) = &H3B764D3B:  Te2(74) = &HD6B761D6:  Te2(75) = &HB37DCEB3:  Te2(76) = &H29527B29:  Te2(77) = &HE3DD3EE3:  Te2(78) = &H2F5E712F:  Te2(79) = &H84139784
    Te2(80) = &H53A6F553:  Te2(81) = &HD1B968D1:  Te2(82) = &H0&:        Te2(83) = &HEDC12CED:  Te2(84) = &H20406020:  Te2(85) = &HFCE31FFC:  Te2(86) = &HB179C8B1:  Te2(87) = &H5BB6ED5B
    Te2(88) = &H6AD4BE6A:  Te2(89) = &HCB8D46CB:  Te2(90) = &HBE67D9BE:  Te2(91) = &H39724B39:  Te2(92) = &H4A94DE4A:  Te2(93) = &H4C98D44C:  Te2(94) = &H58B0E858:  Te2(95) = &HCF854ACF
    Te2(96) = &HD0BB6BD0:  Te2(97) = &HEFC52AEF:  Te2(98) = &HAA4FE5AA:  Te2(99) = &HFBED16FB:  Te2(100) = &H4386C543: Te2(101) = &H4D9AD74D: Te2(102) = &H33665533: Te2(103) = &H85119485
    Te2(104) = &H458ACF45: Te2(105) = &HF9E910F9: Te2(106) = &H2040602:  Te2(107) = &H7FFE817F: Te2(108) = &H50A0F050: Te2(109) = &H3C78443C: Te2(110) = &H9F25BA9F: Te2(111) = &HA84BE3A8
    Te2(112) = &H51A2F351: Te2(113) = &HA35DFEA3: Te2(114) = &H4080C040: Te2(115) = &H8F058A8F: Te2(116) = &H923FAD92: Te2(117) = &H9D21BC9D: Te2(118) = &H38704838: Te2(119) = &HF5F104F5
    Te2(120) = &HBC63DFBC: Te2(121) = &HB677C1B6: Te2(122) = &HDAAF75DA: Te2(123) = &H21426321: Te2(124) = &H10203010: Te2(125) = &HFFE51AFF: Te2(126) = &HF3FD0EF3: Te2(127) = &HD2BF6DD2
    Te2(128) = &HCD814CCD: Te2(129) = &HC18140C:  Te2(130) = &H13263513: Te2(131) = &HECC32FEC: Te2(132) = &H5FBEE15F: Te2(133) = &H9735A297: Te2(134) = &H4488CC44: Te2(135) = &H172E3917
    Te2(136) = &HC49357C4: Te2(137) = &HA755F2A7: Te2(138) = &H7EFC827E: Te2(139) = &H3D7A473D: Te2(140) = &H64C8AC64: Te2(141) = &H5DBAE75D: Te2(142) = &H19322B19: Te2(143) = &H73E69573
    Te2(144) = &H60C0A060: Te2(145) = &H81199881: Te2(146) = &H4F9ED14F: Te2(147) = &HDCA37FDC: Te2(148) = &H22446622: Te2(149) = &H2A547E2A: Te2(150) = &H903BAB90: Te2(151) = &H880B8388
    Te2(152) = &H468CCA46: Te2(153) = &HEEC729EE: Te2(154) = &HB86BD3B8: Te2(155) = &H14283C14: Te2(156) = &HDEA779DE: Te2(157) = &H5EBCE25E: Te2(158) = &HB161D0B:  Te2(159) = &HDBAD76DB
    Te2(160) = &HE0DB3BE0: Te2(161) = &H32645632: Te2(162) = &H3A744E3A: Te2(163) = &HA141E0A:  Te2(164) = &H4992DB49: Te2(165) = &H60C0A06:  Te2(166) = &H24486C24: Te2(167) = &H5CB8E45C
    Te2(168) = &HC29F5DC2: Te2(169) = &HD3BD6ED3: Te2(170) = &HAC43EFAC: Te2(171) = &H62C4A662: Te2(172) = &H9139A891: Te2(173) = &H9531A495: Te2(174) = &HE4D337E4: Te2(175) = &H79F28B79
    Te2(176) = &HE7D532E7: Te2(177) = &HC88B43C8: Te2(178) = &H376E5937: Te2(179) = &H6DDAB76D: Te2(180) = &H8D018C8D: Te2(181) = &HD5B164D5: Te2(182) = &H4E9CD24E: Te2(183) = &HA949E0A9
    Te2(184) = &H6CD8B46C: Te2(185) = &H56ACFA56: Te2(186) = &HF4F307F4: Te2(187) = &HEACF25EA: Te2(188) = &H65CAAF65: Te2(189) = &H7AF48E7A: Te2(190) = &HAE47E9AE: Te2(191) = &H8101808
    Te2(192) = &HBA6FD5BA: Te2(193) = &H78F08878: Te2(194) = &H254A6F25: Te2(195) = &H2E5C722E: Te2(196) = &H1C38241C: Te2(197) = &HA657F1A6: Te2(198) = &HB473C7B4: Te2(199) = &HC69751C6
    Te2(200) = &HE8CB23E8: Te2(201) = &HDDA17CDD: Te2(202) = &H74E89C74: Te2(203) = &H1F3E211F: Te2(204) = &H4B96DD4B: Te2(205) = &HBD61DCBD: Te2(206) = &H8B0D868B: Te2(207) = &H8A0F858A
    Te2(208) = &H70E09070: Te2(209) = &H3E7C423E: Te2(210) = &HB571C4B5: Te2(211) = &H66CCAA66: Te2(212) = &H4890D848: Te2(213) = &H3060503:  Te2(214) = &HF6F701F6: Te2(215) = &HE1C120E
    Te2(216) = &H61C2A361: Te2(217) = &H356A5F35: Te2(218) = &H57AEF957: Te2(219) = &HB969D0B9: Te2(220) = &H86179186: Te2(221) = &HC19958C1: Te2(222) = &H1D3A271D: Te2(223) = &H9E27B99E
    Te2(224) = &HE1D938E1: Te2(225) = &HF8EB13F8: Te2(226) = &H982BB398: Te2(227) = &H11223311: Te2(228) = &H69D2BB69: Te2(229) = &HD9A970D9: Te2(230) = &H8E07898E: Te2(231) = &H9433A794
    Te2(232) = &H9B2DB69B: Te2(233) = &H1E3C221E: Te2(234) = &H87159287: Te2(235) = &HE9C920E9: Te2(236) = &HCE8749CE: Te2(237) = &H55AAFF55: Te2(238) = &H28507828: Te2(239) = &HDFA57ADF
    Te2(240) = &H8C038F8C: Te2(241) = &HA159F8A1: Te2(242) = &H89098089: Te2(243) = &HD1A170D:  Te2(244) = &HBF65DABF: Te2(245) = &HE6D731E6: Te2(246) = &H4284C642: Te2(247) = &H68D0B868
    Te2(248) = &H4182C341: Te2(249) = &H9929B099: Te2(250) = &H2D5A772D: Te2(251) = &HF1E110F:  Te2(252) = &HB07BCBB0: Te2(253) = &H54A8FC54: Te2(254) = &HBB6DD6BB: Te2(255) = &H162C3A16

    Te3(0) = &HC6A56363:   Te3(1) = &HF8847C7C:   Te3(2) = &HEE997777:   Te3(3) = &HF68D7B7B:   Te3(4) = &HFF0DF2F2:   Te3(5) = &HD6BD6B6B:   Te3(6) = &HDEB16F6F:   Te3(7) = &H9154C5C5
    Te3(8) = &H60503030:   Te3(9) = &H2030101:    Te3(10) = &HCEA96767:  Te3(11) = &H567D2B2B:  Te3(12) = &HE719FEFE:  Te3(13) = &HB562D7D7:  Te3(14) = &H4DE6ABAB:  Te3(15) = &HEC9A7676
    Te3(16) = &H8F45CACA:  Te3(17) = &H1F9D8282:  Te3(18) = &H8940C9C9:  Te3(19) = &HFA877D7D:  Te3(20) = &HEF15FAFA:  Te3(21) = &HB2EB5959:  Te3(22) = &H8EC94747:  Te3(23) = &HFB0BF0F0
    Te3(24) = &H41ECADAD:  Te3(25) = &HB367D4D4:  Te3(26) = &H5FFDA2A2:  Te3(27) = &H45EAAFAF:  Te3(28) = &H23BF9C9C:  Te3(29) = &H53F7A4A4:  Te3(30) = &HE4967272:  Te3(31) = &H9B5BC0C0
    Te3(32) = &H75C2B7B7:  Te3(33) = &HE11CFDFD:  Te3(34) = &H3DAE9393:  Te3(35) = &H4C6A2626:  Te3(36) = &H6C5A3636:  Te3(37) = &H7E413F3F:  Te3(38) = &HF502F7F7:  Te3(39) = &H834FCCCC
    Te3(40) = &H685C3434:  Te3(41) = &H51F4A5A5:  Te3(42) = &HD134E5E5:  Te3(43) = &HF908F1F1:  Te3(44) = &HE2937171:  Te3(45) = &HAB73D8D8:  Te3(46) = &H62533131:  Te3(47) = &H2A3F1515
    Te3(48) = &H80C0404:   Te3(49) = &H9552C7C7:  Te3(50) = &H46652323:  Te3(51) = &H9D5EC3C3:  Te3(52) = &H30281818:  Te3(53) = &H37A19696:  Te3(54) = &HA0F0505:   Te3(55) = &H2FB59A9A
    Te3(56) = &HE090707:   Te3(57) = &H24361212:  Te3(58) = &H1B9B8080:  Te3(59) = &HDF3DE2E2:  Te3(60) = &HCD26EBEB:  Te3(61) = &H4E692727:  Te3(62) = &H7FCDB2B2:  Te3(63) = &HEA9F7575
    Te3(64) = &H121B0909:  Te3(65) = &H1D9E8383:  Te3(66) = &H58742C2C:  Te3(67) = &H342E1A1A:  Te3(68) = &H362D1B1B:  Te3(69) = &HDCB26E6E:  Te3(70) = &HB4EE5A5A:  Te3(71) = &H5BFBA0A0
    Te3(72) = &HA4F65252:  Te3(73) = &H764D3B3B:  Te3(74) = &HB761D6D6:  Te3(75) = &H7DCEB3B3:  Te3(76) = &H527B2929:  Te3(77) = &HDD3EE3E3:  Te3(78) = &H5E712F2F:  Te3(79) = &H13978484
    Te3(80) = &HA6F55353:  Te3(81) = &HB968D1D1:  Te3(82) = &H0&:        Te3(83) = &HC12CEDED:  Te3(84) = &H40602020:  Te3(85) = &HE31FFCFC:  Te3(86) = &H79C8B1B1:  Te3(87) = &HB6ED5B5B
    Te3(88) = &HD4BE6A6A:  Te3(89) = &H8D46CBCB:  Te3(90) = &H67D9BEBE:  Te3(91) = &H724B3939:  Te3(92) = &H94DE4A4A:  Te3(93) = &H98D44C4C:  Te3(94) = &HB0E85858:  Te3(95) = &H854ACFCF
    Te3(96) = &HBB6BD0D0:  Te3(97) = &HC52AEFEF:  Te3(98) = &H4FE5AAAA:  Te3(99) = &HED16FBFB:  Te3(100) = &H86C54343: Te3(101) = &H9AD74D4D: Te3(102) = &H66553333: Te3(103) = &H11948585
    Te3(104) = &H8ACF4545: Te3(105) = &HE910F9F9: Te3(106) = &H4060202:  Te3(107) = &HFE817F7F: Te3(108) = &HA0F05050: Te3(109) = &H78443C3C: Te3(110) = &H25BA9F9F: Te3(111) = &H4BE3A8A8
    Te3(112) = &HA2F35151: Te3(113) = &H5DFEA3A3: Te3(114) = &H80C04040: Te3(115) = &H58A8F8F:  Te3(116) = &H3FAD9292: Te3(117) = &H21BC9D9D: Te3(118) = &H70483838: Te3(119) = &HF104F5F5
    Te3(120) = &H63DFBCBC: Te3(121) = &H77C1B6B6: Te3(122) = &HAF75DADA: Te3(123) = &H42632121: Te3(124) = &H20301010: Te3(125) = &HE51AFFFF: Te3(126) = &HFD0EF3F3: Te3(127) = &HBF6DD2D2
    Te3(128) = &H814CCDCD: Te3(129) = &H18140C0C: Te3(130) = &H26351313: Te3(131) = &HC32FECEC: Te3(132) = &HBEE15F5F: Te3(133) = &H35A29797: Te3(134) = &H88CC4444: Te3(135) = &H2E391717
    Te3(136) = &H9357C4C4: Te3(137) = &H55F2A7A7: Te3(138) = &HFC827E7E: Te3(139) = &H7A473D3D: Te3(140) = &HC8AC6464: Te3(141) = &HBAE75D5D: Te3(142) = &H322B1919: Te3(143) = &HE6957373
    Te3(144) = &HC0A06060: Te3(145) = &H19988181: Te3(146) = &H9ED14F4F: Te3(147) = &HA37FDCDC: Te3(148) = &H44662222: Te3(149) = &H547E2A2A: Te3(150) = &H3BAB9090: Te3(151) = &HB838888
    Te3(152) = &H8CCA4646: Te3(153) = &HC729EEEE: Te3(154) = &H6BD3B8B8: Te3(155) = &H283C1414: Te3(156) = &HA779DEDE: Te3(157) = &HBCE25E5E: Te3(158) = &H161D0B0B: Te3(159) = &HAD76DBDB
    Te3(160) = &HDB3BE0E0: Te3(161) = &H64563232: Te3(162) = &H744E3A3A: Te3(163) = &H141E0A0A: Te3(164) = &H92DB4949: Te3(165) = &HC0A0606:  Te3(166) = &H486C2424: Te3(167) = &HB8E45C5C
    Te3(168) = &H9F5DC2C2: Te3(169) = &HBD6ED3D3: Te3(170) = &H43EFACAC: Te3(171) = &HC4A66262: Te3(172) = &H39A89191: Te3(173) = &H31A49595: Te3(174) = &HD337E4E4: Te3(175) = &HF28B7979
    Te3(176) = &HD532E7E7: Te3(177) = &H8B43C8C8: Te3(178) = &H6E593737: Te3(179) = &HDAB76D6D: Te3(180) = &H18C8D8D:  Te3(181) = &HB164D5D5: Te3(182) = &H9CD24E4E: Te3(183) = &H49E0A9A9
    Te3(184) = &HD8B46C6C: Te3(185) = &HACFA5656: Te3(186) = &HF307F4F4: Te3(187) = &HCF25EAEA: Te3(188) = &HCAAF6565: Te3(189) = &HF48E7A7A: Te3(190) = &H47E9AEAE: Te3(191) = &H10180808
    Te3(192) = &H6FD5BABA: Te3(193) = &HF0887878: Te3(194) = &H4A6F2525: Te3(195) = &H5C722E2E: Te3(196) = &H38241C1C: Te3(197) = &H57F1A6A6: Te3(198) = &H73C7B4B4: Te3(199) = &H9751C6C6
    Te3(200) = &HCB23E8E8: Te3(201) = &HA17CDDDD: Te3(202) = &HE89C7474: Te3(203) = &H3E211F1F: Te3(204) = &H96DD4B4B: Te3(205) = &H61DCBDBD: Te3(206) = &HD868B8B:  Te3(207) = &HF858A8A
    Te3(208) = &HE0907070: Te3(209) = &H7C423E3E: Te3(210) = &H71C4B5B5: Te3(211) = &HCCAA6666: Te3(212) = &H90D84848: Te3(213) = &H6050303:  Te3(214) = &HF701F6F6: Te3(215) = &H1C120E0E
    Te3(216) = &HC2A36161: Te3(217) = &H6A5F3535: Te3(218) = &HAEF95757: Te3(219) = &H69D0B9B9: Te3(220) = &H17918686: Te3(221) = &H9958C1C1: Te3(222) = &H3A271D1D: Te3(223) = &H27B99E9E
    Te3(224) = &HD938E1E1: Te3(225) = &HEB13F8F8: Te3(226) = &H2BB39898: Te3(227) = &H22331111: Te3(228) = &HD2BB6969: Te3(229) = &HA970D9D9: Te3(230) = &H7898E8E:  Te3(231) = &H33A79494
    Te3(232) = &H2DB69B9B: Te3(233) = &H3C221E1E: Te3(234) = &H15928787: Te3(235) = &HC920E9E9: Te3(236) = &H8749CECE: Te3(237) = &HAAFF5555: Te3(238) = &H50782828: Te3(239) = &HA57ADFDF
    Te3(240) = &H38F8C8C:  Te3(241) = &H59F8A1A1: Te3(242) = &H9808989:  Te3(243) = &H1A170D0D: Te3(244) = &H65DABFBF: Te3(245) = &HD731E6E6: Te3(246) = &H84C64242: Te3(247) = &HD0B86868
    Te3(248) = &H82C34141: Te3(249) = &H29B09999: Te3(250) = &H5A772D2D: Te3(251) = &H1E110F0F: Te3(252) = &H7BCBB0B0: Te3(253) = &HA8FC5454: Te3(254) = &H6DD6BBBB: Te3(255) = &H2C3A1616

    Te4(0) = &H63636363:   Te4(1) = &H7C7C7C7C:   Te4(2) = &H77777777:   Te4(3) = &H7B7B7B7B:   Te4(4) = &HF2F2F2F2:   Te4(5) = &H6B6B6B6B:   Te4(6) = &H6F6F6F6F:   Te4(7) = &HC5C5C5C5
    Te4(8) = &H30303030:   Te4(9) = &H1010101:    Te4(10) = &H67676767:  Te4(11) = &H2B2B2B2B:  Te4(12) = &HFEFEFEFE:  Te4(13) = &HD7D7D7D7:  Te4(14) = &HABABABAB:  Te4(15) = &H76767676
    Te4(16) = &HCACACACA:  Te4(17) = &H82828282:  Te4(18) = &HC9C9C9C9:  Te4(19) = &H7D7D7D7D:  Te4(20) = &HFAFAFAFA:  Te4(21) = &H59595959:  Te4(22) = &H47474747:  Te4(23) = &HF0F0F0F0
    Te4(24) = &HADADADAD:  Te4(25) = &HD4D4D4D4:  Te4(26) = &HA2A2A2A2:  Te4(27) = &HAFAFAFAF:  Te4(28) = &H9C9C9C9C:  Te4(29) = &HA4A4A4A4:  Te4(30) = &H72727272:  Te4(31) = &HC0C0C0C0
    Te4(32) = &HB7B7B7B7:  Te4(33) = &HFDFDFDFD:  Te4(34) = &H93939393:  Te4(35) = &H26262626:  Te4(36) = &H36363636:  Te4(37) = &H3F3F3F3F:  Te4(38) = &HF7F7F7F7:  Te4(39) = &HCCCCCCCC
    Te4(40) = &H34343434:  Te4(41) = &HA5A5A5A5:  Te4(42) = &HE5E5E5E5:  Te4(43) = &HF1F1F1F1:  Te4(44) = &H71717171:  Te4(45) = &HD8D8D8D8:  Te4(46) = &H31313131:  Te4(47) = &H15151515
    Te4(48) = &H4040404:   Te4(49) = &HC7C7C7C7:  Te4(50) = &H23232323:  Te4(51) = &HC3C3C3C3:  Te4(52) = &H18181818:  Te4(53) = &H96969696:  Te4(54) = &H5050505:   Te4(55) = &H9A9A9A9A
    Te4(56) = &H7070707:   Te4(57) = &H12121212:  Te4(58) = &H80808080:  Te4(59) = &HE2E2E2E2:  Te4(60) = &HEBEBEBEB:  Te4(61) = &H27272727:  Te4(62) = &HB2B2B2B2:  Te4(63) = &H75757575
    Te4(64) = &H9090909:   Te4(65) = &H83838383:  Te4(66) = &H2C2C2C2C:  Te4(67) = &H1A1A1A1A:  Te4(68) = &H1B1B1B1B:  Te4(69) = &H6E6E6E6E:  Te4(70) = &H5A5A5A5A:  Te4(71) = &HA0A0A0A0
    Te4(72) = &H52525252:  Te4(73) = &H3B3B3B3B:  Te4(74) = &HD6D6D6D6:  Te4(75) = &HB3B3B3B3:  Te4(76) = &H29292929:  Te4(77) = &HE3E3E3E3:  Te4(78) = &H2F2F2F2F:  Te4(79) = &H84848484
    Te4(80) = &H53535353:  Te4(81) = &HD1D1D1D1:  Te4(82) = &H0&:        Te4(83) = &HEDEDEDED:  Te4(84) = &H20202020:  Te4(85) = &HFCFCFCFC:  Te4(86) = &HB1B1B1B1:  Te4(87) = &H5B5B5B5B
    Te4(88) = &H6A6A6A6A:  Te4(89) = &HCBCBCBCB:  Te4(90) = &HBEBEBEBE:  Te4(91) = &H39393939:  Te4(92) = &H4A4A4A4A:  Te4(93) = &H4C4C4C4C:  Te4(94) = &H58585858:  Te4(95) = &HCFCFCFCF
    Te4(96) = &HD0D0D0D0:  Te4(97) = &HEFEFEFEF:  Te4(98) = &HAAAAAAAA:  Te4(99) = &HFBFBFBFB:  Te4(100) = &H43434343: Te4(101) = &H4D4D4D4D: Te4(102) = &H33333333: Te4(103) = &H85858585
    Te4(104) = &H45454545: Te4(105) = &HF9F9F9F9: Te4(106) = &H2020202:  Te4(107) = &H7F7F7F7F: Te4(108) = &H50505050: Te4(109) = &H3C3C3C3C: Te4(110) = &H9F9F9F9F: Te4(111) = &HA8A8A8A8
    Te4(112) = &H51515151: Te4(113) = &HA3A3A3A3: Te4(114) = &H40404040: Te4(115) = &H8F8F8F8F: Te4(116) = &H92929292: Te4(117) = &H9D9D9D9D: Te4(118) = &H38383838: Te4(119) = &HF5F5F5F5
    Te4(120) = &HBCBCBCBC: Te4(121) = &HB6B6B6B6: Te4(122) = &HDADADADA: Te4(123) = &H21212121: Te4(124) = &H10101010: Te4(125) = &HFFFFFFFF: Te4(126) = &HF3F3F3F3: Te4(127) = &HD2D2D2D2
    Te4(128) = &HCDCDCDCD: Te4(129) = &HC0C0C0C:  Te4(130) = &H13131313: Te4(131) = &HECECECEC: Te4(132) = &H5F5F5F5F: Te4(133) = &H97979797: Te4(134) = &H44444444: Te4(135) = &H17171717
    Te4(136) = &HC4C4C4C4: Te4(137) = &HA7A7A7A7: Te4(138) = &H7E7E7E7E: Te4(139) = &H3D3D3D3D: Te4(140) = &H64646464: Te4(141) = &H5D5D5D5D: Te4(142) = &H19191919: Te4(143) = &H73737373
    Te4(144) = &H60606060: Te4(145) = &H81818181: Te4(146) = &H4F4F4F4F: Te4(147) = &HDCDCDCDC: Te4(148) = &H22222222: Te4(149) = &H2A2A2A2A: Te4(150) = &H90909090: Te4(151) = &H88888888
    Te4(152) = &H46464646: Te4(153) = &HEEEEEEEE: Te4(154) = &HB8B8B8B8: Te4(155) = &H14141414: Te4(156) = &HDEDEDEDE: Te4(157) = &H5E5E5E5E: Te4(158) = &HB0B0B0B:  Te4(159) = &HDBDBDBDB
    Te4(160) = &HE0E0E0E0: Te4(161) = &H32323232: Te4(162) = &H3A3A3A3A: Te4(163) = &HA0A0A0A:  Te4(164) = &H49494949: Te4(165) = &H6060606:  Te4(166) = &H24242424: Te4(167) = &H5C5C5C5C
    Te4(168) = &HC2C2C2C2: Te4(169) = &HD3D3D3D3: Te4(170) = &HACACACAC: Te4(171) = &H62626262: Te4(172) = &H91919191: Te4(173) = &H95959595: Te4(174) = &HE4E4E4E4: Te4(175) = &H79797979
    Te4(176) = &HE7E7E7E7: Te4(177) = &HC8C8C8C8: Te4(178) = &H37373737: Te4(179) = &H6D6D6D6D: Te4(180) = &H8D8D8D8D: Te4(181) = &HD5D5D5D5: Te4(182) = &H4E4E4E4E: Te4(183) = &HA9A9A9A9
    Te4(184) = &H6C6C6C6C: Te4(185) = &H56565656: Te4(186) = &HF4F4F4F4: Te4(187) = &HEAEAEAEA: Te4(188) = &H65656565: Te4(189) = &H7A7A7A7A: Te4(190) = &HAEAEAEAE: Te4(191) = &H8080808
    Te4(192) = &HBABABABA: Te4(193) = &H78787878: Te4(194) = &H25252525: Te4(195) = &H2E2E2E2E: Te4(196) = &H1C1C1C1C: Te4(197) = &HA6A6A6A6: Te4(198) = &HB4B4B4B4: Te4(199) = &HC6C6C6C6
    Te4(200) = &HE8E8E8E8: Te4(201) = &HDDDDDDDD: Te4(202) = &H74747474: Te4(203) = &H1F1F1F1F: Te4(204) = &H4B4B4B4B: Te4(205) = &HBDBDBDBD: Te4(206) = &H8B8B8B8B: Te4(207) = &H8A8A8A8A
    Te4(208) = &H70707070: Te4(209) = &H3E3E3E3E: Te4(210) = &HB5B5B5B5: Te4(211) = &H66666666: Te4(212) = &H48484848: Te4(213) = &H3030303:  Te4(214) = &HF6F6F6F6: Te4(215) = &HE0E0E0E
    Te4(216) = &H61616161: Te4(217) = &H35353535: Te4(218) = &H57575757: Te4(219) = &HB9B9B9B9: Te4(220) = &H86868686: Te4(221) = &HC1C1C1C1: Te4(222) = &H1D1D1D1D: Te4(223) = &H9E9E9E9E
    Te4(224) = &HE1E1E1E1: Te4(225) = &HF8F8F8F8: Te4(226) = &H98989898: Te4(227) = &H11111111: Te4(228) = &H69696969: Te4(229) = &HD9D9D9D9: Te4(230) = &H8E8E8E8E: Te4(231) = &H94949494
    Te4(232) = &H9B9B9B9B: Te4(233) = &H1E1E1E1E: Te4(234) = &H87878787: Te4(235) = &HE9E9E9E9: Te4(236) = &HCECECECE: Te4(237) = &H55555555: Te4(238) = &H28282828: Te4(239) = &HDFDFDFDF
    Te4(240) = &H8C8C8C8C: Te4(241) = &HA1A1A1A1: Te4(242) = &H89898989: Te4(243) = &HD0D0D0D:  Te4(244) = &HBFBFBFBF: Te4(245) = &HE6E6E6E6: Te4(246) = &H42424242: Te4(247) = &H68686868
    Te4(248) = &H41414141: Te4(249) = &H99999999: Te4(250) = &H2D2D2D2D: Te4(251) = &HF0F0F0F:  Te4(252) = &HB0B0B0B0: Te4(253) = &H54545454: Te4(254) = &HBBBBBBBB: Te4(255) = &H16161616

    Td0(0) = &H50A7F451:   Td0(1) = &H5365417E:   Td0(2) = &HC3A4171A:   Td0(3) = &H965E273A:   Td0(4) = &HCB6BAB3B:   Td0(5) = &HF1459D1F:   Td0(6) = &HAB58FAAC:   Td0(7) = &H9303E34B
    Td0(8) = &H55FA3020:   Td0(9) = &HF66D76AD:   Td0(10) = &H9176CC88:  Td0(11) = &H254C02F5:  Td0(12) = &HFCD7E54F:  Td0(13) = &HD7CB2AC5:  Td0(14) = &H80443526:  Td0(15) = &H8FA362B5
    Td0(16) = &H495AB1DE:  Td0(17) = &H671BBA25:  Td0(18) = &H980EEA45:  Td0(19) = &HE1C0FE5D:  Td0(20) = &H2752FC3:   Td0(21) = &H12F04C81:  Td0(22) = &HA397468D:  Td0(23) = &HC6F9D36B
    Td0(24) = &HE75F8F03:  Td0(25) = &H959C9215:  Td0(26) = &HEB7A6DBF:  Td0(27) = &HDA595295:  Td0(28) = &H2D83BED4:  Td0(29) = &HD3217458:  Td0(30) = &H2969E049:  Td0(31) = &H44C8C98E
    Td0(32) = &H6A89C275:  Td0(33) = &H78798EF4:  Td0(34) = &H6B3E5899:  Td0(35) = &HDD71B927:  Td0(36) = &HB64FE1BE:  Td0(37) = &H17AD88F0:  Td0(38) = &H66AC20C9:  Td0(39) = &HB43ACE7D
    Td0(40) = &H184ADF63:  Td0(41) = &H82311AE5:  Td0(42) = &H60335197:  Td0(43) = &H457F5362:  Td0(44) = &HE07764B1:  Td0(45) = &H84AE6BBB:  Td0(46) = &H1CA081FE:  Td0(47) = &H942B08F9
    Td0(48) = &H58684870:  Td0(49) = &H19FD458F:  Td0(50) = &H876CDE94:  Td0(51) = &HB7F87B52:  Td0(52) = &H23D373AB:  Td0(53) = &HE2024B72:  Td0(54) = &H578F1FE3:  Td0(55) = &H2AAB5566
    Td0(56) = &H728EBB2:   Td0(57) = &H3C2B52F:   Td0(58) = &H9A7BC586:  Td0(59) = &HA50837D3:  Td0(60) = &HF2872830:  Td0(61) = &HB2A5BF23:  Td0(62) = &HBA6A0302:  Td0(63) = &H5C8216ED
    Td0(64) = &H2B1CCF8A:  Td0(65) = &H92B479A7:  Td0(66) = &HF0F207F3:  Td0(67) = &HA1E2694E:  Td0(68) = &HCDF4DA65:  Td0(69) = &HD5BE0506:  Td0(70) = &H1F6234D1:  Td0(71) = &H8AFEA6C4
    Td0(72) = &H9D532E34:  Td0(73) = &HA055F3A2:  Td0(74) = &H32E18A05:  Td0(75) = &H75EBF6A4:  Td0(76) = &H39EC830B:  Td0(77) = &HAAEF6040:  Td0(78) = &H69F715E:   Td0(79) = &H51106EBD
    Td0(80) = &HF98A213E:  Td0(81) = &H3D06DD96:  Td0(82) = &HAE053EDD:  Td0(83) = &H46BDE64D:  Td0(84) = &HB58D5491:  Td0(85) = &H55DC471:   Td0(86) = &H6FD40604:  Td0(87) = &HFF155060
    Td0(88) = &H24FB9819:  Td0(89) = &H97E9BDD6:  Td0(90) = &HCC434089:  Td0(91) = &H779ED967:  Td0(92) = &HBD42E8B0:  Td0(93) = &H888B8907:  Td0(94) = &H385B19E7:  Td0(95) = &HDBEEC879
    Td0(96) = &H470A7CA1:  Td0(97) = &HE90F427C:  Td0(98) = &HC91E84F8:  Td0(99) = &H0&:        Td0(100) = &H83868009: Td0(101) = &H48ED2B32: Td0(102) = &HAC70111E: Td0(103) = &H4E725A6C
    Td0(104) = &HFBFF0EFD: Td0(105) = &H5638850F: Td0(106) = &H1ED5AE3D: Td0(107) = &H27392D36: Td0(108) = &H64D90F0A: Td0(109) = &H21A65C68: Td0(110) = &HD1545B9B: Td0(111) = &H3A2E3624
    Td0(112) = &HB1670A0C: Td0(113) = &HFE75793:  Td0(114) = &HD296EEB4: Td0(115) = &H9E919B1B: Td0(116) = &H4FC5C080: Td0(117) = &HA220DC61: Td0(118) = &H694B775A: Td0(119) = &H161A121C
    Td0(120) = &HABA93E2:  Td0(121) = &HE52AA0C0: Td0(122) = &H43E0223C: Td0(123) = &H1D171B12: Td0(124) = &HB0D090E:  Td0(125) = &HADC78BF2: Td0(126) = &HB9A8B62D: Td0(127) = &HC8A91E14
    Td0(128) = &H8519F157: Td0(129) = &H4C0775AF: Td0(130) = &HBBDD99EE: Td0(131) = &HFD607FA3: Td0(132) = &H9F2601F7: Td0(133) = &HBCF5725C: Td0(134) = &HC53B6644: Td0(135) = &H347EFB5B
    Td0(136) = &H7629438B: Td0(137) = &HDCC623CB: Td0(138) = &H68FCEDB6: Td0(139) = &H63F1E4B8: Td0(140) = &HCADC31D7: Td0(141) = &H10856342: Td0(142) = &H40229713: Td0(143) = &H2011C684
    Td0(144) = &H7D244A85: Td0(145) = &HF83DBBD2: Td0(146) = &H1132F9AE: Td0(147) = &H6DA129C7: Td0(148) = &H4B2F9E1D: Td0(149) = &HF330B2DC: Td0(150) = &HEC52860D: Td0(151) = &HD0E3C177
    Td0(152) = &H6C16B32B: Td0(153) = &H99B970A9: Td0(154) = &HFA489411: Td0(155) = &H2264E947: Td0(156) = &HC48CFCA8: Td0(157) = &H1A3FF0A0: Td0(158) = &HD82C7D56: Td0(159) = &HEF903322
    Td0(160) = &HC74E4987: Td0(161) = &HC1D138D9: Td0(162) = &HFEA2CA8C: Td0(163) = &H360BD498: Td0(164) = &HCF81F5A6: Td0(165) = &H28DE7AA5: Td0(166) = &H268EB7DA: Td0(167) = &HA4BFAD3F
    Td0(168) = &HE49D3A2C: Td0(169) = &HD927850:  Td0(170) = &H9BCC5F6A: Td0(171) = &H62467E54: Td0(172) = &HC2138DF6: Td0(173) = &HE8B8D890: Td0(174) = &H5EF7392E: Td0(175) = &HF5AFC382
    Td0(176) = &HBE805D9F: Td0(177) = &H7C93D069: Td0(178) = &HA92DD56F: Td0(179) = &HB31225CF: Td0(180) = &H3B99ACC8: Td0(181) = &HA77D1810: Td0(182) = &H6E639CE8: Td0(183) = &H7BBB3BDB
    Td0(184) = &H97826CD:  Td0(185) = &HF418596E: Td0(186) = &H1B79AEC:  Td0(187) = &HA89A4F83: Td0(188) = &H656E95E6: Td0(189) = &H7EE6FFAA: Td0(190) = &H8CFBC21:  Td0(191) = &HE6E815EF
    Td0(192) = &HD99BE7BA: Td0(193) = &HCE366F4A: Td0(194) = &HD4099FEA: Td0(195) = &HD67CB029: Td0(196) = &HAFB2A431: Td0(197) = &H31233F2A: Td0(198) = &H3094A5C6: Td0(199) = &HC066A235
    Td0(200) = &H37BC4E74: Td0(201) = &HA6CA82FC: Td0(202) = &HB0D090E0: Td0(203) = &H15D8A733: Td0(204) = &H4A9804F1: Td0(205) = &HF7DAEC41: Td0(206) = &HE50CD7F:  Td0(207) = &H2FF69117
    Td0(208) = &H8DD64D76: Td0(209) = &H4DB0EF43: Td0(210) = &H544DAACC: Td0(211) = &HDF0496E4: Td0(212) = &HE3B5D19E: Td0(213) = &H1B886A4C: Td0(214) = &HB81F2CC1: Td0(215) = &H7F516546
    Td0(216) = &H4EA5E9D:  Td0(217) = &H5D358C01: Td0(218) = &H737487FA: Td0(219) = &H2E410BFB: Td0(220) = &H5A1D67B3: Td0(221) = &H52D2DB92: Td0(222) = &H335610E9: Td0(223) = &H1347D66D
    Td0(224) = &H8C61D79A: Td0(225) = &H7A0CA137: Td0(226) = &H8E14F859: Td0(227) = &H893C13EB: Td0(228) = &HEE27A9CE: Td0(229) = &H35C961B7: Td0(230) = &HEDE51CE1: Td0(231) = &H3CB1477A
    Td0(232) = &H59DFD29C: Td0(233) = &H3F73F255: Td0(234) = &H79CE1418: Td0(235) = &HBF37C773: Td0(236) = &HEACDF753: Td0(237) = &H5BAAFD5F: Td0(238) = &H146F3DDF: Td0(239) = &H86DB4478
    Td0(240) = &H81F3AFCA: Td0(241) = &H3EC468B9: Td0(242) = &H2C342438: Td0(243) = &H5F40A3C2: Td0(244) = &H72C31D16: Td0(245) = &HC25E2BC:  Td0(246) = &H8B493C28: Td0(247) = &H41950DFF
    Td0(248) = &H7101A839: Td0(249) = &HDEB30C08: Td0(250) = &H9CE4B4D8: Td0(251) = &H90C15664: Td0(252) = &H6184CB7B: Td0(253) = &H70B632D5: Td0(254) = &H745C6C48: Td0(255) = &H4257B8D0

    Td1(0) = &HA7F45150:   Td1(1) = &H65417E53:   Td1(2) = &HA4171AC3:   Td1(3) = &H5E273A96:   Td1(4) = &H6BAB3BCB:   Td1(5) = &H459D1FF1:   Td1(6) = &H58FAACAB:   Td1(7) = &H3E34B93
    Td1(8) = &HFA302055:   Td1(9) = &H6D76ADF6:   Td1(10) = &H76CC8891:  Td1(11) = &H4C02F525:  Td1(12) = &HD7E54FFC:  Td1(13) = &HCB2AC5D7:  Td1(14) = &H44352680:  Td1(15) = &HA362B58F
    Td1(16) = &H5AB1DE49:  Td1(17) = &H1BBA2567:  Td1(18) = &HEEA4598:   Td1(19) = &HC0FE5DE1:  Td1(20) = &H752FC302:  Td1(21) = &HF04C8112:  Td1(22) = &H97468DA3:  Td1(23) = &HF9D36BC6
    Td1(24) = &H5F8F03E7:  Td1(25) = &H9C921595:  Td1(26) = &H7A6DBFEB:  Td1(27) = &H595295DA:  Td1(28) = &H83BED42D:  Td1(29) = &H217458D3:  Td1(30) = &H69E04929:  Td1(31) = &HC8C98E44
    Td1(32) = &H89C2756A:  Td1(33) = &H798EF478:  Td1(34) = &H3E58996B:  Td1(35) = &H71B927DD:  Td1(36) = &H4FE1BEB6:  Td1(37) = &HAD88F017:  Td1(38) = &HAC20C966:  Td1(39) = &H3ACE7DB4
    Td1(40) = &H4ADF6318:  Td1(41) = &H311AE582:  Td1(42) = &H33519760:  Td1(43) = &H7F536245:  Td1(44) = &H7764B1E0:  Td1(45) = &HAE6BBB84:  Td1(46) = &HA081FE1C:  Td1(47) = &H2B08F994
    Td1(48) = &H68487058:  Td1(49) = &HFD458F19:  Td1(50) = &H6CDE9487:  Td1(51) = &HF87B52B7:  Td1(52) = &HD373AB23:  Td1(53) = &H24B72E2:   Td1(54) = &H8F1FE357:  Td1(55) = &HAB55662A
    Td1(56) = &H28EBB207:  Td1(57) = &HC2B52F03:  Td1(58) = &H7BC5869A:  Td1(59) = &H837D3A5:   Td1(60) = &H872830F2:  Td1(61) = &HA5BF23B2:  Td1(62) = &H6A0302BA:  Td1(63) = &H8216ED5C
    Td1(64) = &H1CCF8A2B:  Td1(65) = &HB479A792:  Td1(66) = &HF207F3F0:  Td1(67) = &HE2694EA1:  Td1(68) = &HF4DA65CD:  Td1(69) = &HBE0506D5:  Td1(70) = &H6234D11F:  Td1(71) = &HFEA6C48A
    Td1(72) = &H532E349D:  Td1(73) = &H55F3A2A0:  Td1(74) = &HE18A0532:  Td1(75) = &HEBF6A475:  Td1(76) = &HEC830B39:  Td1(77) = &HEF6040AA:  Td1(78) = &H9F715E06:  Td1(79) = &H106EBD51
    Td1(80) = &H8A213EF9:  Td1(81) = &H6DD963D:   Td1(82) = &H53EDDAE:   Td1(83) = &HBDE64D46:  Td1(84) = &H8D5491B5:  Td1(85) = &H5DC47105:  Td1(86) = &HD406046F:  Td1(87) = &H155060FF
    Td1(88) = &HFB981924:  Td1(89) = &HE9BDD697:  Td1(90) = &H434089CC:  Td1(91) = &H9ED96777:  Td1(92) = &H42E8B0BD:  Td1(93) = &H8B890788:  Td1(94) = &H5B19E738:  Td1(95) = &HEEC879DB
    Td1(96) = &HA7CA147:   Td1(97) = &HF427CE9:   Td1(98) = &H1E84F8C9:  Td1(99) = &H0&:        Td1(100) = &H86800983: Td1(101) = &HED2B3248: Td1(102) = &H70111EAC: Td1(103) = &H725A6C4E
    Td1(104) = &HFF0EFDFB: Td1(105) = &H38850F56: Td1(106) = &HD5AE3D1E: Td1(107) = &H392D3627: Td1(108) = &HD90F0A64: Td1(109) = &HA65C6821: Td1(110) = &H545B9BD1: Td1(111) = &H2E36243A
    Td1(112) = &H670A0CB1: Td1(113) = &HE757930F: Td1(114) = &H96EEB4D2: Td1(115) = &H919B1B9E: Td1(116) = &HC5C0804F: Td1(117) = &H20DC61A2: Td1(118) = &H4B775A69: Td1(119) = &H1A121C16
    Td1(120) = &HBA93E20A: Td1(121) = &H2AA0C0E5: Td1(122) = &HE0223C43: Td1(123) = &H171B121D: Td1(124) = &HD090E0B:  Td1(125) = &HC78BF2AD: Td1(126) = &HA8B62DB9: Td1(127) = &HA91E14C8
    Td1(128) = &H19F15785: Td1(129) = &H775AF4C:  Td1(130) = &HDD99EEBB: Td1(131) = &H607FA3FD: Td1(132) = &H2601F79F: Td1(133) = &HF5725CBC: Td1(134) = &H3B6644C5: Td1(135) = &H7EFB5B34
    Td1(136) = &H29438B76: Td1(137) = &HC623CBDC: Td1(138) = &HFCEDB668: Td1(139) = &HF1E4B863: Td1(140) = &HDC31D7CA: Td1(141) = &H85634210: Td1(142) = &H22971340: Td1(143) = &H11C68420
    Td1(144) = &H244A857D: Td1(145) = &H3DBBD2F8: Td1(146) = &H32F9AE11: Td1(147) = &HA129C76D: Td1(148) = &H2F9E1D4B: Td1(149) = &H30B2DCF3: Td1(150) = &H52860DEC: Td1(151) = &HE3C177D0
    Td1(152) = &H16B32B6C: Td1(153) = &HB970A999: Td1(154) = &H489411FA: Td1(155) = &H64E94722: Td1(156) = &H8CFCA8C4: Td1(157) = &H3FF0A01A: Td1(158) = &H2C7D56D8: Td1(159) = &H903322EF
    Td1(160) = &H4E4987C7: Td1(161) = &HD138D9C1: Td1(162) = &HA2CA8CFE: Td1(163) = &HBD49836:  Td1(164) = &H81F5A6CF: Td1(165) = &HDE7AA528: Td1(166) = &H8EB7DA26: Td1(167) = &HBFAD3FA4
    Td1(168) = &H9D3A2CE4: Td1(169) = &H9278500D: Td1(170) = &HCC5F6A9B: Td1(171) = &H467E5462: Td1(172) = &H138DF6C2: Td1(173) = &HB8D890E8: Td1(174) = &HF7392E5E: Td1(175) = &HAFC382F5
    Td1(176) = &H805D9FBE: Td1(177) = &H93D0697C: Td1(178) = &H2DD56FA9: Td1(179) = &H1225CFB3: Td1(180) = &H99ACC83B: Td1(181) = &H7D1810A7: Td1(182) = &H639CE86E: Td1(183) = &HBB3BDB7B
    Td1(184) = &H7826CD09: Td1(185) = &H18596EF4: Td1(186) = &HB79AEC01: Td1(187) = &H9A4F83A8: Td1(188) = &H6E95E665: Td1(189) = &HE6FFAA7E: Td1(190) = &HCFBC2108: Td1(191) = &HE815EFE6
    Td1(192) = &H9BE7BAD9: Td1(193) = &H366F4ACE: Td1(194) = &H99FEAD4:  Td1(195) = &H7CB029D6: Td1(196) = &HB2A431AF: Td1(197) = &H233F2A31: Td1(198) = &H94A5C630: Td1(199) = &H66A235C0
    Td1(200) = &HBC4E7437: Td1(201) = &HCA82FCA6: Td1(202) = &HD090E0B0: Td1(203) = &HD8A73315: Td1(204) = &H9804F14A: Td1(205) = &HDAEC41F7: Td1(206) = &H50CD7F0E: Td1(207) = &HF691172F
    Td1(208) = &HD64D768D: Td1(209) = &HB0EF434D: Td1(210) = &H4DAACC54: Td1(211) = &H496E4DF:  Td1(212) = &HB5D19EE3: Td1(213) = &H886A4C1B: Td1(214) = &H1F2CC1B8: Td1(215) = &H5165467F
    Td1(216) = &HEA5E9D04: Td1(217) = &H358C015D: Td1(218) = &H7487FA73: Td1(219) = &H410BFB2E: Td1(220) = &H1D67B35A: Td1(221) = &HD2DB9252: Td1(222) = &H5610E933: Td1(223) = &H47D66D13
    Td1(224) = &H61D79A8C: Td1(225) = &HCA1377A:  Td1(226) = &H14F8598E: Td1(227) = &H3C13EB89: Td1(228) = &H27A9CEEE: Td1(229) = &HC961B735: Td1(230) = &HE51CE1ED: Td1(231) = &HB1477A3C
    Td1(232) = &HDFD29C59: Td1(233) = &H73F2553F: Td1(234) = &HCE141879: Td1(235) = &H37C773BF: Td1(236) = &HCDF753EA: Td1(237) = &HAAFD5F5B: Td1(238) = &H6F3DDF14: Td1(239) = &HDB447886
    Td1(240) = &HF3AFCA81: Td1(241) = &HC468B93E: Td1(242) = &H3424382C: Td1(243) = &H40A3C25F: Td1(244) = &HC31D1672: Td1(245) = &H25E2BC0C: Td1(246) = &H493C288B: Td1(247) = &H950DFF41
    Td1(248) = &H1A83971:  Td1(249) = &HB30C08DE: Td1(250) = &HE4B4D89C: Td1(251) = &HC1566490: Td1(252) = &H84CB7B61: Td1(253) = &HB632D570: Td1(254) = &H5C6C4874: Td1(255) = &H57B8D042

    Td2(0) = &HF45150A7:   Td2(1) = &H417E5365:   Td2(2) = &H171AC3A4:   Td2(3) = &H273A965E:   Td2(4) = &HAB3BCB6B:   Td2(5) = &H9D1FF145:   Td2(6) = &HFAACAB58:   Td2(7) = &HE34B9303
    Td2(8) = &H302055FA:   Td2(9) = &H76ADF66D:   Td2(10) = &HCC889176:  Td2(11) = &H2F5254C:   Td2(12) = &HE54FFCD7:  Td2(13) = &H2AC5D7CB:  Td2(14) = &H35268044:  Td2(15) = &H62B58FA3
    Td2(16) = &HB1DE495A:  Td2(17) = &HBA25671B:  Td2(18) = &HEA45980E:  Td2(19) = &HFE5DE1C0:  Td2(20) = &H2FC30275:  Td2(21) = &H4C8112F0:  Td2(22) = &H468DA397:  Td2(23) = &HD36BC6F9
    Td2(24) = &H8F03E75F:  Td2(25) = &H9215959C:  Td2(26) = &H6DBFEB7A:  Td2(27) = &H5295DA59:  Td2(28) = &HBED42D83:  Td2(29) = &H7458D321:  Td2(30) = &HE0492969:  Td2(31) = &HC98E44C8
    Td2(32) = &HC2756A89:  Td2(33) = &H8EF47879:  Td2(34) = &H58996B3E:  Td2(35) = &HB927DD71:  Td2(36) = &HE1BEB64F:  Td2(37) = &H88F017AD:  Td2(38) = &H20C966AC:  Td2(39) = &HCE7DB43A
    Td2(40) = &HDF63184A:  Td2(41) = &H1AE58231:  Td2(42) = &H51976033:  Td2(43) = &H5362457F:  Td2(44) = &H64B1E077:  Td2(45) = &H6BBB84AE:  Td2(46) = &H81FE1CA0:  Td2(47) = &H8F9942B
    Td2(48) = &H48705868:  Td2(49) = &H458F19FD:  Td2(50) = &HDE94876C:  Td2(51) = &H7B52B7F8:  Td2(52) = &H73AB23D3:  Td2(53) = &H4B72E202:  Td2(54) = &H1FE3578F:  Td2(55) = &H55662AAB
    Td2(56) = &HEBB20728:  Td2(57) = &HB52F03C2:  Td2(58) = &HC5869A7B:  Td2(59) = &H37D3A508:  Td2(60) = &H2830F287:  Td2(61) = &HBF23B2A5:  Td2(62) = &H302BA6A:   Td2(63) = &H16ED5C82
    Td2(64) = &HCF8A2B1C:  Td2(65) = &H79A792B4:  Td2(66) = &H7F3F0F2:   Td2(67) = &H694EA1E2:  Td2(68) = &HDA65CDF4:  Td2(69) = &H506D5BE:   Td2(70) = &H34D11F62:  Td2(71) = &HA6C48AFE
    Td2(72) = &H2E349D53:  Td2(73) = &HF3A2A055:  Td2(74) = &H8A0532E1:  Td2(75) = &HF6A475EB:  Td2(76) = &H830B39EC:  Td2(77) = &H6040AAEF:  Td2(78) = &H715E069F:  Td2(79) = &H6EBD5110
    Td2(80) = &H213EF98A:  Td2(81) = &HDD963D06:  Td2(82) = &H3EDDAE05:  Td2(83) = &HE64D46BD:  Td2(84) = &H5491B58D:  Td2(85) = &HC471055D:  Td2(86) = &H6046FD4:   Td2(87) = &H5060FF15
    Td2(88) = &H981924FB:  Td2(89) = &HBDD697E9:  Td2(90) = &H4089CC43:  Td2(91) = &HD967779E:  Td2(92) = &HE8B0BD42:  Td2(93) = &H8907888B:  Td2(94) = &H19E7385B:  Td2(95) = &HC879DBEE
    Td2(96) = &H7CA1470A:  Td2(97) = &H427CE90F:  Td2(98) = &H84F8C91E:  Td2(99) = &H0&:        Td2(100) = &H80098386: Td2(101) = &H2B3248ED: Td2(102) = &H111EAC70: Td2(103) = &H5A6C4E72
    Td2(104) = &HEFDFBFF:  Td2(105) = &H850F5638: Td2(106) = &HAE3D1ED5: Td2(107) = &H2D362739: Td2(108) = &HF0A64D9:  Td2(109) = &H5C6821A6: Td2(110) = &H5B9BD154: Td2(111) = &H36243A2E
    Td2(112) = &HA0CB167:  Td2(113) = &H57930FE7: Td2(114) = &HEEB4D296: Td2(115) = &H9B1B9E91: Td2(116) = &HC0804FC5: Td2(117) = &HDC61A220: Td2(118) = &H775A694B: Td2(119) = &H121C161A
    Td2(120) = &H93E20ABA: Td2(121) = &HA0C0E52A: Td2(122) = &H223C43E0: Td2(123) = &H1B121D17: Td2(124) = &H90E0B0D:  Td2(125) = &H8BF2ADC7: Td2(126) = &HB62DB9A8: Td2(127) = &H1E14C8A9
    Td2(128) = &HF1578519: Td2(129) = &H75AF4C07: Td2(130) = &H99EEBBDD: Td2(131) = &H7FA3FD60: Td2(132) = &H1F79F26:  Td2(133) = &H725CBCF5: Td2(134) = &H6644C53B: Td2(135) = &HFB5B347E
    Td2(136) = &H438B7629: Td2(137) = &H23CBDCC6: Td2(138) = &HEDB668FC: Td2(139) = &HE4B863F1: Td2(140) = &H31D7CADC: Td2(141) = &H63421085: Td2(142) = &H97134022: Td2(143) = &HC6842011
    Td2(144) = &H4A857D24: Td2(145) = &HBBD2F83D: Td2(146) = &HF9AE1132: Td2(147) = &H29C76DA1: Td2(148) = &H9E1D4B2F: Td2(149) = &HB2DCF330: Td2(150) = &H860DEC52: Td2(151) = &HC177D0E3
    Td2(152) = &HB32B6C16: Td2(153) = &H70A999B9: Td2(154) = &H9411FA48: Td2(155) = &HE9472264: Td2(156) = &HFCA8C48C: Td2(157) = &HF0A01A3F: Td2(158) = &H7D56D82C: Td2(159) = &H3322EF90
    Td2(160) = &H4987C74E: Td2(161) = &H38D9C1D1: Td2(162) = &HCA8CFEA2: Td2(163) = &HD498360B: Td2(164) = &HF5A6CF81: Td2(165) = &H7AA528DE: Td2(166) = &HB7DA268E: Td2(167) = &HAD3FA4BF
    Td2(168) = &H3A2CE49D: Td2(169) = &H78500D92: Td2(170) = &H5F6A9BCC: Td2(171) = &H7E546246: Td2(172) = &H8DF6C213: Td2(173) = &HD890E8B8: Td2(174) = &H392E5EF7: Td2(175) = &HC382F5AF
    Td2(176) = &H5D9FBE80: Td2(177) = &HD0697C93: Td2(178) = &HD56FA92D: Td2(179) = &H25CFB312: Td2(180) = &HACC83B99: Td2(181) = &H1810A77D: Td2(182) = &H9CE86E63: Td2(183) = &H3BDB7BBB
    Td2(184) = &H26CD0978: Td2(185) = &H596EF418: Td2(186) = &H9AEC01B7: Td2(187) = &H4F83A89A: Td2(188) = &H95E6656E: Td2(189) = &HFFAA7EE6: Td2(190) = &HBC2108CF: Td2(191) = &H15EFE6E8
    Td2(192) = &HE7BAD99B: Td2(193) = &H6F4ACE36: Td2(194) = &H9FEAD409: Td2(195) = &HB029D67C: Td2(196) = &HA431AFB2: Td2(197) = &H3F2A3123: Td2(198) = &HA5C63094: Td2(199) = &HA235C066
    Td2(200) = &H4E7437BC: Td2(201) = &H82FCA6CA: Td2(202) = &H90E0B0D0: Td2(203) = &HA73315D8: Td2(204) = &H4F14A98:  Td2(205) = &HEC41F7DA: Td2(206) = &HCD7F0E50: Td2(207) = &H91172FF6
    Td2(208) = &H4D768DD6: Td2(209) = &HEF434DB0: Td2(210) = &HAACC544D: Td2(211) = &H96E4DF04: Td2(212) = &HD19EE3B5: Td2(213) = &H6A4C1B88: Td2(214) = &H2CC1B81F: Td2(215) = &H65467F51
    Td2(216) = &H5E9D04EA: Td2(217) = &H8C015D35: Td2(218) = &H87FA7374: Td2(219) = &HBFB2E41:  Td2(220) = &H67B35A1D: Td2(221) = &HDB9252D2: Td2(222) = &H10E93356: Td2(223) = &HD66D1347
    Td2(224) = &HD79A8C61: Td2(225) = &HA1377A0C: Td2(226) = &HF8598E14: Td2(227) = &H13EB893C: Td2(228) = &HA9CEEE27: Td2(229) = &H61B735C9: Td2(230) = &H1CE1EDE5: Td2(231) = &H477A3CB1
    Td2(232) = &HD29C59DF: Td2(233) = &HF2553F73: Td2(234) = &H141879CE: Td2(235) = &HC773BF37: Td2(236) = &HF753EACD: Td2(237) = &HFD5F5BAA: Td2(238) = &H3DDF146F: Td2(239) = &H447886DB
    Td2(240) = &HAFCA81F3: Td2(241) = &H68B93EC4: Td2(242) = &H24382C34: Td2(243) = &HA3C25F40: Td2(244) = &H1D1672C3: Td2(245) = &HE2BC0C25: Td2(246) = &H3C288B49: Td2(247) = &HDFF4195
    Td2(248) = &HA8397101: Td2(249) = &HC08DEB3:  Td2(250) = &HB4D89CE4: Td2(251) = &H566490C1: Td2(252) = &HCB7B6184: Td2(253) = &H32D570B6: Td2(254) = &H6C48745C: Td2(255) = &HB8D04257

    Td3(0) = &H5150A7F4:   Td3(1) = &H7E536541:   Td3(2) = &H1AC3A417:   Td3(3) = &H3A965E27:   Td3(4) = &H3BCB6BAB:   Td3(5) = &H1FF1459D:   Td3(6) = &HACAB58FA:   Td3(7) = &H4B9303E3
    Td3(8) = &H2055FA30:   Td3(9) = &HADF66D76:   Td3(10) = &H889176CC:  Td3(11) = &HF5254C02:  Td3(12) = &H4FFCD7E5:  Td3(13) = &HC5D7CB2A:  Td3(14) = &H26804435:  Td3(15) = &HB58FA362
    Td3(16) = &HDE495AB1:  Td3(17) = &H25671BBA:  Td3(18) = &H45980EEA:  Td3(19) = &H5DE1C0FE:  Td3(20) = &HC302752F:  Td3(21) = &H8112F04C:  Td3(22) = &H8DA39746:  Td3(23) = &H6BC6F9D3
    Td3(24) = &H3E75F8F:   Td3(25) = &H15959C92:  Td3(26) = &HBFEB7A6D:  Td3(27) = &H95DA5952:  Td3(28) = &HD42D83BE:  Td3(29) = &H58D32174:  Td3(30) = &H492969E0:  Td3(31) = &H8E44C8C9
    Td3(32) = &H756A89C2:  Td3(33) = &HF478798E:  Td3(34) = &H996B3E58:  Td3(35) = &H27DD71B9:  Td3(36) = &HBEB64FE1:  Td3(37) = &HF017AD88:  Td3(38) = &HC966AC20:  Td3(39) = &H7DB43ACE
    Td3(40) = &H63184ADF:  Td3(41) = &HE582311A:  Td3(42) = &H97603351:  Td3(43) = &H62457F53:  Td3(44) = &HB1E07764:  Td3(45) = &HBB84AE6B:  Td3(46) = &HFE1CA081:  Td3(47) = &HF9942B08
    Td3(48) = &H70586848:  Td3(49) = &H8F19FD45:  Td3(50) = &H94876CDE:  Td3(51) = &H52B7F87B:  Td3(52) = &HAB23D373:  Td3(53) = &H72E2024B:  Td3(54) = &HE3578F1F:  Td3(55) = &H662AAB55
    Td3(56) = &HB20728EB:  Td3(57) = &H2F03C2B5:  Td3(58) = &H869A7BC5:  Td3(59) = &HD3A50837:  Td3(60) = &H30F28728:  Td3(61) = &H23B2A5BF:  Td3(62) = &H2BA6A03:   Td3(63) = &HED5C8216
    Td3(64) = &H8A2B1CCF:  Td3(65) = &HA792B479:  Td3(66) = &HF3F0F207:  Td3(67) = &H4EA1E269:  Td3(68) = &H65CDF4DA:  Td3(69) = &H6D5BE05:   Td3(70) = &HD11F6234:  Td3(71) = &HC48AFEA6
    Td3(72) = &H349D532E:  Td3(73) = &HA2A055F3:  Td3(74) = &H532E18A:   Td3(75) = &HA475EBF6:  Td3(76) = &HB39EC83:   Td3(77) = &H40AAEF60:  Td3(78) = &H5E069F71:  Td3(79) = &HBD51106E
    Td3(80) = &H3EF98A21:  Td3(81) = &H963D06DD:  Td3(82) = &HDDAE053E:  Td3(83) = &H4D46BDE6:  Td3(84) = &H91B58D54:  Td3(85) = &H71055DC4:  Td3(86) = &H46FD406:   Td3(87) = &H60FF1550
    Td3(88) = &H1924FB98:  Td3(89) = &HD697E9BD:  Td3(90) = &H89CC4340:  Td3(91) = &H67779ED9:  Td3(92) = &HB0BD42E8:  Td3(93) = &H7888B89:   Td3(94) = &HE7385B19:  Td3(95) = &H79DBEEC8
    Td3(96) = &HA1470A7C:  Td3(97) = &H7CE90F42:  Td3(98) = &HF8C91E84:  Td3(99) = &H0&:        Td3(100) = &H9838680:  Td3(101) = &H3248ED2B: Td3(102) = &H1EAC7011: Td3(103) = &H6C4E725A
    Td3(104) = &HFDFBFF0E: Td3(105) = &HF563885:  Td3(106) = &H3D1ED5AE: Td3(107) = &H3627392D: Td3(108) = &HA64D90F:  Td3(109) = &H6821A65C: Td3(110) = &H9BD1545B: Td3(111) = &H243A2E36
    Td3(112) = &HCB1670A:  Td3(113) = &H930FE757: Td3(114) = &HB4D296EE: Td3(115) = &H1B9E919B: Td3(116) = &H804FC5C0: Td3(117) = &H61A220DC: Td3(118) = &H5A694B77: Td3(119) = &H1C161A12
    Td3(120) = &HE20ABA93: Td3(121) = &HC0E52AA0: Td3(122) = &H3C43E022: Td3(123) = &H121D171B: Td3(124) = &HE0B0D09:  Td3(125) = &HF2ADC78B: Td3(126) = &H2DB9A8B6: Td3(127) = &H14C8A91E
    Td3(128) = &H578519F1: Td3(129) = &HAF4C0775: Td3(130) = &HEEBBDD99: Td3(131) = &HA3FD607F: Td3(132) = &HF79F2601: Td3(133) = &H5CBCF572: Td3(134) = &H44C53B66: Td3(135) = &H5B347EFB
    Td3(136) = &H8B762943: Td3(137) = &HCBDCC623: Td3(138) = &HB668FCED: Td3(139) = &HB863F1E4: Td3(140) = &HD7CADC31: Td3(141) = &H42108563: Td3(142) = &H13402297: Td3(143) = &H842011C6
    Td3(144) = &H857D244A: Td3(145) = &HD2F83DBB: Td3(146) = &HAE1132F9: Td3(147) = &HC76DA129: Td3(148) = &H1D4B2F9E: Td3(149) = &HDCF330B2: Td3(150) = &HDEC5286:  Td3(151) = &H77D0E3C1
    Td3(152) = &H2B6C16B3: Td3(153) = &HA999B970: Td3(154) = &H11FA4894: Td3(155) = &H472264E9: Td3(156) = &HA8C48CFC: Td3(157) = &HA01A3FF0: Td3(158) = &H56D82C7D: Td3(159) = &H22EF9033
    Td3(160) = &H87C74E49: Td3(161) = &HD9C1D138: Td3(162) = &H8CFEA2CA: Td3(163) = &H98360BD4: Td3(164) = &HA6CF81F5: Td3(165) = &HA528DE7A: Td3(166) = &HDA268EB7: Td3(167) = &H3FA4BFAD
    Td3(168) = &H2CE49D3A: Td3(169) = &H500D9278: Td3(170) = &H6A9BCC5F: Td3(171) = &H5462467E: Td3(172) = &HF6C2138D: Td3(173) = &H90E8B8D8: Td3(174) = &H2E5EF739: Td3(175) = &H82F5AFC3
    Td3(176) = &H9FBE805D: Td3(177) = &H697C93D0: Td3(178) = &H6FA92DD5: Td3(179) = &HCFB31225: Td3(180) = &HC83B99AC: Td3(181) = &H10A77D18: Td3(182) = &HE86E639C: Td3(183) = &HDB7BBB3B
    Td3(184) = &HCD097826: Td3(185) = &H6EF41859: Td3(186) = &HEC01B79A: Td3(187) = &H83A89A4F: Td3(188) = &HE6656E95: Td3(189) = &HAA7EE6FF: Td3(190) = &H2108CFBC: Td3(191) = &HEFE6E815
    Td3(192) = &HBAD99BE7: Td3(193) = &H4ACE366F: Td3(194) = &HEAD4099F: Td3(195) = &H29D67CB0: Td3(196) = &H31AFB2A4: Td3(197) = &H2A31233F: Td3(198) = &HC63094A5: Td3(199) = &H35C066A2
    Td3(200) = &H7437BC4E: Td3(201) = &HFCA6CA82: Td3(202) = &HE0B0D090: Td3(203) = &H3315D8A7: Td3(204) = &HF14A9804: Td3(205) = &H41F7DAEC: Td3(206) = &H7F0E50CD: Td3(207) = &H172FF691
    Td3(208) = &H768DD64D: Td3(209) = &H434DB0EF: Td3(210) = &HCC544DAA: Td3(211) = &HE4DF0496: Td3(212) = &H9EE3B5D1: Td3(213) = &H4C1B886A: Td3(214) = &HC1B81F2C: Td3(215) = &H467F5165
    Td3(216) = &H9D04EA5E: Td3(217) = &H15D358C:  Td3(218) = &HFA737487: Td3(219) = &HFB2E410B: Td3(220) = &HB35A1D67: Td3(221) = &H9252D2DB: Td3(222) = &HE9335610: Td3(223) = &H6D1347D6
    Td3(224) = &H9A8C61D7: Td3(225) = &H377A0CA1: Td3(226) = &H598E14F8: Td3(227) = &HEB893C13: Td3(228) = &HCEEE27A9: Td3(229) = &HB735C961: Td3(230) = &HE1EDE51C: Td3(231) = &H7A3CB147
    Td3(232) = &H9C59DFD2: Td3(233) = &H553F73F2: Td3(234) = &H1879CE14: Td3(235) = &H73BF37C7: Td3(236) = &H53EACDF7: Td3(237) = &H5F5BAAFD: Td3(238) = &HDF146F3D: Td3(239) = &H7886DB44
    Td3(240) = &HCA81F3AF: Td3(241) = &HB93EC468: Td3(242) = &H382C3424: Td3(243) = &HC25F40A3: Td3(244) = &H1672C31D: Td3(245) = &HBC0C25E2: Td3(246) = &H288B493C: Td3(247) = &HFF41950D
    Td3(248) = &H397101A8: Td3(249) = &H8DEB30C:  Td3(250) = &HD89CE4B4: Td3(251) = &H6490C156: Td3(252) = &H7B6184CB: Td3(253) = &HD570B632: Td3(254) = &H48745C6C: Td3(255) = &HD04257B8

    Td4(0) = &H52525252:   Td4(1) = &H9090909:    Td4(2) = &H6A6A6A6A:   Td4(3) = &HD5D5D5D5:   Td4(4) = &H30303030:   Td4(5) = &H36363636:   Td4(6) = &HA5A5A5A5:   Td4(7) = &H38383838
    Td4(8) = &HBFBFBFBF:   Td4(9) = &H40404040:   Td4(10) = &HA3A3A3A3:  Td4(11) = &H9E9E9E9E:  Td4(12) = &H81818181:  Td4(13) = &HF3F3F3F3:  Td4(14) = &HD7D7D7D7:  Td4(15) = &HFBFBFBFB
    Td4(16) = &H7C7C7C7C:  Td4(17) = &HE3E3E3E3:  Td4(18) = &H39393939:  Td4(19) = &H82828282:  Td4(20) = &H9B9B9B9B:  Td4(21) = &H2F2F2F2F:  Td4(22) = &HFFFFFFFF:  Td4(23) = &H87878787
    Td4(24) = &H34343434:  Td4(25) = &H8E8E8E8E:  Td4(26) = &H43434343:  Td4(27) = &H44444444:  Td4(28) = &HC4C4C4C4:  Td4(29) = &HDEDEDEDE:  Td4(30) = &HE9E9E9E9:  Td4(31) = &HCBCBCBCB
    Td4(32) = &H54545454:  Td4(33) = &H7B7B7B7B:  Td4(34) = &H94949494:  Td4(35) = &H32323232:  Td4(36) = &HA6A6A6A6:  Td4(37) = &HC2C2C2C2:  Td4(38) = &H23232323:  Td4(39) = &H3D3D3D3D
    Td4(40) = &HEEEEEEEE:  Td4(41) = &H4C4C4C4C:  Td4(42) = &H95959595:  Td4(43) = &HB0B0B0B:   Td4(44) = &H42424242:  Td4(45) = &HFAFAFAFA:  Td4(46) = &HC3C3C3C3:  Td4(47) = &H4E4E4E4E
    Td4(48) = &H8080808:   Td4(49) = &H2E2E2E2E:  Td4(50) = &HA1A1A1A1:  Td4(51) = &H66666666:  Td4(52) = &H28282828:  Td4(53) = &HD9D9D9D9:  Td4(54) = &H24242424:  Td4(55) = &HB2B2B2B2
    Td4(56) = &H76767676:  Td4(57) = &H5B5B5B5B:  Td4(58) = &HA2A2A2A2:  Td4(59) = &H49494949:  Td4(60) = &H6D6D6D6D:  Td4(61) = &H8B8B8B8B:  Td4(62) = &HD1D1D1D1:  Td4(63) = &H25252525
    Td4(64) = &H72727272:  Td4(65) = &HF8F8F8F8:  Td4(66) = &HF6F6F6F6:  Td4(67) = &H64646464:  Td4(68) = &H86868686:  Td4(69) = &H68686868:  Td4(70) = &H98989898:  Td4(71) = &H16161616
    Td4(72) = &HD4D4D4D4:  Td4(73) = &HA4A4A4A4:  Td4(74) = &H5C5C5C5C:  Td4(75) = &HCCCCCCCC:  Td4(76) = &H5D5D5D5D:  Td4(77) = &H65656565:  Td4(78) = &HB6B6B6B6:  Td4(79) = &H92929292
    Td4(80) = &H6C6C6C6C:  Td4(81) = &H70707070:  Td4(82) = &H48484848:  Td4(83) = &H50505050:  Td4(84) = &HFDFDFDFD:  Td4(85) = &HEDEDEDED:  Td4(86) = &HB9B9B9B9:  Td4(87) = &HDADADADA
    Td4(88) = &H5E5E5E5E:  Td4(89) = &H15151515:  Td4(90) = &H46464646:  Td4(91) = &H57575757:  Td4(92) = &HA7A7A7A7:  Td4(93) = &H8D8D8D8D:  Td4(94) = &H9D9D9D9D:  Td4(95) = &H84848484
    Td4(96) = &H90909090:  Td4(97) = &HD8D8D8D8:  Td4(98) = &HABABABAB:  Td4(99) = &H0&:        Td4(100) = &H8C8C8C8C: Td4(101) = &HBCBCBCBC: Td4(102) = &HD3D3D3D3: Td4(103) = &HA0A0A0A
    Td4(104) = &HF7F7F7F7: Td4(105) = &HE4E4E4E4: Td4(106) = &H58585858: Td4(107) = &H5050505:  Td4(108) = &HB8B8B8B8: Td4(109) = &HB3B3B3B3: Td4(110) = &H45454545: Td4(111) = &H6060606
    Td4(112) = &HD0D0D0D0: Td4(113) = &H2C2C2C2C: Td4(114) = &H1E1E1E1E: Td4(115) = &H8F8F8F8F: Td4(116) = &HCACACACA: Td4(117) = &H3F3F3F3F: Td4(118) = &HF0F0F0F:  Td4(119) = &H2020202
    Td4(120) = &HC1C1C1C1: Td4(121) = &HAFAFAFAF: Td4(122) = &HBDBDBDBD: Td4(123) = &H3030303:  Td4(124) = &H1010101:  Td4(125) = &H13131313: Td4(126) = &H8A8A8A8A: Td4(127) = &H6B6B6B6B
    Td4(128) = &H3A3A3A3A: Td4(129) = &H91919191: Td4(130) = &H11111111: Td4(131) = &H41414141: Td4(132) = &H4F4F4F4F: Td4(133) = &H67676767: Td4(134) = &HDCDCDCDC: Td4(135) = &HEAEAEAEA
    Td4(136) = &H97979797: Td4(137) = &HF2F2F2F2: Td4(138) = &HCFCFCFCF: Td4(139) = &HCECECECE: Td4(140) = &HF0F0F0F0: Td4(141) = &HB4B4B4B4: Td4(142) = &HE6E6E6E6: Td4(143) = &H73737373
    Td4(144) = &H96969696: Td4(145) = &HACACACAC: Td4(146) = &H74747474: Td4(147) = &H22222222: Td4(148) = &HE7E7E7E7: Td4(149) = &HADADADAD: Td4(150) = &H35353535: Td4(151) = &H85858585
    Td4(152) = &HE2E2E2E2: Td4(153) = &HF9F9F9F9: Td4(154) = &H37373737: Td4(155) = &HE8E8E8E8: Td4(156) = &H1C1C1C1C: Td4(157) = &H75757575: Td4(158) = &HDFDFDFDF: Td4(159) = &H6E6E6E6E
    Td4(160) = &H47474747: Td4(161) = &HF1F1F1F1: Td4(162) = &H1A1A1A1A: Td4(163) = &H71717171: Td4(164) = &H1D1D1D1D: Td4(165) = &H29292929: Td4(166) = &HC5C5C5C5: Td4(167) = &H89898989
    Td4(168) = &H6F6F6F6F: Td4(169) = &HB7B7B7B7: Td4(170) = &H62626262: Td4(171) = &HE0E0E0E:  Td4(172) = &HAAAAAAAA: Td4(173) = &H18181818: Td4(174) = &HBEBEBEBE: Td4(175) = &H1B1B1B1B
    Td4(176) = &HFCFCFCFC: Td4(177) = &H56565656: Td4(178) = &H3E3E3E3E: Td4(179) = &H4B4B4B4B: Td4(180) = &HC6C6C6C6: Td4(181) = &HD2D2D2D2: Td4(182) = &H79797979: Td4(183) = &H20202020
    Td4(184) = &H9A9A9A9A: Td4(185) = &HDBDBDBDB: Td4(186) = &HC0C0C0C0: Td4(187) = &HFEFEFEFE: Td4(188) = &H78787878: Td4(189) = &HCDCDCDCD: Td4(190) = &H5A5A5A5A: Td4(191) = &HF4F4F4F4
    Td4(192) = &H1F1F1F1F: Td4(193) = &HDDDDDDDD: Td4(194) = &HA8A8A8A8: Td4(195) = &H33333333: Td4(196) = &H88888888: Td4(197) = &H7070707:  Td4(198) = &HC7C7C7C7: Td4(199) = &H31313131
    Td4(200) = &HB1B1B1B1: Td4(201) = &H12121212: Td4(202) = &H10101010: Td4(203) = &H59595959: Td4(204) = &H27272727: Td4(205) = &H80808080: Td4(206) = &HECECECEC: Td4(207) = &H5F5F5F5F
    Td4(208) = &H60606060: Td4(209) = &H51515151: Td4(210) = &H7F7F7F7F: Td4(211) = &HA9A9A9A9: Td4(212) = &H19191919: Td4(213) = &HB5B5B5B5: Td4(214) = &H4A4A4A4A: Td4(215) = &HD0D0D0D
    Td4(216) = &H2D2D2D2D: Td4(217) = &HE5E5E5E5: Td4(218) = &H7A7A7A7A: Td4(219) = &H9F9F9F9F: Td4(220) = &H93939393: Td4(221) = &HC9C9C9C9: Td4(222) = &H9C9C9C9C: Td4(223) = &HEFEFEFEF
    Td4(224) = &HA0A0A0A0: Td4(225) = &HE0E0E0E0: Td4(226) = &H3B3B3B3B: Td4(227) = &H4D4D4D4D: Td4(228) = &HAEAEAEAE: Td4(229) = &H2A2A2A2A: Td4(230) = &HF5F5F5F5: Td4(231) = &HB0B0B0B0
    Td4(232) = &HC8C8C8C8: Td4(233) = &HEBEBEBEB: Td4(234) = &HBBBBBBBB: Td4(235) = &H3C3C3C3C: Td4(236) = &H83838383: Td4(237) = &H53535353: Td4(238) = &H99999999: Td4(239) = &H61616161
    Td4(240) = &H17171717: Td4(241) = &H2B2B2B2B: Td4(242) = &H4040404:  Td4(243) = &H7E7E7E7E: Td4(244) = &HBABABABA: Td4(245) = &H77777777: Td4(246) = &HD6D6D6D6: Td4(247) = &H26262626
    Td4(248) = &HE1E1E1E1: Td4(249) = &H69696969: Td4(250) = &H14141414: Td4(251) = &H63636363: Td4(252) = &H55555555: Td4(253) = &H21212121: Td4(254) = &HC0C0C0C:  Td4(255) = &H7D7D7D7D

    'for 128-bit blocks, Rijndael never uses more than 10 rcon values
    rco(0) = &H1&:    rco(1) = &H2&:    rco(2) = &H4&:    rco(3) = &H8&:    rco(4) = &H10&
    rco(5) = &H20&:   rco(6) = &H40&:   rco(7) = &H80&:   rco(8) = &H1B&:   rco(9) = &H36&
#If SUPPORT_LEVEL Then
    rco(10) = &H6C&:  rco(11) = &HD8&:  rco(12) = &HAB&:  rco(13) = &H4D&:  rco(14) = &H9A&
    rco(15) = &H2F&:  rco(16) = &H5E&:  rco(17) = &HBC&:  rco(18) = &H63&:  rco(19) = &HC6&
    rco(20) = &H97&:  rco(21) = &H35&:  rco(22) = &H6A&:  rco(23) = &HD4&:  rco(24) = &HB3&
    rco(25) = &H7D&:  rco(26) = &HFA&:  rco(27) = &HEF&:  rco(28) = &HC5&
#End If
#End If
End Sub


'================================================================================
'================================================================================
'================================================================================
Public Function EncryptString(ByVal ToEncryptString As String, ByVal KeyString As String) As String
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(ToEncryptString) = 0 Then
        EncryptString = ""
    Else
        If Len(KeyString) = 0 Then
            EncryptString = ""
        Else
            KeyBits = 192       '128 192 256
            
            BlockBits = 128
            
            pass = GetPassword(KeyString)

            plaintext = StrConv(ToEncryptString, vbFromUnicode)

            SetCipherKey pass, KeyBits
            ArrayEncrypt plaintext, ciphertext, 0
            EncryptString = HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
        End If
    End If
End Function

Private Function GetPassword(ByVal KeyString As String) As Byte()
    Dim Data() As Byte
    Data = StrConv(KeyString, vbFromUnicode)
    ReDim Preserve Data(31)
    GetPassword = Data
    
End Function

'Returns a String containing Hex values of data(0 ... n-1) in groups of k
Private Function HexDisplay(Data() As Byte, n As Long, K As Long) As String
    Dim i As Long
    Dim j As Long
    Dim C As Long
    Dim data2() As Byte

    If LBound(Data) = 0 Then
        ReDim data2(n * 4 - 1 + ((n - 1) \ K) * 4)
        j = 0
        For i = 0 To n - 1
            If i Mod K = 0 Then
                If i <> 0 Then
                    data2(j) = 32
                    data2(j + 2) = 32
                    j = j + 4
                End If
            End If
            C = Data(i) \ 16&
            If C < 10 Then
                data2(j) = C + 48     ' "0"..."9"
            Else
                data2(j) = C + 55     ' "A"..."F"
            End If
            C = Data(i) And 15&
            If C < 10 Then
                data2(j + 2) = C + 48 ' "0"..."9"
            Else
                data2(j + 2) = C + 55 ' "A"..."F"
            End If
            j = j + 4
        Next i
Debug.Assert j = UBound(data2) + 1
        HexDisplay = data2
    End If

End Function


'Reverse of HexDisplay.  Given a String containing Hex values, convert to byte array data()
'Returns number of bytes n in data(0 ... n-1)
Private Function HexDisplayRev(TheString As String, Data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim C As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte

    n = 2 * Len(TheString)
    data2 = TheString

    ReDim Data(n \ 4 - 1)

    d = 0
    i = 0
    j = 0
    Do While j < n
        C = data2(j)
        Select Case C
        Case 48 To 57    '"0" ... "9"
            If d = 0 Then   'high
                d = C
            Else            'low
                Data(i) = (C - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = C - 7
            Else            'low
                Data(i) = (C - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = C - 39
            Else            'low
                Data(i) = (C - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase Data
    Else
        ReDim Preserve Data(n - 1)
    End If
    HexDisplayRev = n

End Function

Public Function DecryptString(ByVal ToDecryptString As String, ByVal KeyString As String) As String
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(ToDecryptString) = 0 Then
        DecryptString = ""
    Else
        If Len(KeyString) = 0 Then
            DecryptString = ""
        Else
            KeyBits = 192        '128 192 256
            BlockBits = 128
            pass = GetPassword(KeyString)

            If HexDisplayRev(ToDecryptString, ciphertext) = 0 Then
                DecryptString = ""
                Exit Function
            End If

            SetCipherKey pass, KeyBits
            If ArrayDecrypt(plaintext, ciphertext, 0) <> 0 Then
                DecryptString = ""
                Exit Function
            End If
            DecryptString = StrConv(plaintext, vbUnicode)
        End If
    End If

End Function



