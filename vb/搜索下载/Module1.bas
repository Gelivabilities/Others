Attribute VB_Name = "Module1"
Option Explicit
Private Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private arrBase64() As String

Public Function Base64Encode(strSource As String) As String '±àÂë
    On Error Resume Next
    If UBound(arrBase64) = -1 Then
        arrBase64 = Split(StrConv(cstBase64, vbUnicode), vbNullChar)
    End If
    Dim arrB() As Byte, bTmp(2) As Byte, bT As Byte
    Dim I As Long, J As Long
    arrB = StrConv(strSource, vbFromUnicode)

    J = UBound(arrB)
    For I = 0 To J Step 3
        Erase bTmp
        bTmp(0) = arrB(I + 0)
        bTmp(1) = arrB(I + 1)
        bTmp(2) = arrB(I + 2)

        bT = (bTmp(0) And 252) / 4
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(0) And 3) * 16
        bT = bT + bTmp(1) \ 16
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(1) And 15) * 4
        bT = bT + bTmp(2) \ 64
        If I + 1 <= J Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If

        bT = bTmp(2) And 63
        If I + 2 <= J Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If
    Next
End Function

Public Function Base64Decode(strEncoded As String) As String '½âÂë
    On Error Resume Next
    Dim arrB() As Byte, bTmp(3) As Byte, bT As Long, bRet() As Byte
    Dim I As Long, J As Long
    arrB = StrConv(strEncoded, vbFromUnicode)
    J = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(J - J \ 4 - 1)
    For I = 0 To J Step 4
        Erase bTmp
        bTmp(0) = (InStr(cstBase64, Chr(arrB(I))) - 1) And 63
        bTmp(1) = (InStr(cstBase64, Chr(arrB(I + 1))) - 1) And 63
        bTmp(2) = (InStr(cstBase64, Chr(arrB(I + 2))) - 1) And 63
        bTmp(3) = (InStr(cstBase64, Chr(arrB(I + 3))) - 1) And 63

        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)

        bRet((I \ 4) * 3) = bT \ 65536
        bRet((I \ 4) * 3 + 1) = (bT And 65280) \ 256
        bRet((I \ 4) * 3 + 2) = bT And 255
    Next
    Base64Decode = StrConv(bRet, vbUnicode)
End Function

