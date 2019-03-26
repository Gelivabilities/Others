Attribute VB_Name = "Module2"
Public Function ToUnicode(ByRef str As String) As String 'Unicode编码
    Dim code As String
    Dim obj As Object
    Set obj = CreateObject("MSScriptControl.ScriptControl")
    obj.AllowUI = True
    obj.Language = "JavaScript"
    
    code = code & "function ToUnicode(str)"
    code = code & "{"
    code = code & "return escape(str).replace(/%/g," & Chr(34) & "\\" & Chr(34) & ").toLowerCase();"
    code = code & "}"
    code = code & "ToUnicode (" & Chr(34) & str & Chr(34) & ")"
    
    ToUnicode = obj.Eval(code) '输出结果
End Function

Public Function UnUnicode(ByRef str As String) As String 'Unicode解码
    Dim code As String
    Dim obj As Object
    Set obj = CreateObject("MSScriptControl.ScriptControl")
    obj.AllowUI = True
    obj.Language = "JavaScript"
    
    code = code & "function UnUnicode(str)"
    code = code & "{"
    code = code & "return unescape(str.replace(/\\/g, " & Chr(34) & "%" & Chr(34) & "));"
    code = code & "}"
    code = code & "UnUnicode (" & Chr(34) & str & Chr(34) & ")"
    
    UnUnicode = obj.Eval(code) '输出结果
End Function

Public Function UTF8_URLEncoding(szInput) 'UTF-8 URL编码
    Dim wch, uch, szRet
    Dim x
    Dim nAsc, nAsc2, nAsc3
    If szInput = "" Then
        UTF8_URLEncoding = szInput
        Exit Function
    End If
    For x = 1 To Len(szInput)
        wch = Mid(szInput, x, 1)
        nAsc = AscW(wch)
        
        If nAsc < 0 Then nAsc = nAsc + 65536
        
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    UTF8_URLEncoding = szRet
End Function

Public Function UTF8_UrlDecode(ByVal Url As String) 'UTF-8 URL解码
    Dim B, ub   ''中文字的Unicode码(2字节)
    Dim UtfB    ''Utf-8单个字节
    Dim UtfB1, UtfB2, UtfB3 ''Utf-8码的三个字节
    Dim i, n, s
    n = 0
    ub = 0
    For i = 1 To Len(Url)
        B = Mid(Url, i, 1)
        Select Case B
        Case "+"
            s = s & " "
        Case "%"
            ub = Mid(Url, i + 1, 2)
            UtfB = CInt("&H" & ub)
            If UtfB < 128 Then
                i = i + 2
                s = s & ChrW(UtfB)
            Else
                UtfB1 = (UtfB And &HF) * &H1000   ''取第1个Utf-8字节的二进制后4位
                UtfB2 = (CInt("&H" & Mid(Url, i + 4, 2)) And &H3F) * &H40      ''取第2个Utf-8字节的二进制后6位
                UtfB3 = CInt("&H" & Mid(Url, i + 7, 2)) And &H3F      ''取第3个Utf-8字节的二进制后6位
                s = s & ChrW(UtfB1 Or UtfB2 Or UtfB3)
                i = i + 8
            End If
        Case Else    ''Ascii码
            s = s & B
        End Select
    Next
    UTF8_UrlDecode = s
End Function

Public Function URLEncode(ByRef StrUrl As String) As String 'GBK URL编码
    Dim i As Long
    Dim tempStr As String
    For i = 1 To Len(StrUrl)
        If Asc(Mid(StrUrl, i, 1)) < 0 Then
            tempStr = "%" & Right(CStr(Hex(Asc(Mid(StrUrl, i, 1)))), 2)
            tempStr = "%" & Left(CStr(Hex(Asc(Mid(StrUrl, i, 1)))), Len(CStr(Hex(Asc(Mid(StrUrl, i, 1))))) - 2) & tempStr
            URLEncode = URLEncode & tempStr
        ElseIf (Asc(Mid(StrUrl, i, 1)) >= 65 And Asc(Mid(StrUrl, i, 1)) <= 90) Or (Asc(Mid(StrUrl, i, 1)) >= 97 And Asc(Mid(StrUrl, i, 1)) <= 122) Then
            URLEncode = URLEncode & Mid(StrUrl, i, 1)
        Else
            URLEncode = URLEncode & "%" & Hex(Asc(Mid(StrUrl, i, 1)))
        End If
    Next
End Function

Public Function URLDecode(ByRef StrUrl As String) As String 'GBK URL解码
    Dim i As Long
    If InStr(StrUrl, "%") = 0 Then URLDecode = StrUrl: Exit Function
    For i = 1 To Len(StrUrl)
        If Mid(StrUrl, i, 1) = "%" Then
            If Val("&H" & Mid(StrUrl, i + 1, 2)) > 127 Then
                URLDecode = URLDecode & Chr(Val("&H" & Mid(StrUrl, i + 1, 2) & Mid(StrUrl, i + 4, 2)))
                i = i + 5
            Else
                URLDecode = URLDecode & Chr(Val("&H" & Mid(StrUrl, i + 1, 2)))
                i = i + 2
            End If
        Else
            URLDecode = URLDecode & Mid(StrUrl, i, 1)
        End If
    Next
End Function


