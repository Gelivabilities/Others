Attribute VB_Name = "Module1"
Public Function py(mystr As String) As String
    If Asc(mystr) < 0 Then
        If Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "0"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "A"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "B"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "C"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "D"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "E"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "F"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "G"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "H"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "J"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "K"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "L"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "M"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("Ŷ") Then
            py = "N"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Ŷ") And Asc(Left$(mystr, 1)) < Asc("ž") Then
            py = "O"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("ž") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "P"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("Ȼ") Then
            py = "Q"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Ȼ") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "R"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "S"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "T"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "W"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") And Asc(Left$(mystr, 1)) < Asc("ѹ") Then
            py = "X"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("ѹ") And Asc(Left$(mystr, 1)) < Asc("��") Then
            py = "Y"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("��") Then
            py = "Z"
            Exit Function
        End If
    Else
        If UCase$(mystr) <= "Z" And UCase$(mystr) >= "A" Then
            py = UCase$(Left$(mystr, 1))
        Else
            py = mystr
        End If
    End If
End Function

Public Function test(str As String) As String
    Dim tmp As String
    For i = 1 To Len(str)
        tmp = tmp & py(Mid$(str, i, 1))
    Next i
    test = tmp
End Function
'-----------------------����Ҳ��ֱ�ӷŵ�ģ����-----------------------

