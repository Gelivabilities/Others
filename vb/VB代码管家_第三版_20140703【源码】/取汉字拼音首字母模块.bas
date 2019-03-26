Attribute VB_Name = "Module1"
Public Function py(mystr As String) As String
    If Asc(mystr) < 0 Then
        If Asc(Left$(mystr, 1)) < Asc("°¡") Then
            py = "0"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("°¡") And Asc(Left$(mystr, 1)) < Asc("°Å") Then
            py = "A"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("°Å") And Asc(Left$(mystr, 1)) < Asc("²Á") Then
            py = "B"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("²Á") And Asc(Left$(mystr, 1)) < Asc("´î") Then
            py = "C"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("´î") And Asc(Left$(mystr, 1)) < Asc("¶ê") Then
            py = "D"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("¶ê") And Asc(Left$(mystr, 1)) < Asc("·¢") Then
            py = "E"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("·¢") And Asc(Left$(mystr, 1)) < Asc("¸Á") Then
            py = "F"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("¸Á") And Asc(Left$(mystr, 1)) < Asc("¹þ") Then
            py = "G"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("¹þ") And Asc(Left$(mystr, 1)) < Asc("»÷") Then
            py = "H"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("»÷") And Asc(Left$(mystr, 1)) < Asc("¿¦") Then
            py = "J"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("¿¦") And Asc(Left$(mystr, 1)) < Asc("À¬") Then
            py = "K"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("À¬") And Asc(Left$(mystr, 1)) < Asc("Âè") Then
            py = "L"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Âè") And Asc(Left$(mystr, 1)) < Asc("ÄÃ") Then
            py = "M"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("ÄÃ") And Asc(Left$(mystr, 1)) < Asc("Å¶") Then
            py = "N"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Å¶") And Asc(Left$(mystr, 1)) < Asc("Å¾") Then
            py = "O"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Å¾") And Asc(Left$(mystr, 1)) < Asc("ÆÚ") Then
            py = "P"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("ÆÚ") And Asc(Left$(mystr, 1)) < Asc("È»") Then
            py = "Q"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("È»") And Asc(Left$(mystr, 1)) < Asc("Èö") Then
            py = "R"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Èö") And Asc(Left$(mystr, 1)) < Asc("Ëú") Then
            py = "S"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Ëú") And Asc(Left$(mystr, 1)) < Asc("ÍÚ") Then
            py = "T"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("ÍÚ") And Asc(Left$(mystr, 1)) < Asc("Îô") Then
            py = "W"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Îô") And Asc(Left$(mystr, 1)) < Asc("Ñ¹") Then
            py = "X"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("Ñ¹") And Asc(Left$(mystr, 1)) < Asc("ÔÑ") Then
            py = "Y"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("ÔÑ") Then
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
'-----------------------ÒÔÉÏÒ²¿ÉÖ±½Ó·Åµ½Ä£¿éÀï-----------------------

