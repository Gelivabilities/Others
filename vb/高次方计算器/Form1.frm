VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����ߴη���ȷ������"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   12015
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   3375
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   10575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaxLength       =   2
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "�η�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   27
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   27
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function CalcDec(ByVal NumA As String, ByVal NumB As String) As String

Dim Temp As String

Dim TempA As String

Dim TempB As String

Dim XS As Long

Dim AFS As Boolean

Dim BFS As Boolean

Dim FS As Boolean

Dim TW As Boolean

'�����ı����

For i = 1 To Len(NumA)

    If Mid(NumA, i, 1) = "-" Then AFS = Not AFS Else: NumA = Mid(NumA, i): Exit For

Next

For i = 1 To Len(NumB)

    If Mid(NumB, i, 1) = "-" Then BFS = Not BFS Else: NumB = Mid(NumB, i): Exit For

Next

If AFS And Not BFS Then

    '����������,���ɼӷ�

    CalcDec = "-" & CalcAdd(NumA, NumB)

    Exit Function

ElseIf Not AFS And BFS Then

    '����������,���ɼӷ�

    CalcDec = CalcAdd(NumA, NumB)

    Exit Function

ElseIf AFS And BFS Then

    '��������������������

    Temp = NumA: NumA = NumB: NumB = Temp

End If

'���룬���С��

XS = XSD(NumA, NumB)

'������ǰ

For i = 1 To Len(NumA)

    If Val(Mid(NumA, i, 1)) < Val(Mid(NumB, i, 1)) Then Temp = NumA: NumA = NumB: NumB = Temp: FS = True: Exit For

    If Val(Mid(NumA, i, 1)) > Val(Mid(NumB, i, 1)) Then: Exit For

Next

'��λ

Temp = ""

For i = Len(NumB) To 1 Step -1

    If IIf(TW, Val(Mid(NumA, i, 1)) - 1, Val(Mid(NumA, i, 1))) < Val(Mid(NumB, i, 1)) Then

        If TW = True Then

            Temp = Chr(65 + Val(Mid(NumA, i, 1)) - 1) & Temp: TW = True

        Else

            Temp = Chr(65 + Val(Mid(NumA, i, 1))) & Temp: TW = True

        End If

    Else

        If TW = True Then

            If Val(Mid(NumA, i, 1)) = 0 Then

                Temp = 9 & Temp

            Else

                Temp = Val(Mid(NumA, i, 1)) - 1 & Temp

                TW = False

            End If

        Else

                Temp = Val(Mid(NumA, i, 1)) & Temp

        End If

    End If

Next

NumA = Temp

'������ʽ

Temp = ""

For i = Len(NumA) To 1 Step -1

TempA = IIf(Mid(NumA, i, 1) > "9", Asc(Mid(NumA, i, 1)) - 55, Mid(NumA, i, 1))

Temp = TempA - Val(Mid(NumB, i, 1)) & Temp

Next

'����С����

If XS > 0 And Temp <> "0" Then

    Temp = Left(Temp, Len(Temp) - XS) & "." & Right(Temp, XS)

End If

'��������

CalcDec = ValNum(IIf(FS, "-" & Temp, Temp))

End Function

'�������ӷ�����

'���躯��������[������]

Public Function CalcAdd(ByVal NumA As String, ByVal NumB As String) As String

Dim Temp As String

Dim V As String

Dim Sw As String

Dim Gw As String

Dim TempS As String

Dim TempD As String

Dim XS As Long

Dim AFS As Boolean

Dim BFS As Boolean

Dim FS As Boolean

'�����ı����

For i = 1 To Len(NumA)

    If Mid(NumA, i, 1) = "-" Then AFS = Not AFS Else: NumA = Mid(NumA, i): Exit For

Next

For i = 1 To Len(NumB)

    If Mid(NumB, i, 1) = "-" Then BFS = Not BFS Else: NumB = Mid(NumB, i): Exit For

Next

If AFS Xor BFS Then

    '�����ӷ����������ɼ���

    If AFS Then

        CalcAdd = CalcDec(NumB, NumA)

    Else

        CalcAdd = CalcDec(NumA, NumB)

    End If

    Exit Function

ElseIf AFS And BFS Then

    FS = True

End If

'���룬���С��

XS = XSD(NumA, NumB)

'�ӷ���ʽ

For i = Len(NumA) To 1 Step -1

    Temp = Format(Val(Mid(NumA, i, 1)) + Val(Mid(NumB, i, 1)), "00")

    Temp = Format(Val(Temp) + Val(Sw), "00")

    Gw = Right(Temp, 1) & Gw

    Sw = Left(Temp, 1)

Next

Temp = Sw & Gw

'����С����

If XS > 0 Then

    Temp = Left(Temp, Len(Temp) - XS) & "." & Right(Temp, XS)

End If

'��������

CalcAdd = ValNum(IIf(FS, "-" & Temp, Temp))

End Function

'�������˷�����

'���躯�����ӷ�[������]

Public Function CalcMul(ByVal NumA As String, ByVal NumB As String) As String

Dim Str() As String

Dim s As String, XS As Long

Dim Temp As String

Dim FS As Boolean

'�������

For i = 1 To Len(NumA)

    If Mid(NumA, i, 1) = "-" Then FS = Not FS Else: NumA = Mid(NumA, i): Exit For

Next

For i = 1 To Len(NumB)

    If Mid(NumB, i, 1) = "-" Then FS = Not FS Else: NumB = Mid(NumB, i): Exit For

Next

'С�������

XS = Len(NumA) - IIf(InStr(NumA, ".") > 0, InStr(NumA, "."), Len(NumA))

XS = XS + Len(NumB) - IIf(InStr(NumB, ".") > 0, InStr(NumB, "."), Len(NumB))

NumA = Replace(NumA, ".", ""): NumB = Replace(NumB, ".", "")

ReDim Str(Len(NumB)) As String

'�˷���ʽ

For y = Len(NumB) To 1 Step -1

s = Mid(NumB, y, 1)

Gw = ""

Sw = ""

     

    For i = Len(NumA) To 1 Step -1

        Temp = Format(Val(s) * Val(Mid(NumA, i, 1)), "00")

        Temp = Format(Val(Temp) + Val(Sw), "00")

     

        Gw = Right(Temp, 1) & Gw

        Sw = Left(Temp, 1)

    Next

Temp = Sw + Gw + String(Len(NumB) - y, "0")

Str(y) = Temp

Next

'���ɼӷ�

For i = 1 To Len(NumB) - 1

Str(i + 1) = CalcAdd(Str(i), Str(i + 1))

Next

Temp = Str(Len(NumB))

'����С����

If XS > 0 Then

    Temp = Left(Temp, Len(Temp) - XS) & "." & Right(Temp, XS)

End If

'��������

CalcMul = ValNum(IIf(FS, "-" & Temp, Temp))

End Function

'----------------------------------------------------------'��������������

'���躯�����˷�[������]������[������]

'��ע���˷������漰���ӷ�[������],����NumLenΪ��ЧС��λ��,Ĭ��Ϊ20λ,���Ϊ255λ,�����޸�ByteΪLong�������С����������������������

'�����������ٶȡ�

Public Function CalcDiv(ByVal NumA As String, ByVal NumB As String, Optional ByVal NumLen As Byte = 20) As String

Dim Temp As String

Dim TempA As String

Dim TempB As String

Dim TempC As Long

Dim DocS As Boolean, DocN As Long

'���С��

If InStr(NumA, ".") = 0 Then NumA = NumA & "."

'����ȥ��

If InStr(NumB, ".") > 0 Then

TempC = Len(Mid(NumB, InStr(NumB, ".") + 1))

NumB = Replace(NumB, ".", "")

Temp = Mid(NumA, InStr(NumA, ".") + 1)

If Len(Temp) < TempC Then

    NumA = Replace(NumA, ".", "") & String(TempC - Len(Temp), "0") & "."

Else

    NumA = Left(NumA, InStr(NumA, ".") - 1) & Left(Temp, TempC) & "." & Mid(Temp, TempC + 1)

End If

End If

Temp = ""

TempC = 0

'������ʽ

CalcStatic:

TempA = TempB & Mid(NumA, TempC + 1, 1)

If (TempB = "0" And Mid(NumA, TempC + 1, 1) = "") Or DocN >= NumLen Then CalcDiv = ValNum(Temp): Exit Function

If TempA = TempB Then TempA = TempA & "0"

If Right(TempA, 1) = "." Then TEMB = TempA: TempC = TempC + 1: DocS = True: Temp = Temp & ".": GoTo CalcStatic:

TempA = Replace(TempA, ".", "0")

'�˷�����

For i = 1 To 10

    If CalcMin(TempA, CalcMul(NumB, i)) Then Temp = Temp & i - 1: Exit For

Next

'�������

TempB = CalcDec(TempA, CalcMul(NumB, i - 1))

TempC = TempC + 1

'С��λ��

If DocS Then DocN = DocN + 1

GoTo CalcStatic:

End Function


'�������Ӽ������ֶ��벢���С����λ��

Public Function XSD(ByRef NumA As String, ByRef NumB As String) As Long

Dim TempA As String

Dim TempB As String

'û��С���������ĩβ

If InStr(NumA, ".") = 0 Then NumA = NumA & "."

If InStr(NumB, ".") = 0 Then NumB = NumB & "."

'��������λ

TempA = Left(NumA, InStr(NumA, "."))

TempB = Left(NumB, InStr(NumB, "."))

If Len(TempA) < Len(TempB) Then

    NumA = String(Len(TempB) - Len(TempA), "0") & NumA

ElseIf Len(TempA) > Len(TempB) Then

    NumB = String(Len(TempA) - Len(TempB), "0") & NumB

End If

'����С��λ

TempA = Mid(NumA, InStr(NumA, ".") + 1)

TempB = Mid(NumB, InStr(NumB, ".") + 1)

If Len(TempA) < Len(TempB) Then

    NumA = NumA & String(Len(TempB) - Len(TempA), "0")

ElseIf Len(TempA) > Len(TempB) Then

    NumB = NumB & String(Len(TempA) - Len(TempB), "0")

End If

'��¼С��λ��

XSD = Len(NumA) - IIf(InStr(NumA, ".") > 0, InStr(NumA, "."), Len(NumA))

'ȥ��С����

NumA = Replace(NumA, ".", ""): NumB = Replace(NumB, ".", "")

End Function

'������ȥ��������

Public Function ValNum(ByVal Num As String) As String

Dim Temp As String

Dim TempA As String

Dim TempB As String

Temp = Len(Num) - IIf(InStr(Num, ".") > 0, InStr(Num, "."), Len(Num))

TempA = Left(Num, Len(Num) - Temp)

TempB = Right(Num, Temp)

Do Until Len(TempA) = 1

    If Left(TempA, 1) <> "0" Then Exit Do

    TempA = Mid(TempA, 2)

Loop

If Left(TempA, 1) = "." Then TempA = "0" & TempA

Do Until Len(TempB) = 0

    If Val(Right(TempB, 1)) <> "0" Then Exit Do

    TempB = Left(TempB, Len(TempB) - 1)

Loop

If TempB = "" And Right(TempA, 1) = "." Then TempA = Left(TempA, Len(TempA) - 1)

ValNum = TempA & TempB

End Function

'�������Ƚϴ�С

Public Function CalcMin(ByVal NumA As String, ByVal NumB As String) As Boolean

XSD NumA, NumB

For i = 1 To Len(NumA)

    If Val(Mid(NumA, i, 1)) < Val(Mid(NumB, i, 1)) Then: CalcMin = True: Exit For

    If Val(Mid(NumA, i, 1)) > Val(Mid(NumB, i, 1)) Then: CalcMin = False: Exit For

Next

End Function
Public Function power(ByVal x As String, ByVal y As String) As String
Dim s As String
s = 1
For i = 1 To y
    s = CalcMul(s, x)
Next
power = s
End Function
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
Text3.Text = "��������������"
Else
If Text1.Text = 0 And Text2.Text = 0 Then
Text3.Text = "������"
Else
Text3.Text = power(Text1.Text, Text2.Text)
End If
End If
End Sub

Private Sub Text1_keypress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If keyascii < 48 Or keyascii > 57 Then keyascii = 0
End Sub
Private Sub Text2_keypress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If keyascii < 48 Or keyascii > 57 Then keyascii = 0
End Sub
