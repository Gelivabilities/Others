VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ȴ���..."
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   4200
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton Option1 
      Caption         =   "�������"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "����"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   4920
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1575
      Left            =   4200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   -1.00000e5
      Width           =   2535
      ExtentX         =   4471
      ExtentY         =   2778
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����ѡ��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "�ؼ���"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim songs(0 To 2000) As String
Dim songPacks(0 To 200) As String
Dim xiangguanxiazai(0 To 100) As String

Public downloadAddress As String
Public newSourceCode As String

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
   Dim lngRetVal As Long
   lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
   If lngRetVal = 0 Then DownloadFile = True
End Function

Private Sub Command1_Click()
Shell "explorer http://pan.baidu.com/s/1gdzRiFL"
End Sub

Private Sub Command2_Click()
a = GetBody("http://jq9837.com/taikojiro/data.htm")
If a = "" Then
MsgBox "û��( �� �� ) ~��"
End
End If

WebBrowser1.Navigate downloadAddress
List1.SetFocus
End Sub

Private Sub Command3_Click()
Shell "explorer http://taikomg.com/"
End Sub

Private Sub Command4_Click()
Shell "explorer http://pan.baidu.com/s/1eQ6EOl8"
End Sub

Private Sub Form_Load()
Form1.Show

newSourceCode = sourceCodeToListCode("http://jq9837.com/taikojiro/data.htm")
buildList (newSourceCode)

Form1.Caption = "̫�Ĵ����ļ����������� - V1.0 Beta"
Text1.ToolTipText = "ֱ�����뽫�Զ�����"
End Sub
'��ȡԴ���룬Դ�������׶�������
Public Function sourceCodeToListCode(a As String)
Dim sourceCode As String
sourceCode = GetBody(a)
If sourceCode = "" Then
MsgBox "û��( �� �� ) ~��"
End
End If
sourceCodeToListCode = sourceCodeProcess(sourceCode)
End Function

'����
Public Function filter(keyWord As String)
n = 0
For i = 0 To 2000
If (InStr(LCase(songs(i)), LCase(keyWord)) <> 0 Or InStr(LCase(getKeyWords(songs(i))), LCase(keyWord)) <> 0) And (songs(i) = "" = False) Then
List1.AddItem songs(i), n
n = n + 1
End If
Next
End Function

'�ַ���C��ȡ��λ���ַ���A��B֮���һ����
Public Function getStr(str, head, ending As String)
headLen = Len(head)
endingLen = Len(ending)
strLen = Len(str)
getStr = Mid(str, headLen + 1, strLen - headLen - endingLen)
End Function

'ȥ������ַ���
Public Function delStrLeft(str, keyWord As String, move As Integer)
keywordplace = InStr(str, keyWord)
lengthStrToDel = Len(keyWord)
delStrLeft = Mid(str, keywordplace + lengthStrToDel + move)
End Function

'ȥ���ұ��ַ���
Public Function delStrRight(str, keyWord As String, move As Integer)
keywordplace = InStr(str, keyWord)
lengthStrToDel = Len(keyWord)
delStrRight = Replace(str, Mid(str, keywordplace + lengthStrToDel + move), "")
End Function
'��ȡ��ҳԴ����
Public Function GetBody(ByVal URL$, Optional ByVal Coding$ = "UTF-8")
    Dim ObjXML
    On Error Resume Next
    Set ObjXML = CreateObject("Microsoft.XMLHTTP")
    With ObjXML
        .Open "Get", URL, False, "", ""
        .setRequestHeader "If-Modified-Since", "0"
        .Send
        GetBody = .responseBody
    End With
    GetBody = BytesToBstr(GetBody, Coding)
    Set ObjXML = Nothing
End Function
'��ȡ��ҳԴ����
Public Function BytesToBstr(strBody, CodeBase)
    Dim ObjStream
    Set ObjStream = CreateObject("Adodb.Stream")
    With ObjStream
        .Type = 1
        .Mode = 3
        .Open
        .Write strBody
        .Position = 0
        .Type = 2
        .Charset = CodeBase
        BytesToBstr = .ReadText
        .Close
    End With
    Set ObjStream = Nothing
End Function

'Դ�������׶������֣�[����]""[�ؼ���]""[���ص�ַ]""��
Public Function sourceCodeProcess(sourceCode As String)
'����ͷ
a = delStrLeft(sourceCode, "<span lang=", 5)
a = "����[" & a

'�����β
a = delStrRight(a, "</body>", -25) & "]"

'����ʣ�²���
a = Replace(a, vbCrLf, "")
a = Replace(a, "</span>&lt;/name&gt;&lt;url&gt;", "]���ص�ַ[")
a = Replace(a, "&lt;/url&gt;", "]" & vbCrLf & "����[")
a = Replace(a, "</p>", "")
a = Replace(a, "<p>&lt;name&gt;", "")
a = Replace(a, "&lt;/name&gt;&lt;url&gt;", "]���ص�ַ[")
a = Replace(a, "&lt;/name&gt;&lt;keywords&gt;", "]�ؼ���[")
a = Replace(a, "&lt;/keywords&gt;&lt;url&gt;", "]���ص�ַ[")
a = "����[" & delStrLeft(a, "��", 0)
a = Replace(a, "&lt;/u", "")
sourceCodeProcess = a
End Function
'�����б�
Public Function buildList(newSourceCode As String)
a = newSourceCode
i = 0

Do While InStr(a, "����") <> 0
    songs(i) = delStrRight(delStrLeft(a, "����[", 0), "]", -1)
    a = delStrLeft(a, "����[", 0)
    i = i + 1
Loop


End Function
'ѡ���б��е�һ��ʱ������ѡ�б��������õ����ص�ַ���Ա���Ӧ����ѡ�а�ť
Private Sub list1_Click()

For i = 0 To List1.ListIndex
    If List1.Selected(i) = True Then
        downloadAddress = delStrLeft(newSourceCode, List1.Text & "]", -1)
        downloadAddress = delStrLeft(downloadAddress, "]���ص�ַ[", 0)
        downloadAddress = delStrRight(downloadAddress, "]", -1)
        Command2.Enabled = True
        i = i + 1
    End If
Next

End Sub
'���س���������ѡ��
Private Sub list1_keypress(keyascii As Integer)
If keyascii = 13 Then Call Command2_Click
End Sub

'��ȡ�ؼ���
Public Function getKeyWords(listText As String)

        words = delStrLeft(newSourceCode, listText & "]", -1)
        words = "]" & delStrLeft(words, "]�ؼ���[", 0)
        words = delStrLeft(words, "]", 0)
        words = delStrRight(words, "]", -1)
getKeyWords = words
End Function

'�ж�ѡ�е��������ϣ������������
Private Sub Option1_Click(index As Integer)
    Command2.Enabled = False
    If Option1(0).Value = True Then
        List1.Clear
        filter Text1.Text
    End If
End Sub
'���ָı�ʱˢ���б�
Private Sub Text1_Change()

    Command2.Enabled = False '����ѡ�в��ɵ�
    
    If Option1(0).Value = True Then '������ѡ��
            If Len(Text1.Text) <> 0 Then
                List1.Clear
                filter (Text1.Text)
            Else
                List1.Clear
                filter ("")
            End If
    End If
End Sub
