VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "等待中..."
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.OptionButton Option1 
      Caption         =   "相关下载"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整合"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "单曲"
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
      Caption         =   "下载选中"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "关键词"
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
MsgBox "没网( ﹁ ﹁ ) ~→"
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

Form1.Caption = "太鼓次郎文件搜索下载器 - V1.0 Beta"
Text1.ToolTipText = "直接输入将自动搜索"
End Sub
'获取源代码，源代码变成易读的文字
Public Function sourceCodeToListCode(a As String)
Dim sourceCode As String
sourceCode = GetBody(a)
If sourceCode = "" Then
MsgBox "没网( ﹁ ﹁ ) ~→"
End
End If
sourceCodeToListCode = sourceCodeProcess(sourceCode)
End Function

'搜索
Public Function filter(keyWord As String)
n = 0
For i = 0 To 2000
If (InStr(LCase(songs(i)), LCase(keyWord)) <> 0 Or InStr(LCase(getKeyWords(songs(i))), LCase(keyWord)) <> 0) And (songs(i) = "" = False) Then
List1.AddItem songs(i), n
n = n + 1
End If
Next
End Function

'字符串C截取出位于字符串A和B之间的一部分
Public Function getStr(str, head, ending As String)
headLen = Len(head)
endingLen = Len(ending)
strLen = Len(str)
getStr = Mid(str, headLen + 1, strLen - headLen - endingLen)
End Function

'去掉左边字符串
Public Function delStrLeft(str, keyWord As String, move As Integer)
keywordplace = InStr(str, keyWord)
lengthStrToDel = Len(keyWord)
delStrLeft = Mid(str, keywordplace + lengthStrToDel + move)
End Function

'去掉右边字符串
Public Function delStrRight(str, keyWord As String, move As Integer)
keywordplace = InStr(str, keyWord)
lengthStrToDel = Len(keyWord)
delStrRight = Replace(str, Mid(str, keywordplace + lengthStrToDel + move), "")
End Function
'获取网页源代码
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
'获取网页源代码
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

'源代码变成易读的文字（[名称]""[关键词]""[下载地址]""）
Public Function sourceCodeProcess(sourceCode As String)
'处理开头
a = delStrLeft(sourceCode, "<span lang=", 5)
a = "名称[" & a

'处理结尾
a = delStrRight(a, "</body>", -25) & "]"

'处理剩下部分
a = Replace(a, vbCrLf, "")
a = Replace(a, "</span>&lt;/name&gt;&lt;url&gt;", "]下载地址[")
a = Replace(a, "&lt;/url&gt;", "]" & vbCrLf & "名称[")
a = Replace(a, "</p>", "")
a = Replace(a, "<p>&lt;name&gt;", "")
a = Replace(a, "&lt;/name&gt;&lt;url&gt;", "]下载地址[")
a = Replace(a, "&lt;/name&gt;&lt;keywords&gt;", "]关键词[")
a = Replace(a, "&lt;/keywords&gt;&lt;url&gt;", "]下载地址[")
a = "名称[" & delStrLeft(a, "　", 0)
a = Replace(a, "&lt;/u", "")
sourceCodeProcess = a
End Function
'生成列表
Public Function buildList(newSourceCode As String)
a = newSourceCode
i = 0

Do While InStr(a, "名称") <> 0
    songs(i) = delStrRight(delStrLeft(a, "名称[", 0), "]", -1)
    a = delStrLeft(a, "名称[", 0)
    i = i + 1
Loop


End Function
'选中列表中的一个时，下载选中变亮，并得到下载地址，以便响应下载选中按钮
Private Sub list1_Click()

For i = 0 To List1.ListIndex
    If List1.Selected(i) = True Then
        downloadAddress = delStrLeft(newSourceCode, List1.Text & "]", -1)
        downloadAddress = delStrLeft(downloadAddress, "]下载地址[", 0)
        downloadAddress = delStrRight(downloadAddress, "]", -1)
        Command2.Enabled = True
        i = i + 1
    End If
Next

End Sub
'按回车就能下载选中
Private Sub list1_keypress(keyascii As Integer)
If keyascii = 13 Then Call Command2_Click
End Sub

'截取关键词
Public Function getKeyWords(listText As String)

        words = delStrLeft(newSourceCode, listText & "]", -1)
        words = "]" & delStrLeft(words, "]关键词[", 0)
        words = delStrLeft(words, "]", 0)
        words = delStrRight(words, "]", -1)
getKeyWords = words
End Function

'判断选中单曲，整合，还是相关下载
Private Sub Option1_Click(index As Integer)
    Command2.Enabled = False
    If Option1(0).Value = True Then
        List1.Clear
        filter Text1.Text
    End If
End Sub
'文字改变时刷新列表
Private Sub Text1_Change()

    Command2.Enabled = False '下载选中不可点
    
    If Option1(0).Value = True Then '单曲被选中
            If Len(Text1.Text) <> 0 Then
                List1.Clear
                filter (Text1.Text)
            Else
                List1.Clear
                filter ("")
            End If
    End If
End Sub
