Attribute VB_Name = "Module4"
Public Function GetData(ByVal Url As String, ByVal CodeBase As String) As Variant
On Error GoTo CHUCUO:
    Dim XMLHTTP As Object
    Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
    XMLHTTP.Open "Get", Url, True
    XMLHTTP.send
    '--------------------------------------发送数据
    While XMLHTTP.ReadyState <> 4
        DoEvents
    Wend
    '--------------------------------------函数返回
    GetData = XMLHTTP.ResponseBody
    Form4.协议头.Text = XMLHTTP.GetAllResponseHeaders
    If CStr(GetData) <> "" Then GetData = BytesToBstr(GetData, CodeBase)
    Set XMLHTTP = Nothing
    Exit Function
CHUCUO:
    Set XMLHTTP = Nothing
    GetData = ""
End Function

Public Function PostData(ByVal StrUrl As String, ByVal StrData As String, ByVal CodeBase As String) As Variant
On Error GoTo CHUCUO:
    Dim XMLHTTP As Object
    Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
    XMLHTTP.Open "POST", StrUrl, True
    XMLHTTP.setRequestHeader "Content-Length", Len(PostData)
    XMLHTTP.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"
    XMLHTTP.send (StrData)
    '--------------------------------------发送数据
    Do Until XMLHTTP.ReadyState = 4
        DoEvents
    Loop
    '--------------------------------------函数返回
    PostData = XMLHTTP.ResponseBody
    Form4.协议头.Text = XMLHTTP.GetAllResponseHeaders
    If CStr(PostData) <> "" Then PostData = BytesToBstr(PostData, CodeBase)
    Set XMLHTTP = Nothing
    Exit Function
CHUCUO:
    Set XMLHTTP = Nothing
    PostData = ""
End Function

Public Function BytesToBstr(strBody, CodeBase) '判断编码
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
