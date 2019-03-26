Attribute VB_Name = "Module3"
Public Function 时间戳A() As String  '输出长度为13位的时间戳
    Dim ShiJianChuocode As String
    Dim obj As Object
    Set obj = CreateObject("MSScriptControl.ScriptControl")
    obj.AllowUI = True
    obj.Language = "JavaScript"
    ShiJianChuocode = ShiJianChuocode & "function abc()" & vbCrLf
    ShiJianChuocode = ShiJianChuocode & "{" & vbCrLf
    ShiJianChuocode = ShiJianChuocode & "var timestamp = new Date().getTime();" & vbCrLf
    ShiJianChuocode = ShiJianChuocode & "return timestamp;" & vbCrLf
    ShiJianChuocode = ShiJianChuocode & "}" & vbCrLf
    ShiJianChuocode = ShiJianChuocode & "abc()" & vbCrLf
    时间戳A = obj.Eval(ShiJianChuocode)
End Function

Public Function 时间戳B() As String  '输出长度为10位的时间戳
    时间戳B = DateDiff("s", "01/01/1970 00:00:00", Now())
End Function
