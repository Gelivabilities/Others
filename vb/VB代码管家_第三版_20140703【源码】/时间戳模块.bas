Attribute VB_Name = "Module3"
Public Function ʱ���A() As String  '�������Ϊ13λ��ʱ���
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
    ʱ���A = obj.Eval(ShiJianChuocode)
End Function

Public Function ʱ���B() As String  '�������Ϊ10λ��ʱ���
    ʱ���B = DateDiff("s", "01/01/1970 00:00:00", Now())
End Function
