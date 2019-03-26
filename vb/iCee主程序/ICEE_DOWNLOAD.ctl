VERSION 5.00
Begin VB.UserControl ICEE_DOWNLOAD 
   BackColor       =   &H00000000&
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1830
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00000000&
   ScaleHeight     =   1560
   ScaleWidth      =   1830
   ToolboxBitmap   =   "ICEE_DOWNLOAD.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ICEE_DOWNLOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'

Option Explicit

Event DownloadProgress(Curbytes As Long, Maxbytes As Long, Total As String) '下载进度
Event DownloadComplete(Maxbytes As Long, SaveFile As String) '下载结束
Event Speed(Spe As String, Elapsed As String, Left As String) '下载速度
Event State(DString As String, SaveName As String) '下载状态

Dim Ti As Long '下载时间
Dim DByte As Long '已下载文件数据大小
Dim DMax As Long '数据总大小

Private Sub Timer1_Timer()  '用来激发事件
Ti = Ti + 1
If DByte = 0 Then Exit Sub
RaiseEvent Speed(ByteSize(DByte / Ti), GetRestTime(Ti), GetRestTime(DMax / DByte * Ti - Ti))
RaiseEvent DownloadProgress(DByte, DMax, ByteSize(DMax))
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
   Timer1.Enabled = False
   On Error Resume Next
   Dim f() As Byte, fn As Long
   If AsyncProp.BytesMax <> 0 Then
      fn = FreeFile
      f = AsyncProp.Value
      Open AsyncProp.PropertyName For Binary Access Write As #fn
      Put #fn, , f  '存文件
      Close #fn
      RaiseEvent State("下载完成! 线程退出....", Right(AsyncProp.PropertyName, Len(AsyncProp.PropertyName) - InStrRev(AsyncProp.PropertyName, "\")))
   Else
      RaiseEvent State("下载失败....", Right(AsyncProp.PropertyName, Len(AsyncProp.PropertyName) - InStrRev(AsyncProp.PropertyName, "\")))
   End If
    RaiseEvent DownloadComplete(CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
   On Error Resume Next
   If AsyncProp.BytesMax <> 0 Then
        DByte = AsyncProp.BytesRead
        DMax = AsyncProp.BytesMax
   End If
End Sub

Public Function BeginDownload(URL As String, SaveFile As String) As Boolean
   On Error GoTo ErrorBeginDownload
   RaiseEvent State("提交下载请求....", Right(SaveFile, Len(SaveFile) - InStrRev(SaveFile, "\")))
   UserControl.AsyncRead URL, vbAsyncTypeByteArray, SaveFile, vbAsyncReadForceUpdate
   Timer1.Enabled = True
   BeginDownload = True
   RaiseEvent State("开始下载,向缓存区写数据....", Right(SaveFile, Len(SaveFile) - InStrRev(SaveFile, "\")))
   Exit Function
ErrorBeginDownload:
   BeginDownload = False
   RaiseEvent State("无法开始下载....", Right(SaveFile, Len(SaveFile) - InStrRev(SaveFile, "\")))
End Function

Public Function CloseDownload(SaveFile As String) As Boolean
    On Error GoTo ErrorCloseDownload
    UserControl.CancelAsyncRead SaveFile
    RaiseEvent State("结束下载,线程退出....", Right(SaveFile, Len(SaveFile) - InStrRev(SaveFile, "\")))
    Timer1.Enabled = False
    CloseDownload = True
    Exit Function
ErrorCloseDownload:
    CloseDownload = False
    RaiseEvent State("无法结束下载....", Right(SaveFile, Len(SaveFile) - InStrRev(SaveFile, "\")))
End Function

Private Function GetRestTime(Position As Long) As String
''这个函数的功能是把以长整型表示的时间转换为电子钟式的"**:**:**"
Dim Min As String, Sec As String, Hou As String
Hou = Position \ 360
Min = (Position Mod 360) \ 60
Sec = Position - Hou * 360 - Min * 60
If Len(Hou) < 2 Then Hou = "0" & Hou
If Len(Min) < 2 Then Min = "0" & Min
If Len(Sec) < 2 Then Sec = "0" & Sec
GetRestTime = Hou & ":" & Min & ":" & Sec
End Function


Public Function ByteSize(DoByte As Long) As String
''这个函数功能是转换字节单位
Select Case DoByte
    Case 0 To 1023      'Byte
        ByteSize = DoByte & " Byte"
    Case 1024 To 1048575       'KB
        ByteSize = DoByte \ 1024 & " KB"
    Case 1048576 To 1073741823      'MB
        ByteSize = Round(DoByte / 1024 / 1024, 2) & " MB"
    Case Is > 1073741823       'GB
        ByteSize = Round(DoByte / 1024 / 1024 / 1024, 2) & " GB"
End Select
End Function

