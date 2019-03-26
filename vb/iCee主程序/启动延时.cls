VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWaitableTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'延时系统
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&
Private Const INFINITE = &HFFFF
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Const UNITS = 4294967296#
Private Const MAX_LONG = -2147483648#

Private Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Private Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long

Private mlTimer As Long
Public Function WriteP(ByRef pc() As Byte, PIC As StdPicture)
'VB鼠标画圆并用数组保存每次操作，做到撤销，重做
Dim pbag As New PropertyBag
pbag.WriteProperty "pic", PIC
pc = pbag.Contents
End Function
Public Function ReadP(ByRef rc() As Byte)
Dim pbagb As New PropertyBag
pbagb.Contents = rc
Set ReadP = pbagb.ReadProperty("pic")
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    If mlTimer <> 0 Then CloseHandle mlTimer
End Sub
Public Sub Wait(MilliSeconds As Long)
    On Error GoTo ErrHandler
    Dim Ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    
    mlTimer = CreateWaitableTimer(0, True, App.exename & "Timer" & Format$(Now(), "NNSS"))
    
    If ERR.LastDllError <> ERROR_ALREADY_EXISTS Then
        Ft.dwLowDateTime = -1
        Ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(mlTimer, Ft, 0, 0, 0, 0)
    End If
    dblDelay = CDbl(MilliSeconds) * 10000#
    Ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))
    
    If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow
    
    Ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, Ft, 0, 0, 0, False)
    
    Do
        lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0
    CloseHandle mlTimer
    mlTimer = 0
    Exit Sub
    
ErrHandler:
    ERR.Raise ERR.Number, ERR.Source, "[clsWaitableTimer.Wait]" & ERR.Description
End Sub
