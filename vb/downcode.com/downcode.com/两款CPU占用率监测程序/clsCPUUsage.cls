VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPUUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PdhVbOpenQuery Lib "PDH.DLL" (ByRef QueryHandle As Long) As Long
Private Declare Function PdhCloseQuery Lib "PDH.DLL" (ByVal hQuery As Long) As Long
Private Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As Long
Private Declare Function PdhRemoveCounter Lib "PDH.DLL" (ByVal CounterHandle As Long) As Long
Private Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As Long
Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Private mhStatKey As Long
Private mhQuery As Long
Private mhCounter As Long
Private Sub Class_Initialize()
Dim lData As Long
Dim hKey As Long
If mIsNT Then
Call PdhVbOpenQuery(mhQuery)
Call PdhVbAddCounter(mhQuery, "\Processor(0)\% Processor Time", mhCounter)
Else
Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
Call RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, REG_DWORD, lData, LenB(lData))
Call RegCloseKey(hKey)
Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", mhStatKey)
End If
lData = Me.Usage
End Sub
Public Property Get Usage() As Long
Dim ret As Long
Dim lData As Long
Dim hKey As Long
If mIsNT Then
Call PdhCollectQueryData(mhQuery)
ret = PdhVbGetDoubleCounterValue(mhCounter, lData)
Else
Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
Call RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, REG_DWORD, lData, LenB(lData))
End If
Usage = ret
End Property
Private Sub Class_Terminate()
Dim lData As Long
Dim hKey As Long
If mIsNT Then
Call PdhRemoveCounter(mhCounter)
Call PdhCloseQuery(mhQuery)
Else
Call RegOpenKey(HKEY_DYN_DATA, "PerfStats\StopStat", hKey)        '9x
Call RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, REG_DWORD, lData, LenB(lData))
Call RegCloseKey(hKey)
End If
End Sub
