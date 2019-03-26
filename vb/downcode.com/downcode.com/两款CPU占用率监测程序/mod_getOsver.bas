Attribute VB_Name = "mod_getOsver"
Option Explicit
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32_NT = 2
Public Function mIsNT() As Boolean
Dim vi As OSVERSIONINFO
vi.dwOSVersionInfoSize = Len(vi)
Call GetVersionEx(vi)
mIsNT = (vi.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function
