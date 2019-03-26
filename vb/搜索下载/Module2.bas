Attribute VB_Name = "Module2"
Option Explicit

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
   ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const BIF_NEWDIALOGSTYLE = &H40
Const BIF_EDITBOX = &H10
Const BIF_USENEWUI = BIF_NEWDIALOGSTYLE Or BIF_EDITBOX
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(Optional sTitle As String = "ÇëÑ¡ÔñÎÄ¼þ¼Ð") As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = 0 ' Me.hWnd
        .lpszTitle = lstrcat(sTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
       sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
       iNull = InStr(sPath, vbNullChar)
        If iNull Then
          sPath = Left$(sPath, iNull - 1)
        End If
    End If

    BrowseForFolder = sPath
End Function
