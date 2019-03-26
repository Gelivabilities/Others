Attribute VB_Name = "下载模块"
Dim i As Integer
Public AUTO_OPEN_IT As Boolean
Public AUTO_OPEN_FOLDER As Boolean
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Enum EnumFilter
    [No_Filter]
    [Only_Enabled]
    [Only_Visible]
    [Only_Enabled_Visible]
    [Only_Enabled_NonVisible]
    [Only_Disabled_Visible]
    [Only_Disabled_NonVisible]
    [Only_Visible_WinTextNotEmpty]
End Enum

Public EnumCondition As EnumFilter

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                        ByVal lpClassName As String, _
                        ByVal lpWindowName As String _
) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                        ByVal hwnd As Long, _
                        ByVal wMsg As Long, _
                        ByVal wParam As Long, _
                        lParam As Any _
) As Long

Private Declare Function GetWindow Lib "user32" ( _
          ByVal hwnd As Long, _
          ByVal wCmd As Long _
) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
          ByVal hwnd As Long, _
          ByVal lpClassName As String, _
          ByVal nMaxCount As Long _
) As Long

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD
                
Public Const MAX_PATH = 260
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
                        ByVal hKey As Long, _
                        ByVal lpSubKey As String, _
                        phkResult As Long _
) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
                        ByVal hKey As Long, _
                        ByVal lpValueName As String, _
                        ByVal lpReserved As Long, _
                        lpType As Long, _
                        lpData As Any, _
                        lpcbData As Long _
) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
                         ByVal hKey As Long, _
                         ByVal lpValueName As String, _
                         ByVal Reserved As Long, _
                         ByVal dwType As Long, _
                         lpData As Any, _
                         ByVal cbData As Long _
) As Long
Public IEver2 As String
Private Const WM_NCDESTROY = &H82
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const OLDWNDPROC = "OldWndProc"
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal PV As Long)

Public Const NOERROR = 0
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                              (ByVal hwndOwner As Long, _
                              ByVal nFolder As SHSpecialFolderIDs, _
                              pidl As Long) As Long

Public Enum SHSpecialFolderIDs
  CSIDL_DESKTOP = &H0
  CSIDL_INTERNET = &H1
  CSIDL_PROGRAMS = &H2
  CSIDL_CONTROLS = &H3
  CSIDL_PRINTERS = &H4
  CSIDL_PERSONAL = &H5
  CSIDL_FAVORITES = &H6
  CSIDL_STARTUP = &H7
  CSIDL_RECENT = &H8
  CSIDL_SENDTO = &H9
  CSIDL_BITBUCKET = &HA
  CSIDL_STARTMENU = &HB
  CSIDL_DESKTOPDIRECTORY = &H10
  CSIDL_DRIVES = &H11
  CSIDL_NETWORK = &H12
  CSIDL_NETHOOD = &H13
  CSIDL_FONTS = &H14
  CSIDL_TEMPLATES = &H15
  CSIDL_COMMON_STARTMENU = &H16
  CSIDL_COMMON_PROGRAMS = &H17
  CSIDL_COMMON_STARTUP = &H18
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19
  CSIDL_APPDATA = &H1A
  CSIDL_PRINTHOOD = &H1B
  CSIDL_ALTSTARTUP = &H1D                      ' ' DBCS
  CSIDL_COMMON_ALTSTARTUP = &H1E    ' ' DBCS
  CSIDL_COMMON_FAVORITES = &H1F
  CSIDL_INTERNET_CACHE = &H20
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22
End Enum

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pidl As Long, _
                              ByVal pszPath As String) As Long

Declare Function SHGetFileInfoPidl Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pidl As Long, _
                              ByVal dwFileAttributes As Long, _
                              psfib As SHFILEINFOBYTE, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As SHGFI_flags) As Long

Public Type SHFILEINFOBYTE   ' sfib
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName(1 To MAX_PATH) As Byte
  szTypeName(1 To 80) As Byte
End Type

Public Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

Enum SHGFI_flags
  SHGFI_LARGEICON = &H0             ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1             ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2               ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                          ' pszPath is pidl, rtns BOOL
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretent pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                     ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200     ' isf.szDisplayName is filled, rtns BOOL
  SHGFI_TYPENAME = &H400           ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                             ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000              ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000     ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000         ' sfi.hIcon is selected icon
End Enum
Private m_hSHNotify As Long      ' the one and only shell change notification handle for the desktop folder
Private m_pidlDesktop As Long   ' the desktop's pidl
Public Const WM_SHNOTIFY = &H401
Public Type PIDLSTRUCT
  pidl As Long
  bWatchSubFolders As Long
End Type
Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" _
                              (ByVal hwnd As Long, _
                              ByVal uFlags As SHCN_ItemFlags, _
                              ByVal dwEventID As SHCN_EventIDs, _
                              ByVal uMsg As Long, _
                              ByVal cItems As Long, _
                              lpps As PIDLSTRUCT) As Long
Type SHNOTIFYSTRUCT
  dwItem1 As Long
  dwItem2 As Long
End Type

Declare Function SHChangeNotifyDeregister Lib "shell32" Alias "#4" (ByVal hNotify As Long) As Boolean
Declare Sub SHChangeNotify Lib "shell32" _
                        (ByVal wEventId As SHCN_EventIDs, _
                        ByVal uFlags As SHCN_ItemFlags, _
                        ByVal dwItem1 As Long, _
                        ByVal dwItem2 As Long)
Public Enum SHCN_EventIDs
  SHCNE_RENAMEITEM = &H1      ' (D) A nonfolder item has been renamed.
  SHCNE_CREATE = &H2                ' (D) A nonfolder item has been created.
  SHCNE_DELETE = &H4                ' (D) A nonfolder item has been deleted.
  SHCNE_MKDIR = &H8                  ' (D) A folder item has been created.
  SHCNE_RMDIR = &H10                ' (D) A folder item has been removed.
  SHCNE_MEDIAINSERTED = &H20     ' (G) Storage media has been inserted into a drive.
  SHCNE_MEDIAREMOVED = &H40      ' (G) Storage media has been removed from a drive.
  SHCNE_DRIVEREMOVED = &H80      ' (G) A drive has been removed.
  SHCNE_DRIVEADD = &H100              ' (G) A drive has been added.
  SHCNE_NETSHARE = &H200             ' A folder on the local computer is being shared via the network.
  SHCNE_NETUNSHARE = &H400        ' A folder on the local computer is no longer being shared via the network.
  SHCNE_ATTRIBUTES = &H800           ' (D) The attributes of an item or folder have changed.
  SHCNE_UPDATEDIR = &H1000          ' (D) The contents of an existing folder have changed, but the folder still exists and has not been renamed.
  SHCNE_UPDATEITEM = &H2000                  ' (D) An existing nonfolder item has changed, but the item still exists and has not been renamed.
  SHCNE_SERVERDISCONNECT = &H4000   ' The computer has disconnected from a server.
  SHCNE_UPDATEIMAGE = &H8000&              ' (G) An image in the system image list has changed.
  SHCNE_DRIVEADDGUI = &H10000               ' (G) A drive has been added and the shell should create a new window for the drive.
  SHCNE_RENAMEFOLDER = &H20000          ' (D) The name of a folder has changed.
  SHCNE_FREESPACE = &H40000                   ' (G) The amount of free space on a drive has changed.

#If (WIN32_IE >= &H400) Then
  SHCNE_EXTENDED_EVENT = &H4000000   ' (G) Not currently used.
#End If     ' WIN32_IE >= &H0400

  SHCNE_ASSOCCHANGED = &H8000000       ' (G) A file type association has changed.

  SHCNE_DISKEVENTS = &H2381F                  ' Specifies a combination of all of the disk event identifiers. (D)
  SHCNE_GLOBALEVENTS = &HC0581E0        ' Specifies a combination of all of the global event identifiers. (G)
  SHCNE_ALLEVENTS = &H7FFFFFFF
  SHCNE_INTERRUPT = &H80000000              ' The specified event occurred as a result of a system interrupt.
                                                                            ' It is stripped out before the clients of SHCNNotify_ see it.
End Enum

#If (WIN32_IE >= &H400) Then   ' ...
 Public Const SHCNEE_ORDERCHANGED = &H2    ' dwItem2 is the pidl of the changed folder
#End If
Public Enum SHCN_ItemFlags
  SHCNF_IDLIST = &H0                ' LPITEMIDLIST
  SHCNF_PATHA = &H1               ' path name
  SHCNF_PRINTERA = &H2         ' printer friendly name
  SHCNF_DWORD = &H3             ' DWORD
  SHCNF_PATHW = &H5              ' path name
  SHCNF_PRINTERW = &H6        ' printer friendly name
  SHCNF_TYPE = &HFF
  SHCNF_FLUSH = &H1000
  SHCNF_FLUSHNOWAIT = &H2000

#If UNICODE Then
  SHCNF_PATH = SHCNF_PATHW
  SHCNF_PRINTER = SHCNF_PRINTERW
#Else
  SHCNF_PATH = SHCNF_PATHA
  SHCNF_PRINTER = SHCNF_PRINTERA
#End If
End Enum
'

' Registers the one and only shell change notification.

Public Function SHNotify_Register(hwnd As Long) As Boolean
  Dim PS As PIDLSTRUCT
  
  ' If we don't already have a notification going...
  If (m_hSHNotify = 0) Then
  
    ' Get the pidl for the desktop folder.
    m_pidlDesktop = GetPIDLFromFolderID(0, CSIDL_DESKTOP)
    If m_pidlDesktop Then
      
      ' Fill the one and only PIDLSTRUCT, we're watching
      ' desktop and all of the it's subfolders, everything...
      PS.pidl = m_pidlDesktop
      PS.bWatchSubFolders = True
      
      ' Register the notification, specifying that we want the dwItem1 and dwItem2
      ' members of the SHNOTIFYSTRUCT to be pidls. We're watching all events.
      m_hSHNotify = SHChangeNotifyRegister(hwnd, SHCNF_TYPE Or SHCNF_IDLIST, _
                                            SHCNE_ALLEVENTS Or SHCNE_INTERRUPT, _
                                            WM_SHNOTIFY, 1, PS)
      Debug.Print Hex(SHCNF_TYPE Or SHCNF_IDLIST)
      Debug.Print Hex(SHCNE_ALLEVENTS Or SHCNE_INTERRUPT)
      Debug.Print m_hSHNotify
      SHNotify_Register = CBool(m_hSHNotify)
    
    Else
      ' If something went wrong...
      Call CoTaskMemFree(m_pidlDesktop)
    
    End If   ' m_pidlDesktop
  End If   ' (m_hSHNotify = 0)
  
End Function

' Unregisters the one and only shell change notification.

Public Function SHNotify_Unregister() As Boolean
  
  ' If we have a registered notification handle.
  If m_hSHNotify Then
    ' Unregister it. If the call is successful, zero the handle's variable,
    ' free and zero the the desktop's pidl.
    If SHChangeNotifyDeregister(m_hSHNotify) Then
      m_hSHNotify = 0
      Call CoTaskMemFree(m_pidlDesktop)
      m_pidlDesktop = 0
      SHNotify_Unregister = True
    End If
  End If

End Function

' Returns the event string associated with the specified event ID value.

Public Function SHNotify_GetEventStr(dwEventID As Long) As String
  Dim sEvent As String
  
  Select Case dwEventID
    Case SHCNE_RENAMEITEM: sEvent = "重命名文件"   ' = &H1"
    Case SHCNE_CREATE: sEvent = "创建文件"   ' = &H2"
    Case SHCNE_DELETE: sEvent = "删除文件"   ' = &H4"
    Case SHCNE_MKDIR: sEvent = "创建文件夹"   ' = &H8"
    Case SHCNE_RMDIR: sEvent = "删除文件夹"   ' = &H10"
    Case SHCNE_MEDIAINSERTED: sEvent = "发现可移动存储设备"   ' = &H20"
    Case SHCNE_MEDIAREMOVED: sEvent = "移除可移动设备"   ' = &H40"
    Case SHCNE_DRIVEREMOVED: sEvent = "删除驱动"   ' = &H80"
    Case SHCNE_DRIVEADD: sEvent = "加入驱动"   ' = &H100"
    Case SHCNE_NETSHARE: sEvent = "网络共享"   ' = &H200"
    Case SHCNE_NETUNSHARE: sEvent = "取消网络共享"   ' = &H400"
    Case SHCNE_ATTRIBUTES: sEvent = "'改变文件目录属性/文件名"   ' = &H800"
    Case SHCNE_UPDATEDIR: sEvent = "更新文件夹"   ' = &H1000"
    Case SHCNE_UPDATEITEM: sEvent = "更新文件/文件名"   ' = &H2000"
    Case SHCNE_SERVERDISCONNECT: sEvent = "断开与服务器的连接"   ' = &H4000"
    Case SHCNE_UPDATEIMAGE: sEvent = "SHCNE_UPDATEIMAGE"   ' = &H8000&"
    Case SHCNE_DRIVEADDGUI: sEvent = "SHCNE_DRIVEADDGUI"   ' = &H10000"
    Case SHCNE_RENAMEFOLDER: sEvent = "重命名文件夹"   ' = &H20000"
    Case SHCNE_FREESPACE: sEvent = "磁盘空间大小改变"   ' = &H40000"
    
#If (WIN32_IE >= &H400) Then
    Case SHCNE_EXTENDED_EVENT: sEvent = "SHCNE_EXTENDED_EVENT"   ' = &H4000000"
#End If     ' WIN32_IE >= &H0400
    
    Case SHCNE_ASSOCCHANGED: sEvent = "SHCNE_ASSOCCHANGED"   ' = &H8000000"
    
    Case SHCNE_DISKEVENTS: sEvent = "SHCNE_DISKEVENTS"   ' = &H2381F"
    Case SHCNE_GLOBALEVENTS: sEvent = "SHCNE_GLOBALEVENTS"   ' = &HC0581E0"
    Case SHCNE_ALLEVENTS: sEvent = "SHCNE_ALLEVENTS"   ' = &H7FFFFFFF"
    Case SHCNE_INTERRUPT: sEvent = "SHCNE_INTERRUPT"   ' = &H80000000"
  End Select
  
  SHNotify_GetEventStr = sEvent

End Function
Public Function GetPIDLFromFolderID(hOwner As Long, nFolder As SHSpecialFolderIDs) As Long
  Dim pidl As Long
  If SHGetSpecialFolderLocation(hOwner, nFolder, pidl) = NOERROR Then
    GetPIDLFromFolderID = pidl
  End If
End Function
Public Function GetDisplayNameFromPIDL(pidl As Long) As String
  Dim sfib As SHFILEINFOBYTE
  If SHGetFileInfoPidl(pidl, 0, sfib, Len(sfib), SHGFI_PIDL Or SHGFI_DISPLAYNAME) Then
    GetDisplayNameFromPIDL = GetStrFromBufferA(StrConv(sfib.szDisplayName, vbUnicode))
  End If
End Function

' Returns a path from only an absolute pidl (relative to the desktop)

Public Function GetPathFromPIDL(pidl As Long) As String
  Dim SPATH As String * MAX_PATH
  If SHGetPathFromIDList(pidl, SPATH) Then   ' rtns TRUE (1) if successful, FALSE (0) if not
    GetPathFromPIDL = GetStrFromBufferA(SPATH)
  End If
End Function
Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    GetStrFromBufferA = sz
  End If
End Function

Public Function FSubClass(hwnd As Long) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  
  If (GetProp(hwnd, OLDWNDPROC) = 0) Then
    lpfnOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)
    If lpfnOld Then
      fSuccess = SetProp(hwnd, OLDWNDPROC, lpfnOld)
    End If
  End If
  
  If fSuccess Then
    FSubClass = True
  Else
    If lpfnOld Then Call UnSubClass(hwnd)
  End If
  
End Function

Public Function UnFSubClass(hwnd As Long) As Boolean
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hwnd, OLDWNDPROC)
  If lpfnOld Then
    If RemoveProp(hwnd, OLDWNDPROC) Then
      UnFSubClass = SetWindowLong(hwnd, GWL_WNDPROC, lpfnOld)
    End If
  End If

End Function

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg
    Case WM_SHNOTIFY
      Call FRMEND.NotificationReceipt(wParam, lParam)
      
    Case WM_NCDESTROY
      Call UnSubClass(hwnd)
  End Select
  
  WndProc = CallWindowProc(GetProp(hwnd, OLDWNDPROC), hwnd, uMsg, wParam, lParam)
End Function
Public Function GetURL() As String
Dim sIEClassName     As String, hIE       As Long, lngRep       As Long
Dim sText     As String * 255, sClass           As String * 255
Dim iNum     As Long, hwndChild       As Long, lngRepClassName       As Long
Dim lngLength     As Long, sURL       As String
    
  Dim a(1 To 7) As String
  
    a(1) = "IEFrame"
    a(2) = "WorkerW"
    a(3) = "ReBarWindow32"
    a(4) = "Address Band Root"
    a(5) = "ComboBoxEx32"
    a(6) = "ComboBox"
    a(7) = "Edit"
    
On Error GoTo Fin
sIEClassName = a(1)
hIE = FindWindow(sIEClassName, vbNullString)
If hIE <> 0 Then
          hwndChild = hIE
          hwndChild = hwndFindWindow(hwndChild, a(2))
          If hwndChild = 0 Then ERR.Raise 10
          hwndChild = hwndFindWindow(hwndChild, a(3))
          If hwndChild = 0 Then ERR.Raise 10
          '判断IE版本
          If IEver2 = 7 Then
              hwndChild = hwndFindWindow(hwndChild, a(4))
              If hwndChild = 0 Then ERR.Raise 10
          End If
          hwndChild = hwndFindWindow(hwndChild, a(5))
          If hwndChild = 0 Then ERR.Raise 10
          hwndChild = hwndFindWindow(hwndChild, a(6))
          If hwndChild = 0 Then ERR.Raise 10
          hwndChild = hwndFindWindow(hwndChild, a(7))
          If hwndChild = 0 Then ERR.Raise 10
          GetURL = ExtractURL(hwndChild)
End If
Exit Function
Fin:
End Function
Public Sub BeginDown(URL As String, Dpath As String, Dname As String)
On Error Resume Next
If FRMDOWN.Downloader1(i).BeginDownload(URL, Dpath & Dname) = True Then
    Call NewDown
    FRMDOWN.LVIEW.ListItems.Add i, , Dname, , 2
    FRMDOWN.LVIEW.ListItems(i).SubItems(7) = Frmadd.Text1.Text
    FRMDOWN.LVIEW.ListItems(i).SubItems(4) = Dpath
    FRMDOWN.LVIEW.ListItems(i).SubItems(2) = "0 %"
End If
End Sub

Private Sub NewDown()
i = i + 1
Load FRMDOWN.Downloader1(i)
End Sub

Private Sub CloseDown()
i = i - 1
Unload FRMDOWN.Downloader1(i + 1)
End Sub

Public Function YesNoUrl(URL As String) As Boolean
If LCase(Left(URL, 7)) = "http://" Or LCase(Left(URL, 6)) = "ftp://" Or LCase(Left(URL, 8)) = "https://" Then
    If InStr(1, URL, ".") >= InStr(1, URL, "//") + 3 Then
        YesNoUrl = True
        Exit Function
    End If
End If
YesNoUrl = False
End Function

Private Function SupprimeNull(sM As String) As String
If (InStr(sM, Chr(0)) > 0) Then
        sM = Left(sM, InStr(sM, Chr(0)) - 1)
End If
SupprimeNull = sM
End Function

Private Function ExtractURL(hwnd As Long) As String
Dim lngLength     As Long, sURL       As String, lngRep       As Long

lngLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, ByVal 0)
sURL = Space(lngLength + 1)
lngRep = SendMessage(hwnd, WM_GETTEXT, lngLength + 1, ByVal sURL)
ExtractURL = SupprimeNull(sURL)
End Function


Private Function hwndFindWindow(hwndParent As Long, sClassName As String) As Long
Dim hwndChild     As Long, sClass       As String * MAX_PATH
Dim bTrouve     As Boolean, lngRepClassName       As String

hwndChild = GetWindow(hwndParent, GW_CHILD)
lngRepClassName = GetClassName(hwndChild, sClass, 255)
If Left(sClass, lngRepClassName) = sClassName Then
          hwndFindWindow = hwndChild
          Exit Function
End If
If hwndChild = 0 Then Exit Function

bTrouve = False
Do Until bTrouve
          hwndChild = GetWindow(hwndChild, GW_HWNDNEXT)
          If hwndChild = 0 Then Exit Do
          lngRepClassName = GetClassName(hwndChild, sClass, MAX_PATH)
          If Left(sClass, lngRepClassName) = sClassName Then
                  hwndFindWindow = hwndChild
                  Exit Function
          End If
Loop
End Function

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    Dim sWinText As String
    Dim lngWinTextLen As Long
    Dim bIsWin As Boolean
    Dim bIsVisible As Boolean
    Dim bIsEnabled As Boolean
    Dim lstItem As ListItem
    Dim lstSubItem As ListSubItem
    Dim lstImage As Integer
    
       
    '返回窗口句柄
    If IsWindow(hwnd) = 0 Then bIsWin = False Else bIsWin = True
    
    '窗口是否可见
    If bIsWin = True Then
        If IsWindowVisible(hwnd) = 0 Then bIsVisible = False Else bIsVisible = True
    
    '窗口是否激活
        If IsWindowEnabled(hwnd) = 0 Then bIsEnabled = False Else bIsEnabled = True
    
    '获取窗口文字长度 ...
        lngWinTextLen = GetWindowTextLength(hwnd)
    
    '获取窗口文字 ...
        sWinText = Space(lngWinTextLen)
        GetWindowText hwnd, sWinText, lngWinTextLen + 1
    
    End If
    
    
    With FRMEND
    .lstWinList.SmallIcons = .ImageList1.Object
    .lstWinList.Icons = .ImageList1.Object
        If bIsWin = True Then
            If bIsEnabled = True Then
                If bIsVisible = True Then
                    lstImage = 5
                Else
                    lstImage = 4
                End If
            Else
                lstImage = 3
            End If
        End If
    
        '列表顺序 ...
        '1 : 窗口句柄
        '2 : 窗口文字
        '3 : 是否可见
        '4 : 是否激活
        
        

       
       Select Case EnumCondition
        
            Case No_Filter:
                Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                With lstItem.ListSubItems
    
                    If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                    Else
                        .Add , , "- NA -"
                    End If
                    
                    If bIsVisible = True Then
                        .Add , , "可见"
                        lstItem.FOREColor = vbRed
                        lstItem.Bold = True
                    Else
                        .Add , , "不可见"
                    End If
                    
                    If bIsEnabled = True Then
                        .Add , , "激活"
                    Else
                        .Add , , "未激活"
                    End If
                
                End With
            
            Case Only_Visible
                If bIsVisible = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
    
                        If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                        Else
                          .Add , , "- NA -"
                        End If
                        
                        .Add , , "可见"
                        lstItem.FOREColor = vbRed
                        lstItem.Bold = True
                        
                        If bIsEnabled = True Then
                            .Add , , "激活"
                        Else
                            .Add , , "未激活"
                        End If
                    End With
                End If
            
            Case Only_Enabled
                If bIsEnabled = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                        Else
                          .Add , , "- NA -"
                        End If
                        
                        If bIsVisible = True Then
                            .Add , , "可见"
                            lstItem.FOREColor = vbRed
                            lstItem.Bold = True
                        Else
                            .Add , , "不可见"
                        End If
                        .Add , , "激活"
                    End With
                End If
                
            Case Only_Visible_WinTextNotEmpty
                If bIsVisible = True And Trim(sWinText) <> "" Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        .Add , , sWinText
                        .Add , , "可见"
                        lstItem.FOREColor = vbRed
                        lstItem.Bold = True
                        
                        If bIsEnabled = True Then
                            .Add , , "激活"
                        Else
                            .Add , , "未激活"
                        End If
                    End With
                End If
                
            Case Only_Enabled_Visible
                If bIsEnabled = True And bIsVisible = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                        Else
                          .Add , , "- NA -"
                        End If
                        .Add , , "可见"
                        lstItem.FOREColor = vbRed
                        lstItem.Bold = True
                        .Add , , "激活"
                    End With
                End If
            
            Case Only_Enabled_NonVisible
                If bIsEnabled = True And bIsVisible = False Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                            .Add , , sWinText
                        Else
                            .Add , , "- NA -"
                        End If
                        .Add , , "不可见"
                        .Add , , "激活"
                    End With
                End If
                
            Case Only_Disabled_NonVisible
                If bIsEnabled = False And bIsVisible = False Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                            .Add , , sWinText
                        Else
                             .Add , , "- NA -"
                        End If
                        .Add , , "不可见"
                        .Add , , "未激活"
                    End With
                End If
            
            Case Only_Disabled_Visible
                If bIsEnabled = False And bIsVisible = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                            .Add , , sWinText
                        Else
                            .Add , , "- NA -"
                        End If
                        .Add , , "可见"
                        lstItem.FOREColor = vbRed
                        lstItem.Bold = True
                        .Add , , "未激活"
                    End With
                End If
        End Select
    End With
        
    '继续同样的过程 ...
    EnumWindowsProc = True
End Function

Public Function GetWinInfo()
    '清除 ListView 存在的...
    FRMEND.lstWinList.ListItems.Clear
    '呼叫 EnumWindowsProc ...
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
End Function

