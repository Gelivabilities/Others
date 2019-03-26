VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FRMEX 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000BC8D7&
   BorderStyle     =   0  'None
   Caption         =   "资源管理"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   954
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.PictureBox PINFO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   11160
      ScaleHeight     =   257
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   14
      Top             =   5040
      Width           =   3015
      Begin VB.PictureBox PICPR 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   41
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picSmall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0047491F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2040
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   240
      End
      Begin VB.PictureBox picLarge 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0047491F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   2400
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   29
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lbldatetxt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "文件大小:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   810
      End
      Begin VB.Label lbldate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   600
         TabIndex        =   27
         Top             =   1680
         Width           =   1710
      End
      Begin VB.Label lbldatetxt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "最后一次访问时间:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1530
      End
      Begin VB.Label lbldate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   25
         Top             =   1200
         Width           =   1710
      End
      Begin VB.Label lbldatetxt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "最后一次修改时间:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lbldatetxt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "创建时间:"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   810
      End
      Begin VB.Label lbldate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000 00:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   22
         Top             =   720
         Width           =   1710
      End
      Begin VB.Label lbldate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0 MB"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   600
         TabIndex        =   21
         Top             =   240
         Width           =   360
      End
   End
   Begin ICEE.ICEE_KEY Cmd_Download 
      Height          =   495
      Left            =   105
      TabIndex        =   11
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9840
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   1440
      Left            =   105
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2540
      _Version        =   393217
      Style           =   3
      ImageList       =   "Imt_Tree"
      Appearance      =   0
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   13545
      Picture         =   "FRMEX.frx":0000
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   7
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   13545
      Picture         =   "FRMEX.frx":00E4
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   6
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   13545
      Picture         =   "FRMEX.frx":01C8
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox Pic_FileICO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   -360
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   8355
      Width           =   255
   End
   Begin MSComctlLib.ImageList Imt_LV 
      Left            =   480
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList Imt_Tree 
      Left            =   480
      Top             =   3870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEX.frx":02AC
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEX.frx":059E
            Key             =   "disk"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMEX.frx":08FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7305
      Left            =   2895
      TabIndex        =   2
      Top             =   1560
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   12885
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6825
      Left            =   105
      TabIndex        =   3
      Top             =   2025
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   12039
      _Version        =   393217
      Style           =   7
      ImageList       =   "Imt_Tree"
      Appearance      =   0
   End
   Begin VB.TextBox Txt_Address 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4800
      TabIndex        =   4
      Text            =   "SSS "
      Top             =   1080
      Width           =   8145
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   17
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   2
      Left            =   2760
      TabIndex        =   18
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin ICEE.ICEE_KEY ICM 
      Height          =   375
      Index           =   3
      Left            =   12960
      TabIndex        =   19
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.PictureBox PVIEW 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   11160
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Image PV 
         Height          =   1335
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
   End
   Begin ICEE.IMUSIC IMS 
      Height          =   3015
      Left            =   11160
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5318
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地址:"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   3
      Left            =   4320
      TabIndex        =   20
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件信息"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   11160
      TabIndex        =   12
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label LA 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   13680
      TabIndex        =   9
      Top             =   8880
      Width           =   540
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件管理"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "FRMEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFileName As String    '定义变量存储
Dim bFileFlag As Boolean   '定义变量 在复制粘贴时标识

Dim coname As String       '要拷贝的文件名
Dim copath As String       '要拷贝的文件名加全路径

Dim P_ofso
Private Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long _
                                                                             ) As Long

Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Private Const SHGFI_LARGEICON = &H0        ' Large icon
Private Const SHGFI_SMALLICON = &H1        ' Small icon
Private Const ILD_TRANSPARENT = &H1        ' Display transparent
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE _
                                 Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME _
                                 Or SHGFI_EXETYPE

Private Type SHFILEINFO
    hIcon As Long                           '文件的图标句柄
    iIcon As Long                           '图标的系统索引号
    dwAttributes As Long                    '文件的属性
    szDisplayName As String * MAX_PATH      '文件的显示名
    szTypeName As String * 80               '文件的类型名
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As SHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long _
                                                                                 ) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, _
                                                            ByVal I&, _
                                                            ByVal hDCDest&, _
                                                            ByVal x&, _
                                                            ByVal y&, _
                                                            ByVal flags& _
                                                          ) As Long

Private shinfo As SHFILEINFO

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long  '删除到回收站

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

'*************************************************************************
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'清空回收站
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
'Download by http://www.codefans.net
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Private Const SHERB_NOCONFIRMATION = &H1
Private Const SHERB_NOPROGRESSUI = &H2
Private Const SHERB_NOSOUND = &H4
'*************************************************************************

'引用一个Stripting Runtime 对象
Private fs As FileSystemObject
Private strs As String
Private strss As String
Private Comes2 As String        '记录缓存值
Private StrNums As Integer
Private SubFileName As String  '记录创建子文件夹的名称
Private FPaths As String
Private addr As String
Private File1Pattern As String
Private filname As String
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function movefile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST  'idl
    mkid As SHITEMID
End Type
'========================================================
'声明打开文件属性窗口中的API函数
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
'声明 API 函数

Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias _
"ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Const OFS_MAXPATHNAME = 128
Private Const OF_READ = &H0

Private Type SYSTEMTIME
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
     Bias As Long
     StandardName(32) As Integer
     StandardDate As SYSTEMTIME
     StandardBias As Long
     DaylightName(32) As Integer
     DaylightDate As SYSTEMTIME
     DaylightBias As Long
End Type

Private Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type

Private Type BY_HANDLE_FILE_INFORMATION
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     dwVolumeSerialNumber As Long
     nFileSizeHigh As Long
     nFileSizeLow As Long
     nNumberOfLinks As Long
     nFileIndexHigh As Long
     nFileIndexLow As Long
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, _
                                                             lpFileInformation As BY_HANDLE_FILE_INFORMATION _
                                                            ) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                                           lpReOpenBuff As OFSTRUCT, _
                                           ByVal wStyle As Long _
                                         ) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long


Private Type SHELLEXECUTEINFO '这可以VB自代的API帮助中找到
    CBSIZE As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const MaxLFNPath = 260
Private Const LB_INITSTORAGE = &H1A8
Private Const LB_ADDSTRING = &H180
Private Const WM_SETREDRAW = &HB
Private Const WM_VSCROLL = &H115
Private Const SB_BOTTOM = 7
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private lhwnd As String
Private dirs, Dir$, files As Long
Private isrun As Boolean
Private WFD As WIN32_FIND_DATA, hItem&, hFile&
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MaxLFNPath
    cShortFileName As String * 14
End Type

Private Function Findfile(xstrfilename) As WIN32_FIND_DATA
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long

    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data)
    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error"
    Else
        Findfile = Win32Data
    End If
    plngRtn = FindClose(plngFirstFileHwnd)
End Function


'========================================================
'使用此函数可以出现一个对话框，
'并返回所选择的路径，若没有选择返回("").
Function FPath$(nhwnd&, Title$)
    Dim bi As BROWSEINFO
    Dim idl As ITEMIDLIST
    Dim rtn&, pidl&, Path$, pos%
    bi.hOwner = nhwnd&
    bi.pidlRoot = idl.mkid.cb
    bi.lpszTitle = Title$
    bi.ulFlags = &H1
    pidl& = SHBrowseForFolder(bi)
    Path$ = Space$(512)
    rtn& = SHGetPathFromIDList(ByVal pidl&, ByVal Path$)
    pos% = InStr(Path$, Chr$(0))
    ''
    FPath$ = Left(Path$, pos - 1)
End Function
'========================================================


Sub PropsShow(filename As String)    '显示属性窗口自定义函数
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .CBSIZE = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = Me.hwnd
        .lpVerb = "properties"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShellExecuteEX SEI
End Sub


Private Sub Delfolders(strpath)
    Set fs = CreateObject("scripting.filesystemobject")
    On Error GoTo x:
    If fs.FolderExists(strpath) Then
        fs.DeleteFolder (Trim(strpath))
        Call SHOWWRONG("文件夹删除成功!", 2)
    End If
    Exit Sub
x:
    Call SHOWWRONG("错误号:" & ERR.Number & vbCrLf & "错误描述:" & ERR.Description, 0)
End Sub


Private Sub Movfolders(soupath, despath)
    Set fs = CreateObject("scripting.filesystemobject")
    On Error GoTo x:
    If fs.FolderExists(soupath) Then
        fs.CopyFolder soupath, despath
        fs.DeleteFolder (soupath)
        Call SHOWWRONG("已将文件夹移动到 " & despath, 2)
    End If
    Exit Sub
x:
  Call SHOWWRONG("错误号:" & ERR.Number & vbCrLf & "错误描述:" & ERR.Description, 0)
End Sub

Private Sub CreateFolders(strss)        '创建ManageFile文件夹
    Set fs = CreateObject("scripting.filesystemobject")    '创建FSO对象
    On Error GoTo x:
    fs.CreateFolder (Trim(strss))          '使用FSO对象的CreateFolder方法创建文件夹
    Exit Sub
x:
End Sub

Private Sub CreateFiles()        '根据当前日期创建文件夹
    Set fs = CreateObject("scripting.filesystemobject")   '创建FSO对象
    On Error GoTo x:
    strs = str(Date)             '根据系统的当前日期创建文件夹
    strss = App.Path & "\ManageFile\" & strs
    fs.CreateFolder (Trim(strss))
    Exit Sub
x:
End Sub

Private Sub CreateSubFiles()        '创建子文件夹
    Set fs = CreateObject("scripting.filesystemobject")
    On Error GoTo x:
    strs = Comes2 & "文件"            '根据系统的当前日期创建文件夹
    strss = App.Path & "\ManageFile\" & str(Date) & "\" & SubFileName & "\" & strs
    fs.CreateFolder (Trim(strss))
    Exit Sub
x:
End Sub

Private Sub CreateDateSubFiles()        '创建当前日期下的子文件夹
    Set fs = CreateObject("scripting.filesystemobject")
    On Error GoTo x:
    strss = App.Path & "\ManageFile\" & str(Date) & "\" & SubFileName
    fs.CreateFolder (Trim(strss))
    Exit Sub
x:
End Sub

Private Sub NumFiles()            '获得文件夹的个数
    Set fs = CreateObject("scripting.filesystemobject")
    On Error GoTo x:
    strss = App.Path & "\ManageFile\" & str(Date) & "\" & SubFileName
    StrNums = fs.GetFolder(strss).SubFolders.Count
x:
End Sub


Private Sub Cmd_Download_Click()    '显示下拉的TreeView控件
    Dim SPATH As String             '定义变量存储路径
    Dim n  As Integer               '定义变量存储"\"出现的位置
    Dim iMaxCount As Integer        '定义变量用于存储
    Dim sSplit() As String          '定义数组 获取zifu
    Dim I As Integer                '循环变量

    Dim sKey As String

    Dim sText As String           '存储要显示在TreeView控件中的文字内容

    Dim MyFSO As New FileSystemObject
    Dim MyDrive As Drive
        
    If Right(Txt_Address.Text, 1) = "\" Then Txt_Address.Text = Left(Txt_Address, Len(Txt_Address.Text) - 1)
    
    TreeView2.Nodes.Clear
    TreeView2.Nodes.Add , , "root", "我的电脑", 3
    For Each MyDrive In MyFSO.drives
        TreeView2.Nodes.Add "root", tvwChild, MyDrive.DriveLetter, MyDrive.DriveLetter, 2
    Next

    sSplit = Split(Txt_Address.Text, "\")
    iMaxCount = UBound(sSplit)
    n = 1
    For I = 0 To iMaxCount
        SPATH = Left(Txt_Address, InStr(n + 1, Txt_Address.Text, "\"))
        If SPATH = "" Then
            SPATH = Txt_Address.Text & "\"
        End If
        n = InStr(n + 1, Txt_Address.Text, "\")
        sText = Right(Left(SPATH, Len(SPATH) - 1), Len(Left(SPATH, Len(SPATH) - 1)) - InStrRev(Left(SPATH, Len(SPATH) - 1), "\", -1, vbTextCompare))
        If I = 1 Then
            TreeView2.Nodes.Add UCase(Left(SPATH, 1)), tvwChild, SPATH, sText, 1
        ElseIf I > 1 Then
            If sText = "" Then Exit Sub
            TreeView2.Nodes.Add sKey, tvwChild, SPATH, sText, 1
        End If
        sKey = SPATH
    Next I
    For I = 1 To TreeView2.Nodes.Count
        TreeView2.Nodes(I).Expanded = True
    Next I
    TreeView2.Visible = True
    TreeView2.SetFocus
End Sub
 
Sub Cmd_Go_Click()       '转到
    Dim SPATH As String          '定义变量存储路径
    Dim n  As Integer            '定义变量存储"\"出现的位置
    Dim iMaxCount As Integer     '定义变量用于存储
    Dim sSplit() As String       '定义数组 获取zifu
    Dim I As Integer             '循环变量

    Dim MyFSO As New FileSystemObject
    Dim MyFolder As Folder
    Dim Folder1 As Folder       '指定文件夹下的子文件夹
  
    Dim sExtension As String    '定义变量存储文件的扩展
    On Error GoTo MyErr
    If Right(Me.Txt_Address.Text, 1) = "\" Then Me.Txt_Address.Text = Left(Me.Txt_Address, Len(Me.Txt_Address) - 1)
    
    sSplit = Split(Txt_Address.Text, "\")
    iMaxCount = UBound(sSplit)
    n = 1
    TreeView1.Nodes.Clear
    For I = 0 To iMaxCount
        SPATH = Left(Txt_Address, InStr(n + 1, Txt_Address.Text, "\"))
        If SPATH = "" Then
            SPATH = Txt_Address.Text & "\"
        End If
        n = InStr(n + 1, Txt_Address.Text, "\")
        If I = 0 Then
            TreeView1.Nodes.Add , , Left(SPATH, Len(SPATH) - 1), Left(SPATH, Len(SPATH) - 1), 2
        Else
            
        End If
        Tree_DataExpanded (SPATH)
        '将文件夹路径指定到当前所选择的路径下
        Set MyFolder = MyFSO.GetFolder(SPATH)
        If TreeView1.Nodes(Left(SPATH, Len(SPATH) - 1)).children = 0 Then
            For Each Folder1 In MyFolder.SubFolders
                TreeView1.Nodes.Add Left(SPATH, Len(SPATH) - 1), tvwChild, SPATH & Folder1.name, Folder1.name, 1
            Next
        End If
        TreeView1.Nodes(Left(SPATH, Len(SPATH) - 1)).Selected = True
        TreeView1.Nodes(Left(SPATH, Len(SPATH) - 1)).Expanded = True
        TreeView1.SetFocus
    Next I
      Drive1.Drive = UCase(Left(Me.Txt_Address.Text, 1))
  Exit Sub
MyErr:
    If ERR.Number = 76 Then
        Call SHOWWRONG("找不到  '" & Txt_Address.Text & "'.请确认地址正确!", 0)
        Exit Sub
    End If
End Sub
Private Sub Form_Activate()
    TreeView1.SetFocus
    Me.BackColor = COLOR_NOR
    PINFO.BackColor = COLOR_NOR
    picLarge.BackColor = COLOR_NOR
    picSmall.BackColor = COLOR_NOR
    Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
 Txt_Address.BackColor = COLOR_NOR
Cmd_Download.SETCOLOR vbWhite, &HDECC5, vbBlack
Dim I As Integer
For I = 0 To ICM.Count - 1
ICM(I).SETCOLOR COLOR_NOR, COLOR_HIGH, vbWhite
ICM(I).HASLINE = False
Next
End Sub

Private Sub Form_Load()
    TreeView1.LineStyle = tvwTreeLines
    TreeView1.PathSeparator = "\"
    ListView1.FullRowSelect = True
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Add , , "名称", ListView1.Width * 30 / 80
    ListView1.ColumnHeaders.Add , , "大小", ListView1.Width * 10 / 80
    ListView1.ColumnHeaders.Add , , "类型", ListView1.Width * 20 / 80
    ListView1.ColumnHeaders.Add , , "修改日期", ListView1.Width * 24 / 80
    ListView1.LabelEdit = lvwManual
    ListView1.GridLines = False      '不显示报表的表格线
    ListView1.ColumnHeaders(2).Alignment = lvwColumnRight   '右对齐

    Tree_Add (Left(UCase(Drive1.Drive), 2) & "\")   '调用自定义过程添加根节点
    TreeView1.Nodes(1).Expanded = True
    Txt_Address.Text = Left(Drive1.Drive, 2) & "\"
    TreeView1.HideSelection = False    '失去焦点时选中效果
    TreeView2.LineStyle = tvwRootLines
    
    ICM(0).SETTXT "新建文件夹"
    ICM(1).SETTXT "重新命名"
    ICM(2).SETTXT "文件属性"
    ICM(3).SETTXT "前往"

Call RE_UI
If LONELY_MODE = True Then Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2: Load Frmm
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If TreeView2.Visible = True Then TreeView2.Visible = False
X1.Visible = True
X2.Visible = False
X3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LONELY_MODE = True Then End
End Sub

Private Sub ICM_Click(Index As Integer)
Select Case Index
Case 0
NewFolder_Click
Case 1
ReName_Click
Case 2
Attribute_Click
Case 3
Call Cmd_Go_Click
End Select
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim I As Integer
    ListView1.Sorted = True
    ListView1.SortKey = ColumnHeader.Index - 1
    If ListView1.SortOrder = lvwDescending Then
        ListView1.SortOrder = lvwAscending
    Else
        ListView1.SortOrder = lvwDescending
    End If
End Sub

Private Sub ListView1_DblClick()    '双击执行或打开文件
    '如果是文件夹展开文件夹，如果是文件则不执行操作”
    Dim MyFSO As New FileSystemObject
    Dim MyFolder As Folder
    Dim Folder1 As Folder   '指定文件夹下的子文件夹

    Dim sKey As String     '定义变量充当TreeView的关键字
    Dim SPATH As String       '定义变量存储路径
    coname = ""
    If ListView1.SelectedItem.SubItems(1) = "" Then    '如果是文件夹
        SPATH = TreeView1.SelectedItem.fullPath & "\" & ListView1.SelectedItem.Text & "\"
        sKey = Left(SPATH, Len(SPATH) - 1)
        '将文件夹路径指定到当前所选择的路径下
        Set MyFolder = MyFSO.GetFolder(SPATH)
        For Each Folder1 In MyFolder.SubFolders
            TreeView1.Nodes.Add sKey, tvwChild, sKey & "\" & Folder1.name, Folder1.name, 1
        Next
        TreeView1.Nodes(sKey).Selected = True
        TreeView1.Nodes(sKey).Expanded = True
        TreeView1.SetFocus
        Tree_DataExpanded (SPATH)    '添加
        bFileFlag = False
    Else             '如果是文件
        If Right(Me.Txt_Address.Text, 1) = "\" Then
            Me.Txt_Address.Text = Left(Me.Txt_Address.Text, Len(Me.Txt_Address.Text) - 1)
        End If
        sFileName = Me.Txt_Address.Text & "\" & ListView1.SelectedItem.Text
        SPATH = Me.Txt_Address.Text
        coname = ListView1.SelectedItem.Text
        bFileFlag = True
        OPEN_CLICK
    End If
    Txt_Address.Text = SPATH
    
    
End Sub
Sub OPEN_CLICK()
    Dim SPATH As String
    If Right(Txt_Address.Text, 1) = "\" Then Txt_Address.Text = Left(Txt_Address.Text, Len(Txt_Address.Text) - 1)
    SPATH = Me.Txt_Address.Text & "\" & coname
    Call ShellExecute(Me.hwnd, "Open", SPATH, vbNullString, App.Path, SW_SHOWNORMAL)    '以文件默认的打开方式打开
End Sub

Sub Attribute_Click()   '显示属性
    PropsShow Txt_Address.Text
End Sub
Sub Copy_Click()   '复制
    If bFileFlag = True Then    '如果是文件
        copath = sFileName
    ElseIf bFileFlag = False Then  '如果是文件夹
        copath = Me.Txt_Address.Text
        coname = ""
    End If
End Sub

Sub Del_Click()   '删除
    Dim n As Integer    '存储要提取的字符串的数量
    If Right(Txt_Address.Text, 1) = "\" Then Txt_Address.Text = Left(Txt_Address.Text, Len(Txt_Address.Text) - 1)
    Delfolders (Txt_Address.Text & "\" & ListView1.SelectedItem.Text)
    n = InStrRev(Me.Txt_Address, "\")
    Me.Txt_Address.Text = Left(Me.Txt_Address.Text, n - 1)
    Cmd_Go_Click
End Sub

Sub NewFolder_Click()   '新建文件夹
    Dim ss As String
    ss = InputBox("输入新建文件夹名称", "新建文件夹", "新建文件夹")
    If Right(Txt_Address.Text, 1) = "\" Then Txt_Address.Text = Left(Txt_Address.Text, Len(Txt_Address.Text) - 1)
    CreateFolders (Txt_Address.Text & "\" & ss)
    Cmd_Go_Click
End Sub

Sub Plaster_Click()   '粘贴
    '    On Error Resume Next
    If Right(Txt_Address.Text, 1) = "\" Then Txt_Address.Text = Left(Txt_Address, Len(Txt_Address) - 1)
    If coname <> "" Then
        FileCopy copath, Txt_Address & "\" & coname
    Else
        Set P_ofso = CreateObject("scripting.filesystemobject")
        '        On Error Resume Next
        P_ofso.CopyFolder Trim(copath), Txt_Address, True
    End If
    Cmd_Go_Click
End Sub

Sub ReName_Click()  '重命名
    On Error GoTo MyErr
    Dim s As String
    Dim t As String
    Dim d As String
    Dim n As String
    Dim m As String
    Dim I As Integer
    Dim filedata As WIN32_FIND_DATA
    If Right(Me.Txt_Address.Text, 1) = "\" Then
        Me.Txt_Address.Text = Left(Me.Txt_Address.Text, Len(Me.Txt_Address.Text) - 1)
    End If
    If coname = "" Then           '如果是文件夹
        For I = Len(Txt_Address.Text) To 1 Step -1
            If Mid$(Txt_Address.Text, I, 1) = "\" Then
                n = Right(Txt_Address.Text, Len(Txt_Address.Text) - I)
                m = Left(Txt_Address.Text, I)
                Exit For
            End If
        Next
        t = InputBox(Txt_Address.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "输入更改后的文件夹名", "重命名", n)
        If t <> "" Then Name Txt_Address.Text As m & t
    Else
        s = Txt_Address.Text & "\" & coname
        For I = Len(coname) To 1 Step -1
            If Mid$(coname, I, 1) = "." Then
                n = Mid$(coname, 1, I - 1)
                m = Mid$(coname, I + 1, Len(coname) - I)
                Exit For
            End If
        Next I
        t = InputBox(s & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "输入更改后的文件名", "重命名", n)
        filedata = Findfile(s)
        If (filedata.dwFileAttributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY Then
            If MsgBox("确实要将只读文件" & coname & "重命名为" & t & "." & m, 36, "提示信息") = vbNo Then
                Exit Sub
            End If
        End If
        d = IIf(Right(Txt_Address.Text, 1) <> "\", Txt_Address.Text & "\", Txt_Address.Text) & t & "." & m
        If t <> "" Then Name s As d
    End If
    Cmd_Go_Click
    Exit Sub
MyErr:
    Call SHOWWRONG("错误 " & ERR.Number & ".  " & ERR.Description, 0)
End Sub
Private Sub Tree_Add(Path As String)
    Dim MyFSO As New FileSystemObject
    Dim MyFolder As Folder
    Dim MyFile As File

    Dim Folder1 As Folder       '指定文件夹下的子文件夹
    Dim Folder As Folder        '指定文件夹下的子文件夹

    Set MyFolder = MyFSO.GetFolder(Path)
    TreeView1.Nodes.Clear
    TreeView1.Nodes.Add , , Path, Left(Path, 2), 2

    For Each Folder In MyFolder.SubFolders
        TreeView1.Nodes.Add Path, tvwChild, Left(Path, 2) & "\" & Folder.name, Folder.name, 1
    Next
    Tree_DataExpanded (Path)    '添加
End Sub

Private Sub Tree_DataExpanded(Path As String)
    Dim MyFSO As New FileSystemObject
    Dim MyFolder As Folder
    Dim MyFile As File

    Dim Folder1 As Folder       '指定文件夹下的子文件夹
    Dim Folder As Folder        '指定文件夹下的子文件夹

    Set MyFolder = MyFSO.GetFolder(Path)
    Dim hImgSmall As Long       ' 存储图片句柄
    Dim filename As String      ' 要获取图片的文件路径
    Dim r As Long

    Dim itmX As ListItem        '定义listItem类型变量
    Dim num                     '定义变量作为标识
    Dim sExtension As String    '定义变量存储文件的扩展名

    ListView1.ListItems.Clear
    ListView1.SmallIcons = Nothing
    Imt_LV.ListImages.Clear
    num = num + 1
    Imt_LV.ListImages.Add , "Folder", Imt_Tree.ListImages(1).PICTURE

    Set MyFolder = MyFSO.GetFolder(Path & "\")

    For Each Folder1 In MyFolder.SubFolders
        ListView1.SmallIcons = Me.Imt_LV
        Set itmX = ListView1.ListItems.Add(, , Folder1.name, , 1)
        itmX.SubItems(2) = "文件夹"
    Next

    For Each MyFile In MyFolder.files
        num = num + 1
        filename$ = Path & MyFile.name

        hImgSmall = SHGetFileInfo(filename$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        Me.Pic_FileICO.PICTURE = Nothing
        r& = ImageList_Draw(hImgSmall&, shinfo.iIcon, Pic_FileICO.hdc, 0, 0, ILD_TRANSPARENT)
        Me.Pic_FileICO.PICTURE = Me.Pic_FileICO.image

        Imt_LV.ListImages.Add , "ico" & num, Me.Pic_FileICO.PICTURE

        ListView1.SmallIcons = Me.Imt_LV
        Set itmX = ListView1.ListItems.Add(, , MyFile.name, , num)
        itmX.SubItems(1) = GetFileSize(Path & MyFile.name)      '添加文件大小
        sExtension = Left$(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr$(0)) - 1)
        itmX.SubItems(2) = sExtension                           '添加文件类型
        itmX.SubItems(3) = GetModifyTime(Path & MyFile.name)    '添加文件修改日期
    Next
    
    LA(0).Caption = Format(Me.ListView1.ListItems.Count, "000")
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim SPATH As String, FILETIT As String, AUTHOR As String, SONGTIT As String, SINGER_LOGO As String
        If Right(Me.Txt_Address.Text, 1) = "\" Then
            Me.Txt_Address.Text = Left(Me.Txt_Address.Text, Len(Me.Txt_Address.Text) - 1)
        End If
        sFileName = Me.Txt_Address.Text & "\" & ListView1.SelectedItem.Text
        SPATH = Me.Txt_Address.Text
        coname = ListView1.SelectedItem.Text
        bFileFlag = True
        Select Case UCase(Right(coname, 3))
        Case "MP3", "WMA", "MID"
        IMS.Visible = True
        PVIEW.Visible = False
IMS.STOP_IT
ID3V1.filename = sFileName
ID3V1.ReadTag
AUTHOR = ID3V1.tagArtist
SONGTIT = ID3V1.tagTitle
IMS.SETTXT SONGTIT, AUTHOR
IMS.MUSIC_URL = sFileName
SINGER_LOGO = App.Path & "\MEDIA\MUSICPICTURE\" & AUTHOR & ".BMP"
If PathFileExists(SINGER_LOGO) = 1 Then IMS.SETPIC SINGER_LOGO Else IMS.Cls

        Case "JPG", "BMP", "GIF"
        PVIEW.Visible = True
        IMS.Visible = False
        PICPR.PICTURE = LoadPicture(sFileName)
        Call SETPRE
        Case Else
        IMS.Visible = False
        PVIEW.Visible = False
        End Select
        Call updatestats(sFileName)
        Call RE_UI
End Sub
Sub SETPRE()
If PICPR.Height > PVIEW.ScaleHeight Or PICPR.Width > PVIEW.ScaleWidth Then
PV.Height = PVIEW.ScaleHeight
PV.Width = PVIEW.ScaleWidth * (PV.Height / PICPR.ScaleHeight)
Dimention2 PV, PICPR, PICPR.ScaleWidth * (PV.Height / PICPR.ScaleHeight), PV.Height
PV.Move (PVIEW.ScaleWidth - PV.Width) / 2, 0
Else
Dimention2 PV, PICPR, PICPR.Width, PICPR.Height
PV.Move (PVIEW.ScaleWidth - PV.Width) / 2, (PVIEW.ScaleHeight - PV.Height) / 2
End If
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If TreeView2.Visible = True Then TreeView2.Visible = False
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu Frmm.文件管理
End Sub

Private Sub PINFO_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub PV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub PVIEW_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CMV(Me)
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If TreeView2.Visible = True Then TreeView2.Visible = False

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    DoEvents
    Dim MyFSO As New FileSystemObject
    Dim MyFolder As Folder
    Dim MyFile As File

    Dim Folder1 As Folder       '指定文件夹下的子文件夹

    Dim hImgSmall As Long       ' The handle to the system image list
    Dim filename As String      ' The file name to get icon from
    Dim r As Long

    Dim itmX As ListItem
    Dim num                     '定义变量作为标识
    Dim sExtension As String    '定义变量存储文件的扩展名

    ListView1.ListItems.Clear
    ListView1.SmallIcons = Nothing
    Imt_LV.ListImages.Clear
    num = num + 1
    Imt_LV.ListImages.Add , "Folder", Imt_Tree.ListImages(1).PICTURE

    '将文件夹路径指定到当前所选择的路径下
    Set MyFolder = MyFSO.GetFolder(Node.fullPath & "\")
    If Node.children = 0 Then
        For Each Folder1 In MyFolder.SubFolders
            TreeView1.Nodes.Add Node.fullPath, tvwChild, Node.fullPath & "\" & Folder1.name, Folder1.name, 1
        Next
    End If
    For Each Folder1 In MyFolder.SubFolders
        ListView1.SmallIcons = Me.Imt_LV
        Set itmX = ListView1.ListItems.Add(, , Folder1.name, , 1)
        itmX.SubItems(2) = "文件夹"
    Next

    For Each MyFile In MyFolder.files
        num = num + 1
        filename$ = Node.fullPath & "\" & MyFile.name
        hImgSmall = SHGetFileInfo(filename$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        Me.Pic_FileICO.PICTURE = Nothing
        r& = ImageList_Draw(hImgSmall&, shinfo.iIcon, Pic_FileICO.hdc, 0, 0, ILD_TRANSPARENT)
        Me.Pic_FileICO.PICTURE = Me.Pic_FileICO.image

        Imt_LV.ListImages.Add , "ico" & num, Me.Pic_FileICO.PICTURE

        ListView1.SmallIcons = Me.Imt_LV
        Set itmX = ListView1.ListItems.Add(, , MyFile.name, , num)
        itmX.SubItems(1) = GetFileSize(Node.fullPath & "\" & MyFile.name)  '添加文件大小
        sExtension = Left$(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr$(0)) - 1)
        itmX.SubItems(2) = sExtension              '添加文件类型
        itmX.SubItems(3) = GetModifyTime(Node.fullPath & "\" & MyFile.name)   '添加文件修改日期
    Next
    Txt_Address.Text = Node.fullPath
    LA(0).Caption = Format(Me.ListView1.ListItems.Count, "000")
End Sub
Private Function GetModifyTime(sFile As String) As String
    Dim dtWrite As Date    '创建时间
    Dim lpReOpenBuff As OFSTRUCT
    Dim FileHandle As Long
    Dim FileInfo As BY_HANDLE_FILE_INFORMATION
    Dim tZone As TIME_ZONE_INFORMATION
    Dim fTime As SYSTEMTIME
    Dim Bias As Long

    FileHandle = OpenFile(sFile, lpReOpenBuff, OF_READ)
    Call GetFileInformationByHandle(FileHandle, FileInfo)    '利用 File Handle 读取文件资讯
    Call CloseHandle(FileHandle)
    Call GetTimeZoneInformation(tZone)                       '读取Time Zone，因为上一步骤的文件时间是格林威治时间
    Bias = tZone.Bias                                        '时间差,以"分"为单位
    Call FileTimeToSystemTime(FileInfo.ftLastWriteTime, fTime)
    dtWrite = DateSerial(fTime.wYear, fTime.wMonth, fTime.wDay) + TimeSerial(fTime.wHour, fTime.wMinute - Bias, fTime.wSecond)
    GetModifyTime = dtWrite
End Function
Private Function GetFileSize(sFile As String) As String
On Error Resume Next
    If Round(FileLen(sFile) / 1024) = 0 Then
        GetFileSize = 1 & " KB"
    Else
        GetFileSize = Round(FileLen(sFile) / 1024) & " KB"
    End If
End Function

Private Function GetExtension(filename As String) As String   '获得文件的扩展名
    Dim I, j, Path, Ext As Integer
    For I = Len(filename) To 1 Step -1      '从文件名的长度到文件名的第一个字符作循环
        If Mid(filename, I, 1) = "." Then   '如果当前的字符是"."
            Ext = I     '设置变量Ext的值为i
            Exit For
        End If
    Next I
    If Ext = 0 Then
        Exit Function
    End If
    GetExtension = UCase(Mid(filename, Ext + 1, Len(filename) - Ext))
End Function

Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo MyErr
    If TreeView2.SelectedItem.Text = "我的电脑" Then
    Else
        If Len(TreeView2.SelectedItem.Key) = 1 Then
            Drive1.Drive = TreeView2.SelectedItem.Key
            Txt_Address.Text = TreeView2.SelectedItem.Key & ":\"
        Else
            Txt_Address.Text = TreeView2.SelectedItem.Key
        End If
        Cmd_Go_Click
    End If
    TreeView2.Visible = False
    Exit Sub
MyErr:
    If ERR.Number = 68 Then
        Call SHOWWRONG("实时错误:68 设备不可用.", 0)
    End If
End Sub

Private Sub Txt_Address_Change()
Cmd_Download.SETTXT UCase(Left(Me.Txt_Address.Text, 1))
End Sub

Private Sub Txt_Address_DblClick()
  TreeView2.Visible = False
End Sub

Private Sub Txt_Address_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Cmd_Go_Click
End Sub

Private Sub x1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then Unload Me
End Sub

Sub RE_UI()
If PVIEW.Visible = True Or IMS.Visible = True Then
PINFO.Move 744, 336
Else
PINFO.Move 744, 128
End If
End Sub
Public Sub updatestats(tfilename As String)
On Error Resume Next
    Dim fTime As SYSTEMTIME
    Dim filedata As WIN32_FIND_DATA
    Dim hImgSmall As Long ' The handle to the system image list
    Dim filename As String ' The file name to get icon from
    Dim hImgLarge&
    Dim r As Long
    filedata = Findfile(tfilename)
    If filedata.nFileSizeHigh = 0 Then
        lbldate(3).Caption = Int(filedata.nFileSizeLow / 1024 / 1024) & " MB"
    Else
        lbldate(3).Caption = Int(filedata.nFileSizeHigh / 1024 / 1024) & " MB"
    End If
    Call FileTimeToSystemTime(filedata.ftCreationTime, fTime)
    lbldate(0) = fTime.wDay & "/" & fTime.wMonth & "/" & fTime.wYear & " " & fTime.wHour & ":" & fTime.wMinute & ":" & fTime.wSecond
    Call FileTimeToSystemTime(filedata.ftLastWriteTime, fTime)
    lbldate(1) = fTime.wDay & "/" & fTime.wMonth & "/" & fTime.wYear & " " & fTime.wHour & ":" & fTime.wMinute & ":" & fTime.wSecond
    Call FileTimeToSystemTime(filedata.ftLastAccessTime, fTime)
    lbldate(2) = fTime.wDay & "/" & fTime.wMonth & "/" & fTime.wYear
filename$ = tfilename
picSmall.Cls
picLarge.Cls
hImgSmall& = SHGetFileInfo(filename$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
hImgLarge& = SHGetFileInfo(filename$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
r& = ImageList_Draw(hImgSmall&, shinfo.iIcon, picSmall.hdc, 0, 0, ILD_TRANSPARENT)
r& = ImageList_Draw(hImgLarge&, shinfo.iIcon, picLarge.hdc, 0, 0, ILD_TRANSPARENT)
End Sub
