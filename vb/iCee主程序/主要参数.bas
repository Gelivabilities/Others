Attribute VB_Name = "MMAIN"
Option Explicit
Public fn As Long, AUTOSERCH As Integer, IS_CHK_LIST As Boolean, IS_M_S As Boolean, IS_MINI_MINI As Boolean
Public rtn As Long, IS_FULLSCREEN As Boolean, IS_CAPTURE As Boolean, IS_AM As Boolean
Public OL As String, IETIP As Integer, WILL_DEL As String, WILL_DEL_IDX As Integer, LONELY_MODE As Boolean
Public MOVEINET As Boolean, FIRSTRUN As Boolean, MOVE_TRANS As Integer, HAS_HEAD As Boolean
Public GAMESOUND As Boolean, H_CHANGE As Boolean, IS_NET As Boolean, IS_MINI As Boolean
Public TIP As New CLSPOP, R_P_THU As Integer, CAN_MINI As Boolean, PLAYDSB As Long, IS_LOCK As Boolean
Public ID3V1 As New clsID3v1, CAN_SHOW_MEUN As Boolean
Public IS_FIRST_LOAD_ACT As Boolean, IS_MINI_LIST As Boolean, USE_PIC_FORM As Boolean
Public GLOADFORM As Boolean, ALWAYSONTOP As Boolean, IS_SET As Boolean, HASUSB As Boolean
Public D_L_SHOW As Boolean, IS_CPU_M As Boolean, FAV_IT As Boolean, AUTO_TIP As Boolean
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public strNewComputerName As String, W_P_URL As String, W_P_CODE As String
Public lngReturn As Long, Dpath As String, HAS_NET As Boolean, objWMIService, colProcessList, objProcess
Global gHW As Long
Public sKeys As Collection, IS_CHECK_CLIP As Boolean, AUTO_SINGER As Boolean, RUN_MODE As Integer
Public H_DOS As Integer
Public SIBRIT As Integer
Public Const GWL_WNDPROCB = -4
Public IEver As String, SONG_SZIE As String, SONG_TIME As String, ITS_KPS As String
Public KBS As Long '为浏览器图片做坐标
Public Col As New clsMD5, RESL As Long, UI_BKCOLOR As Long
Public oMagneticWnd As New cMagneticWnd '磁性窗口
Public RPC As New ROUND_FORM  '圆角图像框picbox
Public Wrn As New FrmWrong
Public fso As New Scripting.FileSystemObject 'Form's Global Declarations
Public READYLOAD As Boolean
Public Init As Integer
Public m_OldProc As Long
Public LOGO As String   '头像变量
Public LastRecvBytes As Long, LastSentBytes As Long
Public MOUSEMO As Boolean
Public MAINSTYLE As Integer
Public gMS As Double
Public NOTECOUND As Long
Public Result As Double
Public sDriveNames As String
Public lBuffer As Long
Public lReturn As Long
Public nLoopCtr As Integer
Public nOffset As Integer
Public sTempStr As String
Public Root As String
Public WinPath As String
Public WinSysPath As String
Public RECLICK As Boolean
Public RCODE As Boolean
Public TheForm As Form
Public OldWindowProc As Long
Public Volume_Name As String
Public Serial_Number As Long
Public Max_Component_Length As Long
Public File_System_Flags As Long
Public File_System_Name As String
Public pos As Integer
Public Dbl_Total As Double
Public Dbl_Free As Double
Public COLOR_NOR As Long, COLOR_HIGH As Long
Public lSectorsPerCluster As Long
Public lBytesPerSector As Long
Public lFreeClusters As Long
Public lTotalClusters As Long
Public sDrive As String
Private rectLastTray As RECT
Private rectLastRebar As RECT
Private rectLastNotify As RECT
Private hwndForm As Long
Private lngTimer As Long
Private IntWidth As Integer
Private IntHeight As Integer
Private blnGrow As Boolean
Public StCT As CONTEXT
Dim gdip_Graphics, gdip_Token, gdip_pngImage
Dim File_Share_Flag As Long '定义锁定文件夹的变量
Dim hDir As Long '定义锁定文件夹的变量
Public NEWS As Integer
Public objTimer As clsWaitableTimer
Public ITIME As Long, Song As Integer  '储存当前 播放的歌曲 位置
Public SONGNAME As String, 专辑 As String, 年代 As String '储存当前 播放的歌曲名称
Public Songpath As String '储存当前 播放的歌曲 路径
Public LOLIPOP As Integer
Public FIRSTTIME As Boolean
Public filename() As String
Public FileCount As Integer
Public Preview_Handle As Long
Public LASTDOWNFILE As String
Public Const EM_GETLINECOUNT = &HBA
Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 32
Public Status As RASCONNSTATUS95
Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
Public Const BUFFER_LEN = 256
Public Const LB_FINDSTRINGEXACT = &H1A2  '在ListBox中精确查找
Public Const LB_FINDSTRING = &H18F   '在ListBox中模糊查找
Public Const CB_FINDSTRINGEXACT = &H158  '在ComboBox中精确查找
Private Const MAXERRORLENGTH = 128   '  max error text length (including NULL)
Private Const MIDIMAPPER = (-1)
Private Const MIDI_MAPPER = (-1)
Public Const CBF_FAIL_ADVISES = &H4000
Public Const CBF_FAIL_ALLSVRXACTIONS = &H3F000
Public Const CBF_FAIL_CONNECTIONS = &H2000
Public Const CBF_FAIL_EXECUTES = &H8000
Public Const CBF_FAIL_POKES = &H10000
Public Const CBF_FAIL_REQUESTS = &H20000
Public Const CBF_FAIL_SELFCONNECTIONS = &H1000
Public Const CP_WINANSI = 1004
Public Const XCLASS_FLAGS = &H4000
Public Const XTYP_EXECUTE = (&H50 Or XCLASS_FLAGS)
Private Const LB_ERR = -1
Global strCFile As String
Private Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Private Const STATUS_ACCESS_DENIED = &HC0000022
Private Const STATUS_INVALID_HANDLE = &HC0000008
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const NO_INHERITANCE = 0
Private Const DACL_SECURITY_INFORMATION = &H4
Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEFIRST = &H200
Public Const WM_KEYLAST = &H108
Public Const WM_KEYFIRST = &H100
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1

Public Const CF_TEXT = 1
Public Const SYNCHRONIZE As Long = &H100000
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2
Public Const SHERB_NOCONFIRMATION = &H1
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_ALWAYS = 4
Public oldproc As Long '这个是屏蔽文本菜单的变量
Public Info As DEV_BROADCAST_HDR
Public vInfo As DEV_BROADCAST_VOLUME
Public PrevProc As Long
Public RecvProc As Long
Public GFE(60000) As String
Public GF(60000) As String
Public HotKey As Long
Public HotKey_Cild As Long
Public sDefInitFileName As String
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const LANG_USER_DEFAULT = &H400&
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOZORDER = &H4
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412
Public Const lType = 4
Public Const lSize = 4
Public Const SIZE_KB As Double = 1024
Public Const SIZE_MB As Double = 1024 * SIZE_KB
Public Const SIZE_GB As Double = 1024 * SIZE_MB
Public Const SIZE_TB As Double = 1024 * SIZE_GB
Public Const LB_ADDSTRING = &H180

Global tmpCol As Long
Global r As Long, G As Long, b As Long
Global larrCol() As Long

Public Type cPoint
    cx As Double
    cy As Double
End Type

Public Type Colors
    lBCol As Long
    lFCol As Long
End Type
Const TH32CS_SNAPheaplist = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPthread = &H4
Const TH32CS_SNAPmodule = &H8
Const TH32CS_SNAPall = (TH32CS_SNAPheaplist Or TH32CS_SNAPPROCESS Or TH32CS_SNAPthread Or TH32CS_SNAPmodule)
Const TH32CS_INHERIT = &H80000000
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long
Public Const HKEY_DYN_DATA As Long = &H80000006
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const OOMPS = SWP_NOSIZE Or SWP_NOMOVE
Public Const offset = 500
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Public Const WS_EX_TRANSPARENT = &H20&
Public Const HIGH_PRIORITY_CLASS = &H80 '新进程有非常高的优先级，它优先于大多数应用程序.基本值是13.注意尽量避免采用这个优先级
Public Const IDLE_PRIORITY_CLASS = &H40 '新进程应该有非常低的优先级——只有在系统空闲的时候才能运行.基本值是4
Public Const NORMAL_PRIORITY_CLASS = &H20   '标准优先级.如进程位于前台，则基本值是9；如在后台，则优先值是7
Public Const REALTIME_PRIORITY_CLASS = &H100 '立即对事件作出响应，执行关键时间的任务.会抢先于操作系统组件之前运行.
Public Const LOWER_PRIORITY_CLASS = &H4000  '未公开. 较低优先级别
Const WM_DEVICECHANGE As Long = &H219
Const DBT_DEVICEARRIVAL As Long = &H8000&
Const DBT_DEVICEREMOVECOMPLETE As Long = &H8004&
Const DBT_DEVTYP_VOLUME As Long = &H2
Private BmpBits() As Long
'目录浏览树变量
Public OsName$, TmpStr$, Ary
Public Const CSIDL_DESKTOP = &H0    '桌面
Public Const CSIDL_PROGRAMS = &H2   '程度组
Public Const CSIDL_CONTROLS = &H3   '控制面板
Public Const CSIDL_PRINTERS = &H4   '打印机目录
Public Const CSIDL_PERSONAL = &H5   '文件目录
Public Const CSIDL_FAVORITES = &H6  '收藏夹
Public Const CSIDL_STARTUP = &H7    '启动目录
Public Const CSIDL_RECENT = &H8     '临时目录
Public Const CSIDL_SENDTO = &H9     '发生到目录
Public Const CSIDL_BITBUCKET = &HA  '删除目录
Public Const CSIDL_STARTMENU = &HB  '开始菜单目录
Public Const CSIDL_DESKTOPDIRECTORY = &H10  'Windows\Desktop目录
Public Const CSIDL_DRIVES = &H11    '我的电脑
Public Const CSIDL_NETWORK = &H12   '网上邻居
Public Const CSIDL_NETHOOD = &H13   '网上邻居目录
Public Const CSIDL_FONTS = &H14     '字体目录
Public Const CSIDL_TEMPLATES = &H15 '新建目录
Private Type IO_STATUS_BLOCK
Status As Long
Information As Long
End Type

Private Type UNICODE_STRING
Length As Integer
MaximumLength As Integer
Buffer As Long
End Type

Public Const SHGFI_ICON = &H100

Private Const OBJ_INHERIT = &H2
Private Const OBJ_PERMANENT = &H10
Private Const OBJ_EXCLUSIVE = &H20
Private Const OBJ_CASE_INSENSITIVE = &H40
Private Const OBJ_OPENIF = &H80
Private Const OBJ_OPENLINK = &H100
Private Const OBJ_KERNEL_HANDLE = &H200
Private Const OBJ_VALID_ATTRIBUTES = &H3F2

Private Type OBJECT_ATTRIBUTES
Length As Long
RootDirectory As Long
ObjectName As Long
Attributes As Long
SecurityDeor As Long
SecurityQualityOfService As Long
End Type

Public Type TypeIcon
    CBSIZE As Long
    PicType As PictureTypeConstants
    hIcon As Long
End Type

Private Type ACL
AclRevision As Byte
Sbz1 As Byte
AclSize As Integer
AceCount As Integer
Sbz2 As Integer
End Type

Public Enum ISPN_FT_SIZETYPE
    ISPN_FT_SIZEINBYTES
    ISPN_FT_SIZEINMB
    ISPN_FT_SIZEINKB
    ISPN_FT_SIZEAUTO
End Enum

Public Type PICINFO
    PicWidth As Long
    PicHeight As Long
End Type
Private Enum ACCESS_MODE
NOT_USED_ACCESS
GRANT_ACCESS
SET_ACCESS
DENY_ACCESS
REVOKE_ACCESS
SET_AUDIT_SUCCESS
SET_AUDIT_FAILURE
End Enum
Public Enum GpStatus
   OK = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum
Public Type BLENDFUNCTION
BlendOp As Byte
BlendFlags As Byte
SourceConstantAlpha As Byte
AlphaFormat As Byte
End Type
Private Enum MULTIPLE_TRUSTEE_OPERATION
NO_MULTIPLE_TRUSTEE
TRUSTEE_IS_IMPERSONATE
End Enum

Private Enum TRUSTEE_FORM
TRUSTEE_IS_SID
TRUSTEE_IS_NAME
End Enum

Private Enum TRUSTEE_TYPE
TRUSTEE_IS_UNKNOWN
TRUSTEE_IS_USER
TRUSTEE_IS_GROUP
End Enum

Private Type TRUSTEE
pMultipleTrustee As Long
MultipleTrusteeOperation As MULTIPLE_TRUSTEE_OPERATION
TrusteeForm As TRUSTEE_FORM
TrusteeType As TRUSTEE_TYPE
ptstrName   As String
End Type

Private Type EXPLICIT_ACCESS
grfAccessPermissions As Long
grfAccessMode   As ACCESS_MODE
grfInheritance  As Long
TRUSTEE As TRUSTEE
End Type

Private Enum SE_OBJECT_TYPE
SE_UNKNOWN_OBJECT_TYPE = 0
SE_FILE_OBJECT
SE_SERVICE
SE_PRINTER
SE_REGISTRY_KEY
SE_LMSHARE
SE_KERNEL_OBJECT
SE_WINDOW_OBJECT
SE_DS_OBJECT
SE_DS_OBJECT_ALL
SE_PROVIDER_DEFINED_OBJECT
SE_WMIGUID_OBJECT
End Enum


Private Type MIB_IFROW
wszName(0 To 511) As Byte
dwIndex As Long '// index of the interface
dwType As Long  '// type of interface
dwMtu As Long   '// max transmission unit
dwSpeed As Long '// speed of the interface
dwPhysAddrLen As Long   '// length of physical address
bPhysAddr(0 To 7) As Byte   '// physical address of adapter
dwAdminStatus As Long   '// administrative status
dwOperStatus As Long '// operational status
dwLastChange As Long '// last time operational status changed
dwInOctets As Long  '// octets received
dwInUcastPkts As Long   '// unicast packets received
dwInNUcastPkts As Long  '// non-unicast packets received
dwInDiscards As Long '// received packets discarded
dwInErrors As Long  '// erroneous packets received
dwInUnknownProtos As Long   '// unknown protocol packets received
dwOutOctets As Long '// octets sent
dwOutUcastPkts As Long  '// unicast packets sent
dwOutNUcastPkts As Long '// non-unicast packets sent
dwOutDiscards As Long   '// outgoing packets discarded
dwOutErrors As Long '// erroneous packets sent
dwOutQLen As Long   '// output queue length
dwDescrLen As Long  '// length of bDescr member
bDescr(0 To 255) As Byte '// interface description
End Type
Private Declare Function GetIfTable Lib "iphlpapi" (ByRef pIfRowTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Const ERROR_NOT_SUPPORTED = 50&
Private Enum InterfaceTypes
MIB_IF_TYPE_OTHER = 1
MIB_IF_TYPE_ETHERNET = 6
MIB_IF_TYPE_TOKENRING = 9
MIB_IF_TYPE_FDDI = 15
MIB_IF_TYPE_PPP = 23
MIB_IF_TYPE_LOOPBACK = 24
MIB_IF_TYPE_SLIP = 28
End Enum
Public Type Flow_INFO
lngBytesReceived As Long
lngBytesSent As Long
End Type
Public SoftSAFE As Long
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Public Declare Function SetThreadContext Lib "kernel32.dll" (ByVal hThread As Long, ByVal lpContext As Long) As Long
Public Declare Function GetLastError Lib "kernel32.dll" () As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const EXCEPTION_MAXIMUM_PARAMETERS  As Long = 15
Public Const EXCEPTION_EXECUTE_HANDLER As Long = 1
Public Const EXCEPTION_CONTINUE_EXECUTION  As Long = -1
Public Const EXCEPTION_CONTINUE_SEARCH As Long = 0
Public Const MAXIMUM_SUPPORTED_EXTENSION As Long = 512
Public Const SIZE_OF_80387_REGISTERS As Long = 80
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_NEW = 1
Public Const MAX_IDE_DRIVES = 4
Public Const READ_ATTRIBUTE_BUFFER_SIZE = 512
Public Const IDENTIFY_BUFFER_SIZE = 512
Public Const READ_THRESHOLD_BUFFER_SIZE = 512
Public Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16
Public Const DFP_GET_VERSION = &H74080
Public Const DFP_SEND_DRIVE_COMMAND = &H7C084
Public Const DFP_RECEIVE_DRIVE_DATA = &H7C088

Public Type GETVERSIONOUTPARAMS
   bVersion   As Byte
   bRevision  As Byte
   bReserved  As Byte
   bIDEDeviceMap  As Byte
   fCapabilities  As Long
   dwReserved(3)  As Long
End Type
Public Declare Sub CopyMemoryIp Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub InitCommonControls Lib "Comctl32" ()
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As Any, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Const SMART_CYL_LOW = &H4F
Public Const SMART_CYL_HI = &HC2
Public Type DRIVERSTATUS
   bDriverError  As Byte
   bIDEStatus As Byte
   bReserved(1)  As Byte
   dwReserved(1) As Long
 End Type
Public Enum DRIVER_ERRORS
   SMART_NO_ERROR = 0
   SMART_IDE_ERROR = 1
   SMART_INVALID_FLAG = 2
   SMART_INVALID_COMMAND = 3
   SMART_INVALID_BUFFER = 4
   SMART_INVALID_DRIVE = 5
   SMART_INVALID_IOCTL = 6
   SMART_ERROR_NO_MEM = 7
   SMART_INVALID_REGISTER = 8
   SMART_NOT_SUPPORTED = 9
   SMART_NO_IDE_DEVICE = 10
End Enum
Public Type IDSECTOR
   wGenConfig As Integer
   wNumCyls   As Integer
   wReserved  As Integer
   wNumHeads  As Integer
   wBytesPerTrack As Integer
   wBytesPerSector As Integer
   wSectorsPerTrack   As Integer
   wVendorUnique(2)   As Integer
   sSerialNumber(19)  As Byte
   wBufferType As Integer
   wBufferSize As Integer
   wECCSize   As Integer
   sFirmwareRev(7) As Byte
   sModelNumber(39)   As Byte
   wMoreVendorUnique  As Integer
   wDoubleWordIO  As Integer
   wCapabilities  As Integer
   wReserved1 As Integer
   wPIOTiming As Integer
   wDMATiming As Integer
   wBS As Integer
   wNumCurrentCyls As Integer
   wNumCurrentHeads   As Integer
   wNumCurrentSectorsPerTrack As Integer
   ulCurrentSectorCapacity As Long
   wMultSectorStuff   As Integer
   ulTotalAddressableSectors  As Long
   wSingleWordDMA As Integer
   wMultiWordDMA  As Integer
   bReserved(127) As Byte
End Type

Public Type SENDCMDOUTPARAMS
  cBufferSize   As Long
  DRIVERSTATUS  As DRIVERSTATUS
  bBuffer() As Byte
End Type

Public Const SMART_READ_ATTRIBUTE_VALUES = &HD0
Public Const SMART_READ_ATTRIBUTE_THRESHOLDS = &HD1
Public Const SMART_ENABLE_DISABLE_ATTRIBUTE_AUTOSAVE = &HD2
Public Const SMART_SAVE_ATTRIBUTE_VALUES = &HD3
Public Const SMART_EXECUTE_OFFLINE_IMMEDIATE = &HD4
Public Const SMART_ENABLE_SMART_OPERATIONS = &HD8
Public Const SMART_DISABLE_SMART_OPERATIONS = &HD9
Public Const SMART_RETURN_SMART_STATUS = &HDA

Public Const NUM_ATTRIBUTE_STRUCTS = 30

Public Type DRIVEATTRIBUTE
   bAttrID As Byte
   wStatusFlags As Integer
   bAttrValue As Byte
   bWorstValue As Byte
   bRawValue(5) As Byte
   bReserved As Byte
End Type
Public Enum APIRegistryRoots
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
End Enum
Public Enum STATUS_FLAGS
   PRE_FAILURE_WARRANTY = &H1
   ON_LINE_COLLECTION = &H2
   PERFORMANCE_ATTRIBUTE = &H4
   ERROR_RATE_ATTRIBUTE = &H8
   EVENT_COUNT_ATTRIBUTE = &H10
   SELF_PRESERVING_ATTRIBUTE = &H20
End Enum
Public Type ATTRTHRESHOLD
   bAttrID As Byte
   bWarrantyThreshold As Byte
   bReserved(9) As Byte
End Type
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type ATTR_DATA
AttrID As Byte
AttrName As String
AttrValue As Byte
ThresholdValue As Byte
WorstValue As Byte
StatusFlags As STATUS_FLAGS
End Type

Public Type DRIVE_INFO
bDriveType As Byte
SerialNumber As String
Model As String
FirmWare As String
Cilinders As Long
Heads As Long
SecPerTrack As Long
BytesPerSector As Long
BytesperTrack As Long
NumAttributes As Byte
Attributes() As ATTR_DATA
End Type
'JPEG（这个好麻烦）
Private Type LSJPEGHeader
  jSOI As Integer           '图像开始标识 0,1 两个字节为 FF D8 低位在前，即 -9985
  jAPP0 As Integer          'APP0块标识 2,3 两个字节为 FF E0
  jAPP0Length(1) As Byte    'APP0块标识后的长度，两个字节，高位在前
'  jJFIFName As Long         'JFIF标识 49(J) 48(F) 44(I) 52(F)
'  jJFIFVer1 As Byte         'JFIF版本
'  jJFIFVer2 As Byte         'JFIF版本
'  jJFIFVer3 As Byte         'JFIF版本
'  jJFIFUnit As Byte
'  jJFIFX As Integer
'  jJFIFY As Integer
'  jJFIFsX As Byte
'  jJFIFsY As Byte
End Type

Private Type LSJPEGChunk
  jcType As Integer         '标识（按顺序）:APPn(0,1~15)为 FF E1~FF EF; DQT为 FF DB(-9217)
                            'SOFn(0~3)为 FF C0(-16129),FF C1(-15873),FF C2(-15617),FF C3(-15361)
                            'DHT为 FF C4(-15105); 图像数据开始为 FF DA
  jcLength(1) As Byte       '标识后的长度，两个字节，高位在前
                            '若标识为SOFn，则读取以下信息；否则按照长度跳过，读下一块
  jBlock As Byte            '数据采样块大小 08 or 0C or 10
  jHeight(1) As Byte        '高度两个字节，高位在前
  jWidth(1) As Byte         '宽度两个字节，高位在前
'  jColorType As Byte        '颜色类型 03，后跟9字节，然后是DHT
End Type

'PNG文件头
Private Type LSPNGHeader
  pType As Long             '标识 0,1,2,3 四个字节为 89 50(P) 4E(N) 47(G) 低位在前，即 1196314761
  pType2 As Long            '标识 4,5,6,7 四个字节为 0D 0A 1A 0A
  pIHDRLength As Long       'IHDR块标识后的长度，疑似固定 00 0D，高位在前，即 13
  pIHDRName As Long         'IHDR块标识 49(I) 48(H) 44(D) 52(R)
  pWidth(3) As Byte            '宽度 16,17,18,19 四个字节，高位在前
  pHeight(3) As Byte           '高度 20,21,22,23 四个字节，高位在前
'  pBitDepth As Byte
'  pColorType As Byte
'  pCompress As Byte
'  pFilter As Byte
'  pInterlace As Byte
End Type

'GIF文件头（这个好简单）
Private Type LSGIFHeader
  gType1 As Long            '标识 0,1,2,3 四个字节为 47(G) 49(I) 46(F) 38(8) 低位在前，即 944130375
  gType2 As Integer         '版本 4,5 两个字节为 7a单幅静止图像9a若干幅图像形成连续动画
  gWidth As Integer         '宽度 6,7 两个字节，低位在前
  gHeight As Integer        '高度 8,9 两个字节，低位在前
End Type
Public Enum IDE_DRIVE_NUMBER
PRIMARY_MASTER
PRIMARY_SLAVE
SECONDARY_MASTER
SECONDARY_SLAVE
End Enum

Public Type EXCEPTION_POINTERS
pExceptionRecord As Long 'pointer to an EXCEPTION_RECORD structure
pContextRecord  As Long 'pointer to a CONTEXT structure
End Type

Public Type EXCEPTION_RECORD
ExceptionCode As Long
ExceptionFlags As Long
pExceptionRecord As Long ' Pointer to an EXCEPTION_RECORD structure
ExceptionAddress As Long
NumberParameters As Long
ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

Public Type FLOATING_SAVE_AREA
ControlWord As Long '   DWORD ControlWord;
StatusWord As Long  '   DWORD StatusWord;
TagWord As Long '   DWORD TagWord;
ErrorOffset As Long '   DWORD ErrorOffset;
ErrorSelector As Long   '   DWORD ErrorSelector;
DataOffset As Long  '   DWORD DataOffset;
DataSelector As Long '   DWORD DataSelector;
RegisterArea(SIZE_OF_80387_REGISTERS - 1) As Byte ' BYTE RegisterArea[SIZE_OF_80387_REGISTERS];
Cr0NpxState As Long '   DWORD Cr0NpxState;
End Type

Public cls_Rijndael As cRijndael 'AES加密类
Private Type TRGB
r As Integer
G As Integer
b As Integer
End Type
Public tmpRGB As TRGB
Public Type EVENTMSG
Message As Long
paramL As Long
paramH As Long
TimE As Long
hwnd As Long
End Type
Public Type RASCONNSTATUS95
  dwSize As Long
  RasConnState As Long
  dwError As Long
  szDeviceType(RAS95_MaxDeviceType) As Byte
  szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Public Type RGB
Red As Byte
Green As Byte
Blue As Byte
End Type
Public Type RASCONN95
  dwSize As Long
  hRasCon As Long
  szEntryName(RAS95_MaxEntryName) As Byte
  szDeviceType(RAS95_MaxDeviceType) As Byte
  szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Dim di As DRIVE_INFO
Dim colAttrNames As Collection

Type MIDIOUTCAPS
wMid As Integer
wPid As Integer ' 产品 ID
vDriverVersion As Long ' 设备版本
szPname As String * 32 ' 设备 name
wTechnology As Integer ' 设备类型
wVoices As Integer
wNotes As Integer
wChannelMask As Integer
dwSupport As Long
End Type
Public Enum ShowStyle
 vbHide
 vbMaximizedFocus
 vbMinimizedFocus
 vbMinimizedNoFocus
 vbNormalFocus
 vbNormalNoFocus
End Enum
Type FileInfo   '存储已转换文件信息的自定义类型
Source As String
Target As String
End Type
Public Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
Private Type GUID
Data1 As Long
data2 As Integer
Data3 As Integer
Data4(0 To 7) As Byte
End Type
Public Type GdiplusStartupInput
GdiplusVersion As Long
DebugEventCallback As Long
SuppressBackgroundThread As Long
SuppressExternalCodecs As Long
End Type
Private Type EncoderParameter
GUID As GUID
NumberOfValues As Long
type As Long
Value As Long
End Type
Public Type EncoderParameters
Count As Long
Parameter As EncoderParameter
End Type
Public Type NOTIFYICONDATA
CBSIZE  As Long   'NOTIFYICONDATA类型的大小，用Len(变量名)获得即可
hwnd  As Long   '窗体的名柄
uId  As Long   '图标资源的ID号，通常使用  vbNull
uFlags  As Long   '使哪些参数有效它是以下枚举类型中的  NIF_INFO  Or  NIF_ICON  Or  NIF_TIP  Or  NIF_MESSAGE  四个常数的组合
uCallBackMessage  As Long   '接受消息的事件
hIcon  As Long   '图标名柄
szTip  As String * 128 '当鼠标停留在图标上时显示的Tip文本
dwState  As Long   '通常为  0
dwStateMask  As Long   '通常为  0
uTimeout As Long
szInfo  As String * 256 'Tip文本正文
uTimeoutOrVersion  As Long   '由于VB中没有Union类型，只能用Long型代替
szInfoTitle  As String * 64 'Tip文本的标题
dwInfoFlags  As Long
End Type
Public Type COPYDATASTRUCT
dwData As Long
cbData As Long
lpData As Long
End Type
Public Type OPENFILENAME
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type
Public Type SizeRect
Left As Long
Top As Long
Width As Long
Height As Long
End Type
'Public Type RectAPI
'Left As Long
'Top As Long
'Right As Long
'Bottom As Long
'End Type
Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Public Const DT_CENTER = &H1
Public Type SHFILEOPSTRUCT
hwnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Long
hNameMappings As Long
lpszProgressTitle As Long
End Type
Public Type PALETTEENTRY
peRed As Byte
peGreen As Byte
peBlue As Byte
peFlags As Byte
End Type
Public Type LOGPALETTE
palVersion As Integer
palNumEntries As Integer
palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors.
End Type

Public Type POINT
X As Long
Y As Long
End Type

Public Type picBmp
Size As Long
type As Long
hBmp As Long
hPal As Long
Reserved As Long
End Type
Public Type FormPosition
Left As Long
Top As Long
Width   As Long
Height  As Long
Maxed   As Boolean
End Type
Public Type DEV_BROADCAST_HDR
ISize As Long
IDevicetpe As Long
IReseved As Long
End Type
Type SHELLEXECUTEINFO
CBSIZE As Long
fMask As Long
hwnd As Long
lpVerb As String
lpFile As String
lpParameters As String
lpDirectory As String
nShow As Long
hInstApp As Long
lpIDList As Long ' Optional parameter
lpClass As String ' Optional parameter
hkeyClass As Long ' Optional parameter
dwHotKey As Long ' Optional parameter
hIcon As Long ' Optional parameter
hProcess As Long ' Optional parameter
End Type
Public Type DEV_BROADCAST_VOLUME
ISize As Long
IDevicetype As Long
IReserved As Long
IUntiMask As Long
iFlag As Long
End Type
Public Type LUID
LowPart As Long
HighPart As Long
End Type
Public Type LUID_AND_ATTRIBUTES
pLuid As LUID
Attributes As Long
End Type
Public Type LARGE_INTEGER
LowPart As Long
HighPart As Long
End Type
Private Type CONTEXT
ContextFlags As Long
Dr0 As Long
Dr1 As Long 'DWORD Dr1;
Dr2 As Long 'DWORD Dr2;
Dr3 As Long 'DWORD Dr3;
Dr6 As Long 'DWORD Dr6;
Dr7 As Long 'DWORD Dr7;
FloatSave As FLOATING_SAVE_AREA 'FLOATING_SAVE_AREA FloatSave;
SegGs As Long   'DWORD SegGs;
SegFs As Long   'DWORD SegFs;
SegEs As Long   'DWORD SegEs;
SegDs As Long   'DWORD SegDs;
regEDI As Long  'DWORD Edi;
regESI As Long  'DWORD Esi;
regEBX As Long  'DWORD Ebx;
regEDX As Long  'DWORD Edx;
regECX As Long  'DWORD Ecx;
regEAX As Long  'DWORD Eax;
regEBP As Long  'DWORD Ebp;
regEIP As Long  'DWORD Eip;
SegCs As Long   'DWORD SegCs; // MUST BE SANITIZED
EFlags As Long  'DWORD EFlags; // MUST BE SANITIZED
regESP As Long  'DWORD Esp;
SegSs As Long   'DWORD SegSs;
ExtendedRegisters(MAXIMUM_SUPPORTED_EXTENSION - 1) As Byte ' BYTE ExtendedRegisters[MAXIMUM_SUPPORTED_EXTENSION];
End Type
Private PrevProcPtr As Long '上一个SEH过程地址
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public hNxtHook As Long   ' handle of Hook Procedure
Public Msg As EVENTMSG
Public bL As Single '转换比例变量
Public Cancelflag As Boolean '用户取消标识
Public TargetFile() As FileInfo '存储已转换文件信息的数组
Public TFnum As Long
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
Private Const OFN_FILEMUSTEXIST = &H1000
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Private Declare Function SetSecurityInfo Lib "advapi32.dll" (ByVal handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any) As Long
Private Declare Function GetSecurityInfo Lib "advapi32.dll" (ByVal handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any, ppSecurityDeor As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Public Declare Function CoCreateGuid Lib "ole32" (pGuid As GUID) As Long
Public Declare Function IsEqualGUID Lib "ole32" (pGuid1 As GUID, pGuid2 As GUID) As Long
Public Declare Function StringFromGUID2 Lib "ole32" (pGuid As GUID, ByVal szGuid As String, ByVal cchMax As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal Scan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function TranslateColor Lib "OLEPRO32.DLL" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Public Declare Function SHEmptyRecycleBin Lib "shell32" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal a As String, ByVal fl As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetTickCount& Lib "kernel32" ()
Private Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclA" (ByVal cCountOfExplicitEntries As Long, pListOfExplicitEntries As EXPLICIT_ACCESS, ByVal OldAcl As Long, NewAcl As Long) As Long
Private Declare Sub BuildExplicitAccessWithName Lib "advapi32.dll" Alias "BuildExplicitAccessWithNameA" (pExplicitAccess As EXPLICIT_ACCESS, ByVal pTrusteeName As String, ByVal AccessPermissions As Long, ByVal AccessMode As ACCESS_MODE, ByVal Inheritance As Long)
Private Declare Sub RtlInitUnicodeString Lib "ntdll.dll" (DestinationString As UNICODE_STRING, ByVal SourceString As Long)
Private Declare Function ZwOpenSection Lib "ntdll.dll" (SectionHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As Any) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal lpDefault As String, ByVal lpReturnString As String, ByVal nSize As Long, ByVal filename As String)
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keyDefault$, ByVal filename$)
Public Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Public Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Public Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal ERR As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszfile As String) As Long
'将整个URL参数作为一个URL段
Public Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
Public Const URL_ESCAPE_PERCENT         As Long = &H1000
Public Const URL_UNESCAPE_INPLACE       As Long = &H100000

'路径中包含#
Public Const URL_INTERNAL_PATH          As Long = &H800000
Public Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
Public Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
Public Const URL_DONT_SIMPLIFY          As Long = &H8000000

'转换不安全字符为相应的退格序列
Public Declare Function UrlEscape Lib "shlwapi" _
   Alias "UrlEscapeA" _
  (ByVal pszURL As String, _
   ByVal pszEscaped As String, _
   pcchEscaped As Long, _
   ByVal dwFlags As Long) As Long
Public Declare Function UrlUnescape Lib "shlwapi" _
   Alias "UrlUnescapeA" _
  (ByVal pszURL As String, _
   ByVal pszUnescaped As String, _
   pcchUnescaped As Long, _
   ByVal dwFlags As Long) As Long
   
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, BITMAP As Long) As Long
Public Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, BITMAP As Long) As GpStatus
Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus
Public Declare Function GdipDrawImage Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal X As Single, ByVal Y As Single) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As String, image As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal BITMAP As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal pub_lngInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function InternetTimeToSystemTime Lib "wininet.dll" (ByVal lpszTime As String, ByRef pst As SYSTEMTIME, ByVal dwReserved As Long) As Long
Public Declare Function AddFontResource Lib "gdi32 " Alias "AddFontResourceA " (ByVal lpFileName As String) As Long
Public Declare Function RemoveFontResource Lib "gdi32 " Alias "RemoveFontResourceA " (ByVal lpFileName As String) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINT)
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (udtRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Public Declare Function PdhVbOpenQuery Lib "PDH.DLL" (ByRef QueryHandle As Long) As Long
Public Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As Long
Public Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As Long
Public Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Const WM_SETTINGCHANGE = &H1A
Public Const NIM_ADD = &H0  '添加图标
Public Const NIM_MODIFY = &H1  '修改图标
Public Const NIM_DELETE = &H2  '删除图标
Public Const SPI_SCREENSAVERRUNNING = 97
Dim boTimeOut As Boolean
Public gColors(128) As Long
Public Const DRIVE_CDROM As Long = 5
Public Const DRIVE_REMOVABLE As Long = 2
Public Const FILE_DEVICE_FILE_SYSTEM As Long = 9
Public Const FILE_DEVICE_MASS_STORAGE As Long = &H2D&
Public Const METHOD_BUFFERED As Long = 0
Public Const FILE_ANY_ACCESS As Long = 0
Public Const FILE_READ_ACCESS As Long = 1
Public Const LOCK_VOLUME As Long = 6
Public Const DISMOUNT_VOLUME As Long = 8
Public Const EJECT_MEDIA As Long = &H202
Public Const MEDIA_REMOVAL As Long = &H201
Public Const IDC_HAND  As Long = 32649&
Public hcursor As Long '定义鼠标
Public Const LOCK_TIMEOUT As Long = 1000
Public Const LOCK_RETRIES As Long = 20
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function LoadCursorBynum& Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long)
Public Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByRef dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long '获取硬盘剩余空间的API
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long '这个API是对程序的运行级别进行操作的
Public Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long '这个API是获得硬盘类型的
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long '这个API是指鼠标的显示或隐藏
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Public Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long '这是托盘函数
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetKeyboardState& Lib "user32" (pbKeyState As Byte)
Public Declare Function GetKeyNameText& Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long)
Public Declare Function MapVirtualKey& Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long)
Public Declare Function GetAsyncKeyState% Lib "user32" (ByVal vkey As Long)
Public Declare Function SetWindowWord& Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long)
Public Declare Function GetKeyState% Lib "user32" (ByVal nVirtKey As Long)
Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long '打开
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOPENFILENAME As OPENFILENAME) As Long  '保存
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, pszPath As String) As Long
Public Const ANYSIZE_ARRAY As Long = 1
Public Const EWX_FORCE As Long = 4
Public Const EWX_FORCEIFHUNG As Long = &H10
Public Const EWX_LOGOFF As Long = 0
Public Const EWX_POWEROFF As Long = &H8
Public Const EWX_REBOOT As Long = 2
Public Const EWX_SHUTDOWN As Long = 1
Public Const MAX_COMPUTERNAME As Long = 15
Public Const TOKEN_ADJUST_DEFAULT As Long = &H80
Public Const TOKEN_ADJUST_GROUPS As Long = &H40
Public Const TOKEN_ADJUST_SESSIONID As Long = &H100
Public Const VER_PLATFORM_WIN32_NT As Long = 2
Public Const HWND_TOPMOST As Long = -1
Public QuickKey As Integer '快捷键
Public CKMSG As Integer '定义是否接受服务器信息
Public GETWEATHER As Integer '天气预报
Public Sound As Integer
Public ATP As Integer
Public SPR As Integer
Public lRet As Long
Public I As Integer
Public blnSSDis As Boolean
Public AutoRes As Long
Public NewX, NewY, cx, cy
Public Const RGN_OR = 2
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&
Const MF_STRING = &H0&
Const MF_POPUP = &H10&
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Long  'change lpLuid from LARGE_INTEGER to LUID
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)
Public m_bEditFromCode As Boolean
'定义托盘图标的函数
Public Const NOTIFYICON_VERSION = 3   '风格
Public Const NOTIFYICON_OLDVERSION = 0 'Win95 任务栏样式
Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIIF_INFO = &H1           '   "消息"图标
Public Const NIIF_NOSOUND = &H10
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
Public Const NIS_HIDDEN = &H1
Public Const NIIF_WARNING = &H2           '   "警告"图标
Public Const NIIF_ERROR = &H3           '   "错误"图标
Public Const NIS_SHAREDICON = &H2
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_NULL = &H0
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_SIZING = &H214
Public Const WM_SIZE = &H5
Public Const WM_MOVE = &H3
Public Const PK_TRAYICON = &H401&
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSEMOVE = &H200 '
Public Const WM_MOUSELEAVE = &H2A3 '
Public Const WM_DRAWITEM = &H2B
Public Const WM_SETFONT = &H30
Public Const WM_COMMAND = &H111
Public Const WM_COPYDATA = &H4A
Public Const GWL_USERDATA = (-21&)
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean '执行托盘的API
Public Const FILE_LIST_DIRECTORY = &H1
Public Const FILE_SHARE_READ = &H1&
Public Const FILE_SHARE_DELETE = &H4&
Public Const OPEN_EXISTING = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal PassZero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal PassZero As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public TimerStr(0 To 5) As Integer
Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40
Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (PicDesc As picBmp, REFIID As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Public mfScale As Single
Public mlOldX As Long
Public mlOldY As Long
Public Const DIB_RGB_COLORS As Long = 0
Public Const SRCCOPY As Long = &HCC0020
Public Const PATCOPY As Long = &HF00021
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1 '拖动无边框窗体
Public Const HTCAPTION = 2 '拖动无边框窗体
Public Const ERROR_SUCCESS As Long = 0
Public Const WS_VERSION_REQD  As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR  As Long = -1
Public Const SW_SHOWNORMAL As Long = 1
Public Type HOSTENT
hName As Long
hAliases As Long
hAddrType  As Integer
hLen As Integer
hAddrList  As Long
End Type
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Type WSAData
wVersion As Integer
wHighVersion  As Integer
szDescription(0 To MAX_WSADescription) As Byte
szSystemStatus(0 To MAX_WSASYSStatus) As Byte
wMaxSockets As Integer
wMaxUDPDG  As Integer
dwVendorInfo  As Long
End Type
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF012
Global Const SWP_MOVE = 2
Global Const flags = SWP_NOSIZE Or SWP_MOVE
Declare Function SetWindowsPos Lib "USER" (ByVal H%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Public Const NIIF_GUID = &H4
Public Const SW_SHOW = 5
Public Const WM_USER = &H400
Public Const WM_MBUTTONUP = &H208
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_WNDPROC = (-4)
Public Const MAX_TOOLTIP As Integer = 64
Public Const SW_RESTORE = 9
Public Const SW_HIDE = 0
Public Const ERROR_NONE = 0
'注册表的入口常量
Public Const KEY_ALL_ACCESS = &H3F
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Public Type MEMORYSTATUS
dwLength As Long
dwMemoryLoad As Long
dwTotalPhys As Long
dwAvailPhys As Long
dwTotalPageFile As Long
dwAvailPageFile As Long
dwTotalVirtual As Long
dwAvailVirtual As Long
End Type
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public TheData As NOTIFYICONDATA
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SC_SCREENSAVE = &HF140&
Public Type BROWSEINFO
 hOwner As Long
 pidlRoot As Long
 pszDisplayName As String
 lpszTitle As String
 ulFlags As Long
 lpfn As Long
 lParam As Long
 iImage As Long
End Type
Public Declare Function SetParent Lib "user32" (ByVal hwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Global lpPrevWndProc As Long
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Public Const ILD_TRANSPARENT = &H1 ' Display transparent
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal I&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal flags&) As Long
Public shinfo As SHFILEINFO

Public Declare Function ChangeClipboardChain Lib "user32" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_DRAWCLIPBOARD = &H308
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Const MaxLFNPath = 260
Dim PicHeight%, hLB&, FileSpec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%
Public Running As Boolean
Dim FilesCounter As Integer
Dim WFD As WIN32_FIND_DATA, hItem&, hFile&
Public Const vbBackslash = "\"
Public Const vbAllFiles = "*.*"
Public Const vbKeyDot = 46
Private verinfo As OSVERSIONINFO
Private g_hNtDLL As Long
Private g_pMapPhysicalMemory As Long
Private g_hMPM As Long
Private aByte(3) As Byte
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
'Api播放mid声明
Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
'Api播放wav声明
Public Const SND_SYNC = &H0 '  play synchronously (default)
Public Const SND_ASYNC = &H1 '  play asynchronously
Public Const SND_NODEFAULT = &H2 '  silence not default, if sound not found
'声音占用内存常数
Public Const SND_MEMORY = &H4   '  lpszSoundName points to a memory file
'声音别名常数
Public Const SND_ALIAS = &H10000 '  name is a WIN.INI [sounds] entry
'声音文件名常数
Public Const SND_FILENAME = &H20000 '  name is a file name
Public Const SND_RESOURCE = &H40004 '  name is a resource name or atom
'声音别名标识常数
Public Const SND_ALIAS_ID = &H110000 '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0 '  must be > 4096 to keep strings in same section of resource file
Public Const SND_LOOP = &H8 '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10  '  don't stop any currently playing sound
Public Const SND_VALID = &H1F   '  valid flags  / ;Internal /
Public Const SND_NOWAIT = &H2000 '  don't wait if the driver is busy
'声音有效标志常数
Public Const SND_VALIDFLAGS = &H17201F  '  Set of valid flag bits.  Anything outside
'声音保留常数
Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Public Const SND_TYPE_MASK = &H170007
Public Const SPI_GETWORKAREA = 48
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64
Type NEWTEXTMETRIC
tmHeight As Long
tmAscent As Long
tmDescent As Long
tmInternalLeading As Long
tmExternalLeading As Long
tmAveCharWidth As Long
tmMaxCharWidth As Long
tmWeight As Long
tmOverhang As Long
tmDigitizedAspectX As Long
tmDigitizedAspectY As Long
tmFirstChar As Byte
tmLastChar As Byte
tmDefaultChar As Byte
tmBreakChar As Byte
tmItalic As Byte
tmUnderlined As Byte
tmStruckOut As Byte
tmPitchAndFamily As Byte
tmCharSet As Byte
ntmFlags As Long
ntmSizeEM As Long
ntmCellHeight As Long
ntmAveWidth As Long
End Type
Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Public hwndNextViewer As Long

Public Sub Main()
Ary = Array("", "Windows 95", "Windows 98", "Windows Me", "Windows NT4.0", "Windows 2000", "Windows XP", "Windows 2003", "Windows Vista", "Windows 7")
'If UCase(Ary(GetVersion)) <> "WINDOWS XP" Then Call SHOWWRONG("抱歉,本程序只在Windows XP系统上运行完美", 0): Exit Sub
'由于现在很少有窗体用阴影控件了,所以直接把这句备注
RUN_MODE = GetInitEntry("SYSTEM", "AUTORUN", 0)
COLOR_NOR = vbBlack
COLOR_HIGH = &HAD7900
ALWAYSONTOP = GetInitEntry("SYSTEM", "ONTOP", False)
CAN_MINI = GetInitEntry("SYSTEM", "MINI_PLAYER", True)
AUTO_SINGER = GetInitEntry("SYSTEM", "AUTO_S", True)
On Error Resume Next '防止错误直接崩溃
Clipboard.Clear
Dim Kiss, ISS, RSS, fst As Integer
Dim hwndold As Long, YY As Long, LOCALEXENAMe As String
Kiss = Left(App.Path, 3)
RSS = GetDriveType(Kiss)
Sound = GetSetting("ICEE", "Main", "SOUND", 1)
READYLOAD = True
If RSS = 3 Then '开始运行环境的验证
LOCALEXENAMe = App.Path & "\" & App.exename & ".EXE"
Call 加壳(LOCALEXENAMe, LOCALEXENAMe)
Call GetVer
File_Share_Flag = 0
KBS = 0
hwndold = FindWindowEx(0, 0, "ThunderRT6FormDC", "ICEE") '禁止多个ICEE运行(即使你改了文件名，换了路径.都无法再次运行)
YY = FindWindowEx(0, 0, "ThunderRT6FormDC", "安全验证") '同样
fst = GetSetting("ICEE", "Main", "FIRSTUSE", 0) '初始化第一次使用的函数
NEWS = GetSetting("ICEE", "Main", "news", 1) '是否打开每日资讯
fn = GetSetting("ICEE", "Main", "passyn", 0) '是否使用了密码
AutoRes = GetInitEntry("LocalSafe", "SelfHelp", 1) '是否启动了修复快捷键
If hwndold <> 0 Or YY <> 0 Then '如果不成功的话，则结束(成功的值为0，不成功则为不为0的数字)
With Wrn
.Move (Screen.Width - Wrn.Width) / 2, (Screen.Height - Wrn.Height) / 2
.ts.Caption = "  尊敬的用户,为了保证数据数据被恶意更改,禁止在同一系统上同时运行多个程序"
.DRAWINFOICO 2
.Show vbModal
End With
Exit Sub
End
Else
Call HideCurrentProcess
Call OutPutMain '先创建文件夹及皮肤.
Call OutPutSound '后输出音频.
If fn = 0 Then '如果没有密码的话
Select Case RUN_MODE
Case 0
LONELY_MODE = False
If Command$ = "" Then frmma.Show
Select Case UCase(Right(Command$, 3))
Case "BMP", "JPG", "GIF", "PNG", "PSD"
LONELY_MODE = True
FRMBOARD.Show
Call FRMBOARD.OpenFile(Command$)
Case "MP3"
Call frmma.PlayMusic(Command$)
Case "M3U"
Call frmma.Playlist(Command$)
End Select
Case 1
LONELY_MODE = True
FRMBOARD.Show
Case 2
LONELY_MODE = True
frmGraphic.Show
Case 3
LONELY_MODE = True
FRMEX.Show
End Select
Else '如果有密码的话
frmpass.Show
End If '结束密码的验证
If AutoRes = 1 Then Call mShellLnk("ICEE", App.Path & "\" & App.exename, , App.Path & "\" & App.exename & ".EXE", , "ICEE")  '修复快捷键
End If '结束重复运行的验证
Else
Call SHOWWRONG("  对不起，为避免程序运行过慢，请不要让ICEE在移动设备上运行", 2)
Exit Sub
End If ' 不要在移动设备上运行
End Sub
Public Sub Guanlian(ByVal AUTO As Boolean)
Dim hKey As Long, ExePath As String, ExeIcon As String '注册文件关联的
If AUTO Then
ExePath = App.Path & "\" & App.exename & ".exe %1"
If RegOpenKey(HKEY_CLASSES_ROOT, ".mp3", hKey) <> 0 Then
Call RegCreateKey(HKEY_CLASSES_ROOT, ".mp3", hKey)
Call RegSetValue(HKEY_CLASSES_ROOT, ".mp3", REG_SZ, "mywfile", 9)
Call RegCreateKey(HKEY_CLASSES_ROOT, ".mp3\shellnew", hKey)
Call RegSetValueEx(hKey, "NullFile", "0", REG_SZ, "", 0)
Call RegSetValue(HKEY_CLASSES_ROOT, "mp3", REG_SZ, "音频文件", LenB(StrConv("系统日志", vbFromUnicode)) + 1)
End If
Call RegSetValue(HKEY_CLASSES_ROOT, ".mp3", REG_SZ, "mywfile", 9)
ExeIcon = App.Path & "\" & App.exename & ".exe,0"
Call RegSetValue(HKEY_CLASSES_ROOT, "mywfile\DefaultIcon", REG_SZ, ExeIcon, LenB(StrConv(ExeIcon, vbFromUnicode)) + 1)
Call RegSetValue(HKEY_CLASSES_ROOT, "mywfile\shell\open\command", REG_SZ, ExePath, LenB(StrConv(ExePath, vbFromUnicode)) + 1)
Else
DeleteKey HKEY_CLASSES_ROOT, "", "mywfile"
DeleteKey HKEY_CLASSES_ROOT, "", ".mp3"
End If
UpgradeDesktop (&H1E)
UpgradeDesktop (&H20)
RegCloseKey hKey
End Sub
Public Sub UpgradeDesktop(IcoSize As String)
Dim hKey As Long
RegOpenKeyEx HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", 0, 0, hKey
RegCreateKey HKEY_CURRENT_USER, "Control Panel\Desktop\WindowMetrics", hKey
RegSetValueEx hKey, "Shell Icon Size", 0, 1, ByVal IcoSize, 2
Call SendMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0)
End Sub
Function DeleteKey(RootKey As Long, ParentKeyName As String, SubKeyName As String)
Dim hKey As Long
RegOpenKey RootKey, ParentKeyName, hKey
RegDeleteKey hKey, SubKeyName
RegCloseKey hKey
End Function
Public Sub NoNoNo()
On Error Resume Next
If IS_SET = True Then Exit Sub
frmmp.Hide
Frmm.TimeHon.Enabled = False
With frmma
Call .HIDEBK
.Hide
.Timers.Enabled = False
End With
If frmma.TMP.Enabled = False Then Exit Sub
'    With FRMTASK.PTASK
'    .Top = 0
'    .Left = 0
'    .Width = FRMTASK.ScaleWidth
'    .Height = FRMTASK.ScaleHeight
'    End With
If CAN_MINI = False Then Exit Sub
FRMTASK.Show
End Sub
Sub DataNew()
On Error Resume Next
Call Frmm.CHECKNET
If Status.RasConnState <> &H2000 Then Exit Sub
If Left(IEver, 1) >= 7 Then FRMNEWS.Show
End Sub
Public Sub StartMonitoring(ByVal hwnd As Long)
hwndNextViewer = SetClipboardViewer(hwnd)
End Sub

Public Sub StopMonitoring(ByVal hwnd As Long)
If hwndNextViewer <> 0 Then
Call ChangeClipboardChain(hwnd, hwndNextViewer)
End If
End Sub
Public Function FileExists(sFName As String) As Boolean
On Error GoTo handler
FileAttr sFName
FileExists = True
Exit Function
handler:
FileExists = False
End Function
Public Function ShowProperties(filename As String, OwnerhWnd As Long) As Long
Dim SEI As SHELLEXECUTEINFO
Dim r As Long
With SEI
.CBSIZE = Len(SEI)
.fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
.hwnd = OwnerhWnd
.lpVerb = "properties"
.lpFile = filename
.lpParameters = vbNullChar
.lpDirectory = vbNullChar
.nShow = 0
.hInstApp = 0
.lpIDList = 0
End With
r = ShellExecuteEX(SEI)
ShowProperties = SEI.hInstApp

End Function
Public Sub mShellLnk(ByVal LnkName As String, ByVal FILEPATH As String, Optional ByVal StrArg As String, Optional ByVal IconFileIconIndex As String = vbNullString, Optional ByVal HookKey As String = "", Optional ByVal StrRemark As String = "")
'调用说明
'LnkName = 快捷方式文件名,如果无路径则自动新建到桌面;无后缀名(.lnk)会自动补齐.
'FilePath = 目标文件名,全路径.
'StrArg = 参数,可选.
'IconFileIconIndex = 图标所在库及索引,由逗号分隔,可选.如: "c:\windows\system32\notepad.exe,0"
'HookKey = 热键,值未知,可选.
'StrRemark = 备注,可选.
Dim WshShell As Object, oShellLink As Object, strDesktop As String
Set WshShell = CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")  '桌面路径
If UCase(Right(LnkName, 4)) <> ".LNK" Then
LnkName = LnkName & ".lnk"
End If
If InStr(1, LnkName, "\", vbTextCompare) = 0 Then   '如果不包含全路径,则在桌面创建快捷方式
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & LnkName)
Else '否则在指定位置创建
Set oShellLink = WshShell.CreateShortcut(LnkName)
End If
oShellLink.TargetPath = FILEPATH
oShellLink.Arguments = StrArg
oShellLink.WindowStyle = 1   '风格
oShellLink.HotKey = HookKey   '热键
If IconFileIconIndex = vbNullString Then   '图标
oShellLink.IconLocation = FILEPATH & ",0"   '默认使用目标文件图标
Else
oShellLink.IconLocation = IconFileIconIndex
End If

oShellLink.Description = StrRemark   '快捷方式备注内容
oShellLink.WorkingDirectory = Mid(FILEPATH, 1, InStrRev(FILEPATH, "\")) '源文件所在目录
oShellLink.Save '保存创建的快捷方式

Set WshShell = Nothing
Set oShellLink = Nothing
End Sub

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
ByVal FontType As Long, lParam As LISTBOX) As Long 'Make font parameters
Dim FaceName As String
Dim FullName As String
On Error Resume Next
FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
EnumFontFamProc = 1
End Function 'EnumFontFamProc
Sub FillListWithFonts(LB As LISTBOX)     'Adds system fonts to list box
Dim hdc As Long
On Error Resume Next
LB.Clear
hdc = GetDC(LB.hwnd)
EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, LB
ReleaseDC LB.hwnd, hdc
End Sub
Sub FillComboWithFonts(cb As ComboBox)
Dim hdc As Long
On Error Resume Next
cb.Clear
hdc = GetDC(cb.hwnd)
EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, cb
ReleaseDC cb.hwnd, hdc
End Sub
Public Function GetTaskbarHeight() As Integer
Dim lRes As Long
Dim rectVal As RECT
lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function
Public Sub PlayWav(SoundName As String)
Dim tmpSoundName As String
Dim wFlags As Long
Dim X As Long
  tmpSoundName = SoundName
  wFlags = SND_ASYNC Or SND_LOOP
  X = sndPlaySound(tmpSoundName, wFlags)
End Sub
Public Function BrowseFolder(ByVal aTitle As String, ByVal aForm As Form) As String
Dim bInfo As BROWSEINFO, t As String
Dim rtn&, pidl&, Path$, pos%, Browse
Dim BrowsePath As String
bInfo.hOwner = aForm.hwnd
bInfo.lpszTitle = aTitle
bInfo.ulFlags = &H40
pidl& = SHBrowseForFolder(bInfo)
Path = Space(512)
t = SHGetPathFromIDList(ByVal pidl&, ByVal Path)
pos% = InStr(Path$, Chr$(0))
BrowseFolder = Left(Path$, pos - 1)
If Right$(Browse, 1) = "\" Then
BrowseFolder = BrowseFolder
Else
BrowseFolder = BrowseFolder + "\"
End If
If Right(BrowseFolder, 2) = "\\" Then BrowseFolder = Left(BrowseFolder, Len(BrowseFolder) - 1)
If BrowseFolder = "\" Then BrowseFolder = ""
End Function
Public Function ShortName(LongPath As String) As String
Dim ShortPath As String
Const MAX_PATH = 260
Dim Ret&
ShortPath = Space$(MAX_PATH)
'取得短文件名.
Ret& = GetShortPathName(LongPath, ShortPath, MAX_PATH)
If Ret& Then
ShortName = Left$(ShortPath, Ret&)
End If
End Function
'屏保开启
Public Sub SetSSEnabled(blnEnable As Boolean)
Dim lFlag As Long
lFlag = IIf(blnEnable, 1&, 0&) '1 允许, 0 禁止
Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, lFlag, 0&, 0&)
End Sub
'运行当前屏保
Public Sub StartSS()
Dim lDesktop As Long
lDesktop = GetDesktopWindow()
'发送消息执行屏保
SendMessage lDesktop, WM_SYSCOMMAND, SC_SCREENSAVE, 0&
End Sub
Sub OutPutMain()
On Error Resume Next
If Dir(App.Path + "\Cofing", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Cofing")
If Dir(App.Path + "\Cofing\CPU", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Cofing\CPU")

If Dir(App.Path + "\Skin", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Skin")
If Dir(App.Path + "\Skin\PHOTO", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Skin\PHOTO")
If Dir(App.Path + "\Skin\BK", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Skin\BK")

If Dir(App.Path + "\Sound", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Sound")

If Dir(App.Path + "\Thumbs", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Thumbs")
If Dir(App.Path + "\Thumbs\Singer_Thumbs", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Thumbs\Singer_Thumbs")
If Dir(App.Path + "\Download", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\DOWNLOAD")

If Dir(App.Path + "\Media", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media")
If Dir(App.Path + "\Media\MusicPicture", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media\MusicPicture")
If Dir(App.Path + "\Media\LRC", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media\Lrc")
If Dir(App.Path + "\Media\Paint", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media\Paint")
If Dir(App.Path + "\Media\MUSICBOX", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media\MusicBox")
If Dir(App.Path + "\Media\SEARCH", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media\Search")
If Dir(App.Path + "\Media\FAVOURITE", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media\Favourite")
If Dir(App.Path + "\Media\PIC", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\Media\PIC")

If Dir(App.Path + "\USER", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\User")
If Dir(App.Path + "\User\Act", vbNormal Or vbHide Or vbSystem Or vbReadOnly Or vbDirectory Or vbArchive) = "" Then MkDir (App.Path + "\User\Act")
DoEvents
Dim BY() As Byte, I As Integer

For I = 0 To 8
BY = LoadResData(I + 289, "CUSTOM")
Open App.Path + "\MEDIA\PIC\" & I & ".JPG" For Binary As #1
Put #1, , BY
Close #1
Next

For I = 0 To 9
BY = LoadResData(I + 321, "CUSTOM")
Open App.Path + "\SKIN\BT" & I & ".PNG" For Binary As #1
Put #1, , BY
Close #1
Next

BY = LoadResData(331, "CUSTOM")
Open App.Path + "\SKIN\BP.PNG" For Binary As #1
Put #1, , BY
Close #1

For I = 0 To 8
BY = LoadResData(I + 289, "CUSTOM")
Open App.Path + "\SKIN\PHOTO\WALL" & I & ".JPG" For Binary As #1
Put #1, , BY
Close #1
Next

For I = 0 To 5
BY = LoadResData(I + 309, "CUSTOM")
Open App.Path + "\SKIN\AB" & I & ".PNG" For Binary As #1
Put #1, , BY
Close #1
Next


For I = 0 To 9
BY = LoadResData(I + 267, "CUSTOM")
Open App.Path + "\SKIN\T" & I & ".PNG" For Binary As #1
Put #1, , BY
Close #1
Next

For I = 0 To 7
BY = LoadResData(I + 259, "CUSTOM")
Open App.Path + "\Skin\LOADING" & I & ".png" For Binary As #1
Put #1, , BY
Close #1
Next

For I = 0 To 15
BY = LoadResData(I + 332, "CUSTOM")
Open App.Path + "\Skin\BK\" & I + 10 & ".JPG" For Binary As #1
Put #1, , BY
Close #1
Next
For I = 0 To 9
BY = LoadResData(I + 299, "CUSTOM")
Open App.Path + "\SKIN\BK\" & I & ".JPG" For Binary As #1
Put #1, , BY
Close #1
Next
BY = LoadResData("LOADING", "CUSTOM")
Open App.Path + "\Skin\LOADING.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("S_TIP", "CUSTOM")
Open App.Path + "\Skin\S_TIP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("S_MARK", "CUSTOM")
Open App.Path + "\Skin\S_MARK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DELETE", "CUSTOM")
Open App.Path + "\Skin\DELETE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ADD_ALL", "CUSTOM")
Open App.Path + "\Skin\ADD_ALL.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("REFRESH", "CUSTOM")
Open App.Path + "\Skin\REFRESH.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NT_S", "CUSTOM")
Open App.Path + "\Skin\NT_S.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MI_N", "CUSTOM")
Open App.Path + "\Skin\MI_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SS_N", "CUSTOM")
Open App.Path + "\Skin\SS_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SS_H", "CUSTOM")
Open App.Path + "\Skin\SS_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MI_H", "CUSTOM")
Open App.Path + "\Skin\MI_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DD_N", "CUSTOM")
Open App.Path + "\Skin\DD_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DD_H", "CUSTOM")
Open App.Path + "\Skin\DD_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SEND_H", "CUSTOM")
Open App.Path + "\Skin\SEND_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SEND_N", "CUSTOM")
Open App.Path + "\Skin\SEND_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DE_N", "CUSTOM")
Open App.Path + "\Skin\DE_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DE_H", "CUSTOM")
Open App.Path + "\Skin\DE_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("KU_N", "CUSTOM")
Open App.Path + "\Skin\KU_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("KU_H", "CUSTOM")
Open App.Path + "\Skin\KU_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FA_T", "CUSTOM")
Open App.Path + "\Skin\FA_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FA_F", "CUSTOM")
Open App.Path + "\Skin\FA_F.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("T_P", "CUSTOM")
Open App.Path + "\Skin\T_P.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("I_ON", "CUSTOM")
Open App.Path + "\Skin\I_ON.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("I_OFF", "CUSTOM")
Open App.Path + "\Skin\I_OFF.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LOAD_FM", "CUSTOM")
Open App.Path + "\Skin\LOAD_FM.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CPU_BK", "CUSTOM")
Open App.Path + "\Skin\CPU_BK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DISK", "CUSTOM")
Open App.Path + "\Skin\DISK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SQ", "CUSTOM")
Open App.Path + "\Skin\SQ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SHARE_N", "CUSTOM")
Open App.Path + "\Skin\SHARE_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SHARE_H", "CUSTOM")
Open App.Path + "\Skin\SHARE_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FORM", "CUSTOM")
Open App.Path + "\Skin\FORM.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FAV_NO", "CUSTOM")
Open App.Path + "\Skin\FAV_NO.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FAV_YES", "CUSTOM")
Open App.Path + "\Skin\FAV_YES.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FAV_T", "CUSTOM")
Open App.Path + "\Skin\FAV_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SET", "THUM4")
Open App.Path + "\Skin\D_SET.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DA", "CUSTOM")
Open App.Path + "\Skin\DA.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DF", "CUSTOM")
Open App.Path + "\Skin\DF.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ZI_N", "CUSTOM")
Open App.Path + "\Skin\ZI_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PT_S", "CUSTOM")
Open App.Path + "\Skin\PT_S.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("ZI_H", "CUSTOM")
Open App.Path + "\Skin\ZI_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ZO_N", "CUSTOM")
Open App.Path + "\Skin\ZO_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ZO_H", "CUSTOM")
Open App.Path + "\Skin\ZO_H.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("AD", "CUSTOM")
Open App.Path + "\Skin\AD.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LG_N", "CUSTOM")
Open App.Path + "\Skin\LG_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LG_H", "CUSTOM")
Open App.Path + "\Skin\LG_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DR", "CUSTOM")
Open App.Path + "\Skin\DR.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("FM_T", "CUSTOM")
Open App.Path + "\Skin\FM_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("UP_T", "CUSTOM")
Open App.Path + "\Skin\UP_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ZIN", "CUSTOM")
Open App.Path + "\Skin\ZIN.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ZOUT", "CUSTOM")
Open App.Path + "\Skin\ZOUT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NO_FAV", "CUSTOM")
Open App.Path + "\Skin\NO_FAV.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NHQ", "CUSTOM")
Open App.Path + "\Skin\NHQ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FM", "CUSTOM")
Open App.Path + "\Skin\FM.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PAUSE", "CUSTOM")
Open App.Path + "\Skin\PAUSE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PLAY", "CUSTOM")
Open App.Path + "\Skin\PLAY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SYS", "CUSTOM")
Open App.Path + "\Skin\SYS.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("WHITE", "CUSTOM")
Open App.Path + "\Skin\WHITE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NEXT", "CUSTOM")
Open App.Path + "\Skin\NEXT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("P_LIST", "CUSTOM")
Open App.Path + "\Skin\P_LIST.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SKIN", "CUSTOM")
Open App.Path + "\Skin\SKIN.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("HDSK", "CUSTOM")
Open App.Path + "\Skin\HDSK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FD", "CUSTOM")
Open App.Path + "\Skin\FD.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DEL", "CUSTOM")
Open App.Path + "\Skin\DEL.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("EMP", "CUSTOM")
Open App.Path + "\Skin\EMP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ADD", "CUSTOM")
Open App.Path + "\Skin\ADD.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MORE", "CUSTOM")
Open App.Path + "\Skin\MORE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("STD", "CUSTOM")
Open App.Path + "\Skin\STD.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("QC", "CUSTOM")
Open App.Path + "\Skin\QC.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("UP", "CUSTOM")
Open App.Path + "\Skin\UP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("HELP", "CUSTOM")
Open App.Path + "\Skin\HELP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DLRC", "CUSTOM")
Open App.Path + "\Skin\DLRC.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SE", "CUSTOM")
Open App.Path + "\Skin\SE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SJ", "CUSTOM")
Open App.Path + "\Skin\SJ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SX", "CUSTOM")
Open App.Path + "\Skin\SX.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LB", "CUSTOM")
Open App.Path + "\Skin\LB.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DQ", "CUSTOM")
Open App.Path + "\Skin\DQ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CPU", "CUSTOM")
Open App.Path + "\Skin\CPU.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MEM", "CUSTOM")
Open App.Path + "\Skin\MEM.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("BD", "CUSTOM")
Open App.Path + "\Skin\BD.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("GS", "CUSTOM")
Open App.Path + "\Skin\GS.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SMSG", "ICO")
Open App.Path + "\Skin\SMSG.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("OFF", "ICO")
Open App.Path + "\Skin\OFF.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("EXP", "ICO")
Open App.Path + "\Skin\EXP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("OPEN", "ICO")
Open App.Path + "\Skin\OPEN.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PIC", "ICO")
Open App.Path + "\Skin\PIC.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MUS", "ICO")
Open App.Path + "\Skin\MUS.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("EXE", "ICO")
Open App.Path + "\Skin\EXE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("VID", "ICO")
Open App.Path + "\Skin\VID.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("OT", "ICO")
Open App.Path + "\Skin\OT.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("FMSG", "ICO")
Open App.Path + "\Skin\FMSG.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CZ", "ICO")
Open App.Path + "\Skin\CZ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PIN", "CUSTOM")
Open App.Path + "\Skin\PIN.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SIN", "CUSTOM")
Open App.Path + "\Skin\SIN.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("UN_CO", "CUSTOM")
Open App.Path + "\Skin\UN_CO.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("UI_LOAD", "CUSTOM")
Open App.Path + "\Skin\UI_LOAD.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("TIP", "CUSTOM")
Open App.Path + "\Skin\TIP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LINE", "CUSTOM")
Open App.Path + "\Skin\LINE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NO_PIC", "CUSTOM")
Open App.Path + "\Skin\NO_PIC.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FRAME", "CUSTOM")
Open App.Path + "\Skin\FRAME.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("Pass", "CUSTOM")
Open App.Path + "\Skin\Pass.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MSG_ASK", "MSG")
Open App.Path + "\Skin\MSG_ASK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MSG_INFO", "MSG")
Open App.Path + "\Skin\MSG_INFO.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("D_H", "UI")
Open App.Path + "\Skin\D_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("D_N", "UI")
Open App.Path + "\Skin\D_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("L_H", "UI")
Open App.Path + "\Skin\L_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("L_N", "UI")
Open App.Path + "\Skin\L_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MB_N", "UI")
Open App.Path + "\Skin\MB_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MB_H", "UI")
Open App.Path + "\Skin\MB_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("D_T", "CUSTOM")
Open App.Path + "\Skin\D_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("HD_T", "CUSTOM")
Open App.Path + "\Skin\HD_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("P_T", "CUSTOM")
Open App.Path + "\Skin\P_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CTS", "CUSTOM")
Open App.Path + "\Skin\CTS.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("S_T", "CUSTOM")
Open App.Path + "\Skin\S_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("M_T", "CUSTOM")
Open App.Path + "\Skin\M_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FC_T", "CUSTOM")
Open App.Path + "\Skin\FC_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ID_T", "CUSTOM")
Open App.Path + "\Skin\ID_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PH_T", "CUSTOM")
Open App.Path + "\Skin\PH_T.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("N_T", "CUSTOM")
Open App.Path + "\Skin\N_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ND_T", "CUSTOM")
Open App.Path + "\Skin\ND_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CF_T", "CUSTOM")
Open App.Path + "\Skin\CF_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("A_T", "CUSTOM")
Open App.Path + "\Skin\A_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("HS_T", "CUSTOM")
Open App.Path + "\Skin\HS_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("FL_T", "CUSTOM")
Open App.Path + "\Skin\FL_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DA_T", "CUSTOM")
Open App.Path + "\Skin\DA_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("W_T", "CUSTOM")
Open App.Path + "\Skin\W_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("I_T", "CUSTOM")
Open App.Path + "\Skin\I_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("H_T", "CUSTOM")
Open App.Path + "\Skin\H_T.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("MSG_DONE", "MSG")
Open App.Path + "\Skin\MSG_DONE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MSG_WRN", "MSG")
Open App.Path + "\Skin\MSG_WRN.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("INSH", "CUSTOM")
Open App.Path + "\Skin\INSH.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("POINT", "CUSTOM")
Open App.Path + "\SKIN\POINT.PNG" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NO_FIND", "CUSTOM")
Open App.Path + "\SKIN\NO_FIND.PNG" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("L_SHD", "CUSTOM")
Open App.Path + "\SKIN\L_SHD.PNG" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LOGO_65", "CUSTOM")
Open App.Path + "\Skin\LOGO_65.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("BK_N", "CUSTOM")
Open App.Path + "\Skin\BK_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("BK_H", "CUSTOM")
Open App.Path + "\Skin\BK_H.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("BAOG", "ICO")
Open App.Path + "\Skin\BAOG.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("PRO", "ICO")
Open App.Path + "\Skin\PRO.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SAV", "ICO")
Open App.Path + "\Skin\SAV.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CPT", "ICO")
Open App.Path + "\Skin\CPT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("TCP", "ICO")
Open App.Path + "\Skin\TCP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SOFT", "ICO")
Open App.Path + "\Skin\SOFT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("WIN", "ICO")
Open App.Path + "\Skin\WIN.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("EWM", "ICO")
Open App.Path + "\Skin\EWM.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LVJ", "ICO")
Open App.Path + "\Skin\LVJ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("XUANZ", "ICO")
Open App.Path + "\Skin\XUANZ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("VIEW", "ICO")
Open App.Path + "\Skin\VIEW.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("EMPLY", "CUSTOM")
Open App.Path + "\Skin\EMPLY.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("C_LOAD", "CUSTOM")
Open App.Path + "\Skin\C_LOAD.GIF" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DQ_N", "PLAYER")
Open App.Path + "\Skin\DQ_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DQ_H", "PLAYER")
Open App.Path + "\Skin\DQ_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NX_N", "PLAYER")
Open App.Path + "\Skin\NX_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NX_H", "PLAYER")
Open App.Path + "\Skin\NX_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("P_H", "PLAYER")
Open App.Path + "\Skin\P_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("P_N", "PLAYER")
Open App.Path + "\Skin\P_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PA_N", "PLAYER")
Open App.Path + "\Skin\PA_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PA_h", "PLAYER")
Open App.Path + "\Skin\pa_h.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PR_N", "PLAYER")
Open App.Path + "\Skin\PR_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PR_H", "PLAYER")
Open App.Path + "\Skin\PR_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SJ_N", "PLAYER")
Open App.Path + "\Skin\SJ_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SJ_H", "PLAYER")
Open App.Path + "\Skin\SJ_H.png" For Binary As #1
Put #1, , BY
Close #1


BY = LoadResData("SX_H", "PLAYER")
Open App.Path + "\Skin\SX_H.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("SX_N", "PLAYER")
Open App.Path + "\Skin\SX_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("VOL", "PLAYER")
Open App.Path + "\Skin\VOL.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("XH_H", "PLAYER")
Open App.Path + "\Skin\XH_H.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("XH_N", "PLAYER")
Open App.Path + "\Skin\XH_N.png" For Binary As #1
Put #1, , BY
Close #1









BY = LoadResData("CAL", "CUSTOM")
Open App.Path + "\Skin\CAL.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("M_LOAD", "CUSTOM")
Open App.Path + "\Skin\M_LOAD.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("BACK", "CUSTOM")
Open App.Path + "\Skin\BACK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SHB", "CUSTOM")
Open App.Path + "\Skin\SHB.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PLAYBOX", "CUSTOM")
Open App.Path + "\Skin\PLAYBOX.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SHD_TXT", "CUSTOM")
Open App.Path + "\Skin\SHD_TXT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SINGER_D", "CUSTOM")
Open App.Path + "\Skin\SINGER_D.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PO_T", "CUSTOM")
Open App.Path + "\Skin\PO_T.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("AWAY27", "STATUS")
Open App.Path + "\Skin\AWAY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ONLINE27", "STATUS")
Open App.Path + "\Skin\ONLINE27.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("HIDE27", "STATUS")
Open App.Path + "\Skin\HIDE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("OFFLINE27", "STATUS")
Open App.Path + "\Skin\OFFLINE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("BUSY17", "STATUS")
Open App.Path + "\Skin\BUSY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ABOUTME", "CUSTOM")
Open App.Path + "\Skin\ABOUTME.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("M_LIST", "CUSTOM")
Open App.Path + "\Skin\M_LIST.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ISAY", "CUSTOM")
Open App.Path + "\Skin\ISAY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("TM", "CUSTOM")
Open App.Path + "\Skin\TM.png" For Binary As #1
Put #1, , BY
Close #1

BY = LoadResData("IMUN", "CUSTOM")
Open App.Path + "\Skin\IMUN.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("IMBUG", "CUSTOM")
Open App.Path + "\Skin\IMBUG.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("IMINFO", "CUSTOM")
Open App.Path + "\Skin\IMINFO.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("IMPASS", "CUSTOM")
Open App.Path + "\Skin\IMPASS.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("USAY", "CUSTOM")
Open App.Path + "\Skin\USAY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PHOTOSET", "CUSTOM")
Open App.Path + "\Skin\PHOTOSET.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LOCKED", "CUSTOM")
Open App.Path + "\Skin\LOCKED.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("OTHERSAY", "CUSTOM")
Open App.Path + "\Skin\OTHERSAY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("DL_N", "CUSTOM")
Open App.Path + "\Skin\DL_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SET_H", "CUSTOM")
Open App.Path + "\Skin\SET_H.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SET_N", "CUSTOM")
Open App.Path + "\Skin\SET_N.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LOGO", "CUSTOM")
Open App.Path + "\Skin\LOGO.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MINI_P", "CUSTOM")
Open App.Path + "\Skin\MINI_P.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("STTIT", "CUSTOM")
Open App.Path + "\Skin\STTIT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SER_NT", "CUSTOM")
Open App.Path + "\Skin\SER_NT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("HQ", "CUSTOM")
Open App.Path + "\Skin\HQ.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ONLINE", "CUSTOM")
Open App.Path + "\Skin\ONLINE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SINGER", "CUSTOM")
Open App.Path + "\Skin\SINGER.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("UI_TIT", "CUSTOM")
Open App.Path + "\Skin\UI_TIT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LIST", "CUSTOM")
Open App.Path + "\Skin\LIST.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("INSH", "CUSTOM")
Open App.Path + "\Skin\INSH.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("QMLIST", "PLUG")
Open App.Path + "\QMLISTBOX.OCX" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MSWINSCK", "PLUG")
Open App.Path + "\MSWINSCK.OCX" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MSCOMCTL", "PLUG")
Open App.Path + "\MSCOMCTL.OCX" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SHADOW", "PLUG")
Open App.Path + "\Shadow.OCX" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("WINSUBHOOK", "PLUG")
Open App.Path + "\WINSUBHOOK.tlb" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("IE", "PLUG")
Open App.Path + "\IE.REG" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LOGIN", "CUSTOM")
Open App.Path + "\Skin\LOGIN.png" For Binary As #1
Put #1, , BY
Close #1
Call OUTPUTTHUMB
BY = LoadResData("UNDER", "CUSTOM")
Open App.Path + "\Skin\UNDER.png" For Binary As #1
Put #1, , BY
Close #1
Dim BOY() As Byte
BOY = LoadResData("HEAD", "CUSTOM")
Open App.Path + "\Skin\HEAD.png" For Binary As #1
Put #1, , BOY
Close #1
Dim BOXER() As Byte
BOXER = LoadResData("DIP", "CUSTOM")
Open App.Path + "\Skin\DEAP.png" For Binary As #1
Put #1, , BOXER
Close #1
Dim ABC() As Byte
ABC = LoadResData("ABOUT", "TEXT")
Open App.Path + "\COFING\WHATNEW.TXT" For Binary As #1
Put #1, , ABC
Close #1
Dim BBG() As Byte
BBG = LoadResData("hand", "CUSTOM")
Open (App.Path + "\Skin\DefaultHead.Bmp") For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("CHATBK", "CUSTOM")
Open (App.Path + "\Skin\CHAT.PNG") For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("EMO", "GIF")
Open App.Path + "\SKIN\EMO.GIF" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("WIN", "GIF")
Open App.Path + "\SKIN\WIN.GIF" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("LOSE", "GIF")
Open App.Path + "\SKIN\LOSE.GIF" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("OKOKOK", "GIF")
Open App.Path + "\SKIN\OKOKOK.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("NONONO", "GIF")
Open App.Path + "\SKIN\NONONO.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("SHOP", "GIF")
Open App.Path + "\SKIN\SHOP.GIF" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("BORING", "FACE")
Open App.Path + "\SKIN\BORING.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("FUCK", "FACE")
Open App.Path + "\SKIN\FUCK.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("GOOD_M", "FACE")
Open App.Path + "\SKIN\GOOD_M.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("GOOD_N", "FACE")
Open App.Path + "\SKIN\GOOD_N.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("NO_WORD", "FACE")
Open App.Path + "\SKIN\NO_WORD.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("OMG", "FACE")
Open App.Path + "\SKIN\OMG.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("SHINE", "FACE")
Open App.Path + "\SKIN\SHINE.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("SHOW_LV", "FACE")
Open App.Path + "\SKIN\SHOW_LV.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("SHY", "FACE")
Open App.Path + "\SKIN\SHY.PNG" For Binary As #1
Put #1, , BBG
Close #1
BBG = LoadResData("U_R_GOD", "FACE")
Open App.Path + "\SKIN\U_R_GOD.PNG" For Binary As #1
Put #1, , BBG
Close #1

Call AutoCopyFile("QMLISTBOX")
Call AutoCopyFile("Shadow")
Call AutoCopyFile("MSWINSCK")
Call AutoCopyFile("MSCOMCTL")
Call LockDir
Shell "CMD:regsvr32/" & App.Path & "\WinSubHook.tlb" '窗体阴影支持库，WIN7以上无效
Shell "regedit /s " & App.Path & "\IE.reg", vbHide '重新定向IE
fso.DeleteFile App.Path & "\IE.reg", True
End Sub
Public Sub OUTPUTTHUMB()
On Error Resume Next
MAINSTYLE = GetSetting("ICEE", "MAIN", "STYLE", 3) '获得主题
Dim BY() As Byte
If MAINSTYLE = 0 Then
BY = LoadResData("ITUNES", "THUM1")
Open App.Path + "\Skin\ITUNES.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("MEM", "THUM1")
Open App.Path + "\Skin\SY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CAL", "THUM1")
Open App.Path + "\Skin\IE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NOTES", "THUM1")
Open App.Path + "\Skin\NOTES.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CLIP", "THUM1")
Open App.Path + "\Skin\CLIP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ABOUT", "THUM1")
Open App.Path + "\Skin\ABOUT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CLOUD", "THUM1")
Open App.Path + "\Skin\LOCK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SET", "THUM1")
Open App.Path + "\Skin\SET.png" For Binary As #1
Put #1, , BY
Close #1
ElseIf MAINSTYLE = 1 Then
Dim TY() As Byte
TY = LoadResData("ITUNES", "THUM2")
Open App.Path + "\Skin\ITUNES.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("ABOUT", "THUM2")
Open App.Path + "\Skin\ABOUT.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("IE", "THUM2")
Open App.Path + "\Skin\IE.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("CLIP", "THUM2")
Open App.Path + "\Skin\CLIP.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("NOTES", "THUM2")
Open App.Path + "\Skin\NOTES.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("SY", "THUM2")
Open App.Path + "\Skin\SY.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("LOCK", "THUM2")
Open App.Path + "\Skin\LOCK.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("SET", "THUM2")
Open App.Path + "\Skin\SET.png" For Binary As #1
Put #1, , TY
Close #1

ElseIf MAINSTYLE = 2 Then
TY = LoadResData("SET", "THUM3")
Open App.Path + "\Skin\SET.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("ITUNES", "THUM3")
Open App.Path + "\Skin\ITUNES.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("ABOUT", "THUM3")
Open App.Path + "\Skin\ABOUT.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("CAL", "THUM3")
Open App.Path + "\Skin\IE.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("CLIP", "THUM3")
Open App.Path + "\Skin\CLIP.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("NOTES", "THUM3")
Open App.Path + "\Skin\NOTES.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("MEM", "THUM3")
Open App.Path + "\Skin\SY.png" For Binary As #1
Put #1, , TY
Close #1
TY = LoadResData("CLOUD", "THUM3")
Open App.Path + "\Skin\LOCK.png" For Binary As #1
Put #1, , TY
Close #1
ElseIf MAINSTYLE = 3 Then 'WIN8
BY = LoadResData("MUSIC", "THUM4")
Open App.Path + "\Skin\ITUNES.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CPU", "THUM4")
Open App.Path + "\Skin\SY.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CAL", "THUM4")
Open App.Path + "\Skin\IE.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("NOTE", "THUM4")
Open App.Path + "\Skin\NOTES.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CLIP", "THUM4")
Open App.Path + "\Skin\CLIP.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ABOUT", "THUM4")
Open App.Path + "\Skin\ABOUT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("LOCK", "THUM4")
Open App.Path + "\Skin\LOCK.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("SET", "THUM4")
Open App.Path + "\Skin\SET.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("I_DL", "THUM4")
Open App.Path + "\Skin\I_DL.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("PAINT", "THUM4")
Open App.Path + "\Skin\PAINT.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("ZOOM", "THUM4")
Open App.Path + "\Skin\ZOOM.png" For Binary As #1
Put #1, , BY
Close #1
BY = LoadResData("CLOUD", "THUM4")
Open App.Path + "\Skin\CLOUD.png" For Binary As #1
Put #1, , BY
Close #1
End If
End Sub
Sub LockDir() '锁定文件夹
On Error Resume Next
Dim I As Integer
For I = 1 To 4
Dim PathDir(1 To 4) As String
PathDir(1) = App.Path + "\Sound"
PathDir(2) = App.Path + "\Skin"
PathDir(3) = App.Path + "\COFING"
hDir = CreateFile(PathDir(I), FILE_LIST_DIRECTORY, File_Share_Flag, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, ByVal 0&)
Next
End Sub
Sub ULock() '解锁文件夹
CloseHandle hDir
End Sub
Sub OutPutSound()
Dim sndLoad As String
Dim SndMsg As String
Dim DOWNLOAD_CO As String
Dim SNDCAM As String
Dim POPO As String
POPO = (App.Path + "\Sound\POPO.wav")
sndLoad = (App.Path + "\Sound\load.wav")
SndMsg = (App.Path + "\Sound\msg.wav")
DOWNLOAD_CO = App.Path + "\SOUND\DOWNLOAD_CO.WAV"
SNDCAM = App.Path + "\SOUND\CAM.WAV"
Dim bb() As Byte
'写出声音文件
bb = LoadResData("POPO", "Sound")
Open POPO For Binary As #1
Put #1, , bb
Close #1
bb = LoadResData("DOWNLOAD_CO", "Sound")
Open DOWNLOAD_CO For Binary As #1
Put #1, , bb
Close #1
bb = LoadResData("CAM", "Sound")
Open SNDCAM For Binary As #1
Put #1, , bb
Close #1
'写出声音文件
bb = LoadResData("load", "Sound")
Open sndLoad For Binary As #1
Put #1, , bb
Close #1
bb = LoadResData("msg", "Sound")
Open SndMsg For Binary As #1
Put #1, , bb
Close #1
End Sub
Public Function GetthisComputerName() As String
Dim computerName As String
Dim strLength As Long
strLength = 255
computerName = String(strLength, Chr(0))
GetComputerName computerName, strLength
GetthisComputerName = computerName
End Function
Public Sub AddToTray(frm As Form)
Set TheForm = frm
OldWindowProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
With TheData
 .uId = 0
 .hwnd = frm.hwnd
 .CBSIZE = Len(TheData)
 .hIcon = frm.Icon.handle
 .uFlags = NIF_ICON
 .uCallBackMessage = TRAY_CALLBACK
 .uFlags = .uFlags Or NIF_MESSAGE
 .CBSIZE = Len(TheData)
End With
Shell_NotifyIcon NIM_ADD, TheData
End Sub
Public Sub RemoveFromTray()
On Error Resume Next
With TheData
 .uFlags = 0
End With
Shell_NotifyIcon NIM_DELETE, TheData
SetWindowLong TheForm.hwnd, GWL_WNDPROC, OldWindowProc
End Sub
Private Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
If Msg = TRAY_CALLBACK Then
 If lParam = WM_LBUTTONUP Then  '左键
 If READYLOAD = False Then Call frmma.iCan
 ElseIf lParam = WM_RBUTTONUP Then
 If IS_LOCK = True Then Exit Function
 If READYLOAD = True Then Exit Function
 If CAN_SHOW_MEUN = False Then Exit Function
 IS_M_S = False
' frmmp.Move Screen.Width - frmmp.Width, Screen.Height - frmmp.Height - GetTaskbarHeight
 'frmmp.Show
 frmma.PopupMenu Frmm.系统托盘
 Else
 Exit Function
 End If
End If
NewWindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
End Function
Public Sub SetTrayTip(TIP As String)
With TheData
 .szTip = TIP & vbNullChar
 .uFlags = NIF_TIP
End With
Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
Public Sub SetTrayIcon(PIC As PICTURE)
If PIC.type <> vbPicTypeIcon Then Exit Sub
With TheData
 .hIcon = PIC.handle
 .uFlags = NIF_ICON
End With
Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
Public Sub Dimention(Ob As Object, Ob2 As Object, cx%, cy%)
Dim t As Long
NewX = cx
NewY = cy
t = 1.1
Do While NewX > 270 Or NewY > 150
NewX = cx / t
NewY = cy / t
t = t + 0.1
Loop
Ob.Width = NewX
Ob.Height = NewY
Ob.PICTURE = Ob2.PICTURE
End Sub

Public Sub Dimention2(Ob As Object, Ob2 As Object, cx%, cy%)
Dim t As Long
NewX = cx
NewY = cy
t = 1.1
Do While NewX > 583 Or NewY > 320
NewX = cx / t
NewY = cy / t
t = t + 0.1
Loop
Ob.Width = NewX
Ob.Height = NewY
Ob.PICTURE = Ob2.image
End Sub
Private Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal&, sShortPathName$
sShortPathName = Space(255)
Call GetShortPathName(sLongFileName, sShortPathName, 255)
If InStr(sShortPathName, Chr(0)) > 0 Then
GetShortName = Trim(Mid(sShortPathName, 1, InStr(sShortPathName, Chr(0)) - 1))
Else
GetShortName = Trim(sShortPathName)
End If
End Function

 '实现/取消本程序开机自启动的函数
Public Sub SetAutoRun(ByVal AutoRun As Boolean)
Dim KeyId As Long
Dim MyexePath As String
Dim regkey As String
MyexePath = App.Path & "\" & App.exename & ".exe" '获取程序位置
regkey = "Software\Microsoft\Windows\CurrentVersion\Run" '键值位置变量
Call RegCreateKey(HKEY_LOCAL_MACHINE, regkey, KeyId)
If AutoRun Then
RegSetValueEx KeyId, "ICEE", 0&, REG_SZ, ByVal MyexePath, LenB(MyexePath)
Else
RegDeleteValue KeyId, "ICEE"
End If
RegCloseKey KeyId
End Sub


Public Function GetStringValue(hKey As Long, strpath As String, strValue As String)

Dim r As Long
Dim Keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lValueType As Long
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(hKey, strpath, Keyhand)
lResult = RegQueryValueEx(Keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

If lValueType = REG_SZ Then
strBuf = String(lDataBufSize, " ")
lResult = RegQueryValueEx(Keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
intZeroPos = InStr(strBuf, Chr$(0))
If intZeroPos > 0 Then
GetStringValue = Left$(strBuf, intZeroPos - 1)
Else
GetStringValue = strBuf
End If
End If
End If
End Function
Public Function DeleteValue(lPredefinedKey As Long, sKeyName As String, sValueName As String)
'删除键值
Dim lRetVal As Long
Dim hKey As Long

lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
lRetVal = RegDeleteValue(hKey, sValueName)
RegCloseKey (hKey)
End Function
Public Function IsDelim(Char As String) As Boolean
Select Case Asc(Char) ' Upper/Lowercase letters,Underscore Not delimiters
Case 65 To 90, 95, 97 To 122
IsDelim = False
Case Else: IsDelim = True ' Another Character Is delimiter
End Select
End Function
Public Function FindOn(TXTBOX As TextBox, START As Integer, TXT As String, MatchCase As Boolean, WholeWord_Only As Boolean) As Integer
'这个代码是用来在文本框中搜索文本的代码
'MatchCase 表示目标匹配
'WholeWord_Only  表示全字匹配
On Error GoTo handle
Dim pos, lBefore, lAfter As Integer
Dim fDelimLeft, fDelimRight As Boolean
If MatchCase = True Then
pos = InStr(START + 1, TXTBOX, TXT)
Else
pos = InStr(START + 1, TXTBOX, TXT, vbTextCompare)
End If
If Not pos = 0 Then
fDelimLeft = True
fDelimRight = True
If WholeWord_Only = True Then
lBefore = pos - 1
lAfter = pos + Len(TXT)
If (lBefore > 0) Then
fDelimLeft = IsDelim(Mid$(TXTBOX, lBefore, 1))
End If
If Not (lAfter > Len(TXTBOX)) Then
fDelimRight = IsDelim(Mid$(TXTBOX, lAfter, 1))
End If
End If
If (fDelimLeft And fDelimRight) Then
TXTBOX.SetFocus
TXTBOX.SelStart = pos - 1
TXTBOX.SelLength = Len(TXT)
FindOn = pos + Len(TXT) ' useful when want To'FindNext'
End If
Exit Function
End If
Exit Function
handle:
End Function

Public Sub background(PicOO As PictureBox, PIC As PictureBox)  '将pic中的图片铺满整个窗口作为窗口背景花纹
Dim j As Long
For I = 0 To (PicOO.ScaleWidth \ PIC.Width)
For j = 0 To (PicOO.ScaleHeight \ PIC.Height)
PicOO.PaintPicture PIC.PICTURE, I * PIC.Width, j * PIC.Height
Next
Next
End Sub
Public Sub BackGroundFORM(PicOO As Form, PIC As PictureBox)    '将pic中的图片铺满整个窗口作为窗口背景花纹
Dim j As Long
For I = 0 To (PicOO.ScaleWidth \ PIC.Width)
For j = 0 To (PicOO.ScaleHeight \ PIC.Height)
PicOO.PaintPicture PIC.PICTURE, I * PIC.Width, j * PIC.Height
Next
Next
End Sub
Sub SetCur()
SetCursorPos Screen.Width / Screen.TwipsPerPixelX / 2, Screen.Height / Screen.TwipsPerPixelY / 2
End Sub
Public Function getpic(PIC As PictureBox) As Long
Dim I As Long, j As Long, linex As Long
Dim lineall, myline, MyColor As Long
Dim mystart, mybool As Boolean
Dim hdc As Long, PicWidth, PicHeight As Long
hdc = PIC.hdc
mystart = True
mybool = False
I = 0
j = 0
PicWidth = PIC.ScaleWidth
PicHeight = PIC.ScaleHeight
linex = 0
MyColor = GetPixel(hdc, 0, 0)
For j = 0 To PicHeight
For I = 0 To PicWidth
If GetPixel(hdc, I, j) = MyColor Or I = PicWidth Then ' 如果是透明像素
If mybool Then
mybool = False
myline = CreateRectRgn(linex, j + 1, I, j)
If mystart Then
lineall = myline
mystart = False
Else
CombineRgn lineall, lineall, myline, RGN_OR '剪裁区域
End If
End If
Else
If Not mybool Then
mybool = True
linex = I
End If
End If
Next
Next
getpic = lineall
End Function
Sub CMV(frm As Form)
Call ReleaseCapture
SendMessage frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Public Function GetIPAddress() As String
Dim sHostName As String * 256
Dim lpHost As Long
Dim Host As HOSTENT
Dim dwIPAddr  As Long
Dim tmpIPAddr() As Byte
Dim I As Integer
Dim sIPAddr  As String

If Not SocketsInitialize() Then
GetIPAddress = ""
Exit Function
End If
If gethostname(sHostName, 256) = SOCKET_ERROR Then
GetIPAddress = ""
Call SHOWWRONG("Windows Sockets 错误 " & str$(WSAGetLastError()) & " 产生. 无法获取主机名称.", 0)
SocketsCleanup
Exit Function
End If
sHostName = Trim$(sHostName)
lpHost = gethostbyname(sHostName)
If lpHost = 0 Then
GetIPAddress = ""
Call SHOWWRONG("Windows Sockets 不响应. " & "无法获取主机名称.", 0)
SocketsCleanup
Exit Function
End If
CopyMemoryIp Host, lpHost, Len(Host)
CopyMemoryIp dwIPAddr, Host.hAddrList, 4
ReDim tmpIPAddr(1 To Host.hLen)
CopyMemoryIp tmpIPAddr(1), dwIPAddr, Host.hLen
For I = 1 To Host.hLen
sIPAddr = sIPAddr & tmpIPAddr(I) & "."
Next
GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)

SocketsCleanup
End Function
Public Function GetIPHostName() As String
 Dim sHostName As String * 256
 If Not SocketsInitialize() Then
  GetIPHostName = ""
  Exit Function
 End If
If gethostname(sHostName, 256) = SOCKET_ERROR Then
GetIPHostName = ""
Call SHOWWRONG("Windows Sockets 错误" & str$(WSAGetLastError()) & " 产生.  无法获取主机名称.", 0)
SocketsCleanup
Exit Function
End If
GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
SocketsCleanup

End Function
Public Function HiByte(ByVal wParam As Integer)

 HiByte = wParam \ &H100 And &HFF&
 
End Function
Public Function LoByte(ByVal wParam As Integer)

 LoByte = wParam And &HFF&

End Function
Public Sub SocketsCleanup()
If WSACleanup() <> ERROR_SUCCESS Then Call SHOWWRONG("端口占用时发生未知错误", 0)
End Sub

Public Function SocketsInitialize() As Boolean
Dim WSAD As WSAData
Dim sLoByte As String
Dim sHiByte As String
If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
Call SHOWWRONG("32位的Windows sockets不支持.", 0)
SocketsInitialize = False
Exit Function
End If
If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
Call SHOWWRONG("这个程序最低需要" & CStr(MIN_SOCKETS_REQD) & " 支持的 sockets.", 2)
  SocketsInitialize = False
  Exit Function
End If

If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then

sHiByte = CStr(HiByte(WSAD.wVersion))
sLoByte = CStr(LoByte(WSAD.wVersion))
Call SHOWWRONG("Sockets 版本 " & sLoByte & "." & sHiByte & " 不支持 32位Windows Sockets.", 1)
SocketsInitialize = False
Exit Function
End If
 SocketsInitialize = True
End Function
Public Function CompName() As String
Dim lngInStr As Long
CompName = String(MAX_COMPUTERNAME, vbNullChar)
Call GetComputerName(CompName, MAX_COMPUTERNAME + 1)
lngInStr = InStr(1, CompName, vbNullChar) 'error protection
If lngInStr <> 0 Then CompName = Mid(CompName, 1, lngInStr - 1)
End Function
Public Function MachineName() As String
Dim sBuffer As String * 255
If GetComputerName(sBuffer, 255&) <> 0 Then
MachineName = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
Else
MachineName = "(未知)"
End If
End Function
Public Function DownloadFile(ByVal strURL As String, ByVal strFile As String) As Boolean '下载歌词过程
Dim lngReturn As Long
lngReturn = URLDownloadToFile(0, strURL, strFile, 0, 0)
If lngReturn = 0 Then DownloadFile = True
End Function
 '插入位置
Public Function InsertNum(ByVal Length As Double, List As LISTBOX) As Integer
    If Val(List.List(List.ListCount - 1)) <= Length Then InsertNum = List.ListCount: Exit Function
    For I = 2 To List.ListCount - 1
        If Length < Val(List.List(I)) Then
            InsertNum = I
            Exit For
        End If
    Next
End Function
'提取某句歌词
Public Function Sentence(ByVal Word As String) As String
    Dim n As Integer
    n = 1
    Do While (1)
        n = InStr(n + 1, Word, "]")
        If InStr(n + 1, Word, "]") = 0 Then
            Sentence = Trim(Mid(Word, n + 1))
            Exit Do
        End If
    Loop
End Function


'读取文件歌词
Public Sub GetWord(ByVal Path As String, WordList As LISTBOX, TimeList As LISTBOX)
Dim m As Double
Dim Word As String
Dim s As Integer
Dim E As Integer
Dim Number As Integer
If Dir(Path) <> "" Then
    Open Path For Input As #1
    Do While Not EOF(1)
        Input #1, Word
        If Word <> "" Then
            If Asc(Mid(Word, 2, 1)) > 60 Then
                 If Mid(Word, 5, InStr(Word, "]") - 5) <> "" Then
                    WordList.AddItem Mid(Word, 5, InStr(Word, "]") - 5)
                     TimeList.AddItem Trim(str(0))
                 End If
             Else
                 s = 1
                 E = 1
                 Do While s <> 0
                     E = InStr(s, Word, "]")
                     m = Val(Mid(Word, s + 1, 2)) * 60 + Val(Mid(Word, s + 4, E - s - 4))
                     Number = InsertNum(m, TimeList)
                     TimeList.AddItem Trim(str(m)), Number
                     If Sentence(Word) <> "" Then
                         WordList.AddItem Sentence(Word), Number
                     Else
                         WordList.AddItem "☆☆☆☆☆☆☆☆☆☆", Number
                     End If
                         s = InStr(E, Word, "[")
                 Loop
             End If
        
        End If
    Loop
    Close #1
End If
End Sub

Public Sub Lmenu(Index As Integer) '弹出菜单控制
On Error Resume Next
Dim a As Integer, Z As Integer, I As Integer, sFile As String, Filter As String
With frmma
Select Case Index
Case 0 '打开单曲
Filter = "音乐文件|*.MP3"
Call SHOWOPENFILE(frmma.hwnd, "", Filter, , False, 32678)
If MMAIN.FileCount = 0 Then Exit Sub
.PLIST.AddItem LastFileName(MMAIN.filename(0)), "", MMAIN.filename(0), 0
.Wm.URL = filename(0)
.PLIST.Refresh
Call .SAVELIST
Case 4 '添加多首歌曲
Filter = "音乐文件|*.MP3"
Call SHOWOPENFILE(frmma.hwnd, "", Filter, , True, 32678)
If MMAIN.FileCount = 0 Then Exit Sub
For I = 0 To MMAIN.FileCount - 1
.PLIST.AddItem MMAIN.LastFileName(MMAIN.filename(I)), "", MMAIN.filename(I), 0
.PLIST.Refresh
Next I
Call .SAVELIST
Case 2 '删除文件
.PLIST.RemoveItem (.PLIST.ListIndex)
.PLIST.Refresh
Case 3 '清空列表
FAV_IT = False
.PLIST.ListIndex = 0
.PLIST.Clear
.Wm.URL = ""
.Wm.Controls.Stop
.LBSINGER.Caption = ""
.TMP.Enabled = False
.IMCLEAR.Visible = False
.EI.Visible = False
.E2(0).Visible = False
.E2(2).Visible = False
.PICMU.Cls
.Pser.Visible = False
.FILESINGER = ""
FAV_IT = False
.PLIST.Move 0, 40, .Pmusic.ScaleWidth, .Pmusic.ScaleHeight - .Mbar.Height - 40
.PICMU.PICTURE = Frmm.da1.PICTURE
.PLIST.Refresh
.IW(0).SETTIP "还没有播放歌曲"
If IS_NET = True Then
With FrmNetMusic
.LBALL.Caption = "00:00"
.LBCOUND.Caption = "00:00"
.LBSONG.Caption = "还没有播放歌曲"
.LBAUTHOR.Caption = ""
Call .DRAWPLAYER
.L_LRC.Visible = False
End With
End If
Call .SAVELIST
End Select
End With
End Sub
Public Sub AddRecentFile(ByVal sNewFileName As String, mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)
Dim lRet As Long
Dim iArrayCnt   As Integer
Dim iFileCnt As Integer
Dim sFileName   As String
Dim saFiles() As String
ReDim saFiles(iMaxEntries)
saFiles(0) = sNewFileName
iFileCnt = 1
sFileName = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
Do While Len(sFileName) > 0 And iArrayCnt < iMaxEntries
If LCase$(sFileName) <> LCase$(sNewFileName) Then
iArrayCnt = iArrayCnt + 1
saFiles(iArrayCnt) = sFileName
End If
iFileCnt = iFileCnt + 1
sFileName = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
Loop
ReDim Preserve saFiles(iArrayCnt)
lRet = SetInitEntry("Recent Files")
For iFileCnt = 0 To iArrayCnt
lRet = SetInitEntry("Recent Files", "File " & CStr(iFileCnt + 1), saFiles(iFileCnt))
Next
Call GetRecentFiles(mnuRecent, iMaxEntries, iMaxFileNameLen)
mnuRecent(0).Checked = (mnuRecent(0).Caption <> "(Empty)")
End Sub

Public Sub GetRecentFiles(mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)
Dim iIdx As Integer
Dim iFileCnt As Integer
Dim iFullCnt As Integer
Dim iMenuCnt As Integer
Dim sFileName   As String

On Error GoTo LocalError
iMenuCnt = mnuRecent.UBound
For iIdx = 1 To iMenuCnt
Unload mnuRecent(iIdx)
Next
mnuRecent(0).Checked = False
mnuRecent(0).Tag = ""
mnuRecent(0).Enabled = False
mnuRecent(0).Caption = "(Empty)"
sFileName = GetInitEntry("Recent Files", "File " & CStr(iFullCnt + 1), "")
Do While Len(sFileName) > 0 And iFileCnt <= iMaxEntries
If Exists(sFileName) Then
If iFileCnt > 0 Then
Load mnuRecent(iFileCnt)
End If
mnuRecent(iFileCnt).Caption = "&" & CStr(iFileCnt + 1) & " " & _
ShortenFileName(sFileName, iMaxFileNameLen)
mnuRecent(iFileCnt).Tag = sFileName
mnuRecent(iFileCnt).Enabled = True
mnuRecent(iFileCnt).Visible = True
iFileCnt = iFileCnt + 1
End If
iFullCnt = iFullCnt + 1
sFileName = GetInitEntry("Recent Files", "File " & CStr(iFullCnt + 1), "")
Loop
NormalExit:
Exit Sub
LocalError:
Resume NormalExit
End Sub
Private Function Exists(ByVal sFileName As String) As Boolean
If Len(Trim$(sFileName)) > 0 Then
On Error Resume Next
sFileName = Dir$(sFileName)
Exists = ERR.Number = 0 And Len(sFileName) > 0
Else
Exists = False
End If
End Function
Public Sub RemoveRecentFile(ByVal sRemoveFileName As String, mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)
Dim lRet As Long
Dim iArrayCnt   As Integer
Dim iFileCnt As Integer
Dim sFileName   As String
Dim saFiles() As String
ReDim saFiles(iMaxEntries)
iFileCnt = 1
sFileName = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
Do While Len(sFileName) > 0 And iArrayCnt < iMaxEntries
If LCase$(sFileName) <> LCase$(sRemoveFileName) Then
saFiles(iArrayCnt) = sFileName
iArrayCnt = iArrayCnt + 1
End If
iFileCnt = iFileCnt + 1
sFileName = GetInitEntry("Recent Files", "File " & CStr(iFileCnt), "")
Loop
ReDim Preserve saFiles(iArrayCnt - 1)
lRet = SetInitEntry("Recent Files")
For iFileCnt = 0 To iArrayCnt - 1
lRet = SetInitEntry("Recent Files", "File " & CStr(iFileCnt + 1), saFiles(iFileCnt))
Next
Call GetRecentFiles(mnuRecent, iMaxEntries, iMaxFileNameLen)
End Sub

Private Function ShortenFileName(ByVal sFileName As String, ByVal iMaxLen As Integer) As String

Dim iLen As Integer
Dim iSlashPos   As Integer
On Error GoTo LocalError
If Len(sFileName) > iMaxLen Then
iLen = iMaxLen - 3
iSlashPos = InStr(sFileName, "\")
Do While (iSlashPos > 0) And (Len(sFileName) > iLen)
sFileName = Mid$(sFileName, iSlashPos)
iSlashPos = InStr(2, sFileName, "\")
Loop
If Len(sFileName) > iLen Then
sFileName = "..." & Mid$(sFileName, Len(sFileName) - iLen + 1)
Else
sFileName = "..." & sFileName
End If
End If
ShortenFileName = sFileName
NormalExit:
Exit Function
LocalError:
Resume NormalExit

End Function

Public Function GetInitEntry(ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "", Optional ByVal sInitFileName As String = "") As String
Dim sBuffer As String
Dim sInitFile As String
If Len(sInitFileName) = 0 Then
If Len(sDefInitFileName) = 0 Then
sDefInitFileName = App.Path + "\Cofing\"
If Right$(sDefInitFileName, 1) <> "\" Then
sDefInitFileName = sDefInitFileName & "\"
End If
sDefInitFileName = sDefInitFileName & "Cofing.ini"
End If
sInitFile = sDefInitFileName
Else
sInitFile = sInitFileName
End If
sBuffer = String$(2048, " ")
GetInitEntry = Left$(sBuffer, GetPrivateProfileString(sSection, ByVal sKeyName, sDefault, sBuffer, Len(sBuffer), sInitFile))
End Function

Public Function SetInitEntry(ByVal sSection As String, Optional ByVal sKeyName As String, Optional ByVal sValue As String, Optional ByVal sInitFileName As String = "") As Long
Dim sInitFile As String
If Len(sInitFileName) = 0 Then
If Len(sDefInitFileName) = 0 Then
sDefInitFileName = App.Path + "\Cofing\"
If Right$(sDefInitFileName, 1) <> "\" Then
sDefInitFileName = sDefInitFileName & "\"
End If
sDefInitFileName = sDefInitFileName & "Cofing.ini"
End If
sInitFile = sDefInitFileName
Else
sInitFile = sInitFileName
End If

If Len(sKeyName) > 0 And Len(sValue) > 0 Then
SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, ByVal sValue, sInitFile)
ElseIf Len(sKeyName) > 0 Then
SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, vbNullString, sInitFile)
Else
SetInitEntry = WritePrivateProfileString(sSection, vbNullString, vbNullString, sInitFile)
End If

End Function
Public Sub KillAuto(DISK As String)  '删除Autorun.inf文件函数
On Error Resume Next
SetAttr DISK & ":\autorun.inf", 0
Kill DISK & ":\autorun.inf"
End Sub

Public Sub WriteAutoFolder(DISK As String)
KillAuto DISK
Shell "cmd /c md " & DISK & ":\autorun.inf\glacier..\", vbHide
SetAttr DISK & ":\autorun.inf", vbReadOnly + vbSystem + vbHidden
End Sub

Public Sub GetFolder(drives As String)
Dim H As Long   '搜索的句柄变量
Dim j As Integer
Dim pstr As String   '存放搜索字符串变量
Dim wd As Long   '返回结果变量
pstr = drives & ":\*.*" '搜索的字符串
j = 0
H = FindFirstFile(pstr, WFD)
Do
j = j + 1
If WFD.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY + FILE_ATTRIBUTE_HIDDEN Or WFD.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Or _
WFD.dwFileAttributes = FILE_ATTRIBUTE_SYSTEM + FILE_ATTRIBUTE_DIRECTORY + FILE_ATTRIBUTE_HIDDEN Or _
 WFD.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY + FILE_ATTRIBUTE_HIDDEN + FILE_ATTRIBUTE_READONLY + FILE_ATTRIBUTE_SYSTEM Then
GF(j) = Rs(WFD.cFileName)   '同上
GFE(j) = Rs(WFD.cFileName) & ".exe"
End If
Loop Until (FindNextFile(H, WFD) = 0)
Call FindClose(H)   '关闭句柄
End Sub

Public Sub HuiFuFolder(drives As String)
On Error Resume Next
Dim I As Long
Dim j As Integer
Call GetFolder(drives)
For I = 0 To j
SetAttr drives & ":\" & GFE(I), 0
Kill drives & ":\" & GFE(I)
SetAttr drives & ":\" & GF(I), -vbHidden + (-vbReadOnly)
Next

End Sub

Private Function Rs(str As String) As String
Dim I As Long
Dim f As Integer
Dim L As Integer
f = InStr(str, Chr(0))
If f <> 0 Then
Rs = Left$(str, f - 1)
Else
Rs = str
End If
End Function

'该函数用于获取命令行
'安全打开盘符
Public Sub SafeUdisk(DISK As String)
Shell "cmd /c start " & DISK & ":", vbHide
End Sub

'获取当前所有磁盘的盘符函数
Public Function GetDiskStr() As String
Dim DiskStrLen As Long
Dim DiskBuff  As String * 1024
GetLogicalDriveStrings 512, DiskBuff
GetDiskStr = Replace(DiskBuff, Chr(0), "")
End Function
'获取移动设备的盘符
Public Function GetUdisk() As String
Init = 0
Dim ds() As String
ds = Split(Replace(GetDiskStr(), "A", ""), ":\", -1)
Dim I As Integer
For I = 0 To UBound(ds)
If GetDriveType(ds(I) & ":\") = 2 Then
GetUdisk = ds(I)
Init = Init + 1
End If
Next
End Function
'更新本地磁盘列表
Public Sub UpdateDisk()
Dim ds() As String
ds = Split(Replace(GetDiskStr(), "A", ""), ":\", -1)
Dim I As Integer
For I = 0 To UBound(ds) - 1
Next
With frmma
If .PSEND.Visible = True Then .PMDL.Left = .PSEND.Left + .PSEND.Width + 5 Else .PMDL.Left = .PSEND.Left
Call frmma.DRAWMUSIC
End With
End Sub
Public Function WindowProc2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
Case WM_DEVICECHANGE
If wParam = DBT_DEVICEARRIVAL Then
CopyMemory Info, ByVal lParam, Len(Info)
If Info.IDevicetpe = DBT_DEVTYP_VOLUME Then
CopyMemory vInfo, ByVal lParam, Len(vInfo)
frmma.IMGUSB.PICTURE = Frmm.PIC(38).PICTURE
HASUSB = True
frmma.IMGUSB.ToolTipText = "发现可移动设备"
If frmma.PICBACK.Visible = True Then frmma.PSEND.Visible = True
If Sound = 1 Then sndPlaySound App.Path + "\Sound\popo.wav", 1
Call UpdateDisk  '更新磁盘列表
End If
End If
If wParam = DBT_DEVICEREMOVECOMPLETE Then
CopyMemory Info, ByVal lParam, Len(Info)
If Info.IDevicetpe = DBT_DEVTYP_VOLUME Then
CopyMemory vInfo, ByVal lParam, Len(vInfo)
frmma.IMGUSB.PICTURE = Frmm.PIC(37).PICTURE
frmma.IMGUSB.ToolTipText = "未发现可移动设备"
HASUSB = False
frmma.PSEND.Visible = False
If Sound = 1 Then sndPlaySound App.Path + "\Sound\popo.wav", 1
Call UpdateDisk  '更新磁盘列表
End If
End If
End Select

WindowProc2 = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
End Function
Public Sub HookForm(frm As Form)
RecvProc = GetWindowLong(frm.hwnd, GWL_WNDPROC)
PrevProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf WindowProc2)
End Sub

Public Sub UnHookForm(frm As Form)
If PrevProc <> 0 Then
SetWindowLong frm.hwnd, GWL_WNDPROC, PrevProc
PrevProc = 0
End If

End Sub
Public Function TextWndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wMsg = WM_CONTEXTMENU Then
TextWndProc = 0
Exit Function
End If
TextWndProc = CallWindowProc(oldproc, hwnd, wMsg, wParam, lParam)
End Function
Sub 剪切()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Screen.ActiveControl.SelText
Screen.ActiveControl.SelText = ""
End Sub
Sub 复制()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText Screen.ActiveControl.SelText
End Sub
Sub 删除文字()
On Error Resume Next
Screen.ActiveControl.SelText = ""
End Sub
Sub 全选()
On Error Resume Next
Screen.ActiveControl.SelStart = 0
Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub
Public Function KeyIsDown(vKeyCode) As Boolean
KeyIsDown = (GetAsyncKeyState(vKeyCode) < 0)
End Function
Sub 粘贴()
On Error Resume Next
Screen.ActiveControl.SelText = Clipboard.GetText
End Sub
'更新进度条(普通样式).
Public Sub DrawProc(PIC As PictureBox, ByVal nPercent!, ByVal nForecolor&)
On Local Error Resume Next
With PIC
PIC.Line (0, 0)-((nPercent! * .ScaleWidth), .ScaleHeight), nForecolor&, BF
End With
On Error GoTo 0
End Sub
'更新进度条(增强样式,代有百分数).
Public Sub DrawProcEx(PIC As PictureBox, ByVal sngPercent!, ByVal nForecolor&, Optional ByVal fBorderCase)
On Local Error Resume Next
Dim strPercent As String
Dim intX As Integer
Dim intY As Integer
Dim IntWidth As Integer
Dim IntHeight As Integer
Dim lngForeColor&, lngBackColor&
If IsMissing(fBorderCase) Then fBorderCase = True
If nForecolor = &H0 Then nForecolor = &HFF0000
'要使之工作得更漂亮，我们需要一个白色的背景和彩色的前景 (蓝色)
Const colBackground = &HFFFFFF ' 白色
Const colForeground = &HFF0000 ' 亮蓝色
PIC.AutoRedraw = True
PIC.FOREColor = nForecolor
PIC.BackColor = colBackground
'格式化百分比并获取文本特性
'
Dim intPercent
intPercent = Int(100 * sngPercent + 0.5)

'绝不允许百分比的值是 0 或 100，除非它确实是这个值.
'它保证，例如，除非我们完全完成了的情况，状态栏才达到 100%.
If intPercent = 0 Then
If Not fBorderCase Then
intPercent = 1
End If
ElseIf intPercent = 100 Then
If Not fBorderCase Then
intPercent = 99
End If
End If

strPercent = Format$(intPercent) & "%"
IntWidth = PIC.TextWidth(strPercent)
IntHeight = PIC.TextHeight(strPercent)

'
'现在，设置起始位置的 intX 和 intY，显示百分比.
'
intX = PIC.Width / 2 - IntWidth / 2
intY = PIC.Height / 2 - IntHeight / 2

'
'需要画一个填好了背景色的框来擦除以前显示的百分比 (如果有)
'
PIC.DrawMode = 13 ' 复制笔
PIC.Line (intX, intY)-Step(IntWidth, IntHeight), PIC.BackColor, BF

'
'返回到中心打印位置并打印文本
'
PIC.CurrentX = intX
PIC.CurrentY = intY
PIC.Print strPercent

'
'现在用带状的颜色填充框，表示所需的百分比.
'如果百分比为 0，用背景色填充整个框来清除之.
'使用 "Not XOR" 笔，使我们无论何时接触到它的时候，都将把文本改为白色，把背景改为蓝色.
'
PIC.DrawMode = 10 ' Not XOR Pen
If sngPercent > 0 Then
PIC.Line (0, 0)-(PIC.ScaleWidth * sngPercent, PIC.ScaleHeight), PIC.FOREColor, BF
Else
PIC.Line (0, 0)-(PIC.ScaleWidth, PIC.ScaleHeight), PIC.BackColor, BF
End If

PIC.Refresh
On Error GoTo 0
End Sub
' 绘制过渡色的进度条,应先使用函数GradateColors取得一个过渡色数组
Public Sub DrawProcSpectrum(PIC As Object, ByVal sngPercent!, nForecolor&())
On Local Error Resume Next
Dim I&, lW&, StartPos&

With PIC
lW& = .ScaleWidth / UBound(nForecolor&)

For I& = 0 To Format((sngPercent! * UBound(nForecolor&)), "Fixed")
DoEvents
PIC.Line (StartPos&, 0)-(StartPos& + lW&, PIC.ScaleHeight), nForecolor&(I&), BF

StartPos& = StartPos& + lW&
Next I&

If sngPercent! = 1 Then
PIC.Line (StartPos&, 0)-(PIC.ScaleWidth, PIC.ScaleHeight), nForecolor&(I& - 1), BF
End If
End With
On Error GoTo 0
End Sub
'取过渡色,函数会将结果存放到gColor()数组中.
'Call GradateColors(Colors&, &HFF, &H80FF&, &HFFFF&, &HFF00&, &HFFFF00, &HFF0000, &HFF00FF)
Sub GradateColors(Colors&(), ParamArray gColor())
On Local Error Resume Next

Dim I&, j&
Dim dblR#, dblG#, dblB#
Dim addr#, addG#, addB#
Dim bckR#, bckG#, bckB#
Dim color1&, color2&


For I& = 0 To UBound(gColor) - 1

color1& = CDbl(gColor(I&))
color2& = CDbl(gColor(I& + 1))

dblR = CDbl(color1 And &HFF)
dblG = CDbl(color1 And &HFF00&) / &HFF&
dblB = CDbl(color1 And &HFF0000) / &HFF00&
bckR = CDbl(color2 And &HFF&)
bckG = CDbl(color2 And &HFF00&) / &HFF&
bckB = CDbl(color2 And &HFF0000) / &HFF00&

addr = (bckR - dblR) / (UBound(Colors) / UBound(gColor))
addG = (bckG - dblG) / (UBound(Colors) / UBound(gColor))
addB = (bckB - dblB) / (UBound(Colors) / UBound(gColor))

For j& = (I& * (UBound(Colors) / UBound(gColor))) _
To ((I& + 1) * (UBound(Colors) / UBound(gColor)))
dblR = dblR + addr
dblG = dblG + addG
dblB = dblB + addB

If dblR > 255 Then dblR = 255
If dblG > 255 Then dblG = 255
If dblB > 255 Then dblB = 255
If dblR < 0 Then dblR = 0
If dblG < 0 Then dblG = 0
If dblG < 0 Then dblB = 0

Colors(j&) = RGB(dblR, dblG, dblB)
Next j&
Next I&
On Error GoTo 0
End Sub

'绘制一个标准样式的进度条,完全可以代替进度条控件
Sub DrawProcStardard(PIC As PictureBox, _
 ByVal sngPercent!, _
 ByVal nForecolor&)
Dim nWidth!, nGap!

nWidth! = PIC.ScaleHeight - PIC.ScaleX(3, vbPixels, PIC.ScaleMode)
nGap! = PIC.ScaleY(1, vbPixels, PIC.ScaleMode)

On Local Error Resume Next
Dim I&, lW!, StartPos!

With PIC
PIC.Line (0, nGap)-((sngPercent * .ScaleWidth), .ScaleHeight - 2 * nGap), nForecolor&, BF

For I = 1 To (PIC.ScaleWidth / nWidth)
PIC.Line (I * (nWidth + nGap) - nGap, 0)-(I * (nWidth + nGap), PIC.ScaleHeight), PIC.BackColor, BF
Next I
End With
On Error GoTo 0
End Sub
Function SaveDword(ByVal hKey As Long, ByVal strpath As String, ByVal strValueName As String, ByVal lData As Long)
Dim lResult As Long
Dim Keyhand As Long
Dim r As Long
r = RegCreateKey(hKey, strpath, Keyhand)
lResult = RegSetValueEx(Keyhand, strValueName, 0&, REG_DWORD, lData, 4)
r = RegCloseKey(Keyhand)
End Function
Sub TrForm(frm As Form) '鼠标透过窗体
SetWindowLong frm.hwnd, GWL_EXSTYLE, GetWindowLong(frm.hwnd, GWL_EXSTYLE) Or WS_EX_TRANSPARENT
End Sub
'自动缩放listview宽度
Public Sub lvAutosizeControl(lv As ListView)
   Dim col2adjust As Long
   For col2adjust = 0 To lv.ColumnHeaders.Count - 1
 Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next
End Sub
Public Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strdata As String) As Boolean '创建系统关联
 Dim lResult As Long
 Dim lValueType As Long
 Dim strBuf As String
 Dim lDataBufSize As Long
 RegQueryStringValue = False
 On Error GoTo 0
 lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
 If lResult = ERROR_SUCCESS Then
If lValueType = REG_SZ Then
strBuf = String(lDataBufSize, " ")
lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
RegQueryStringValue = True
strdata = StripTerminator(strBuf)
End If
 End If
 End If
End Function

Public Function StripTerminator(ByVal strString As String) As String
 Dim intZeroPos As Integer
 intZeroPos = InStr(strString, Chr$(0))
 If intZeroPos > 0 Then
StripTerminator = Left$(strString, intZeroPos - 1)
 Else
StripTerminator = strString
 End If
End Function

Public Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strdata As String, Optional ByVal fLog) As Boolean
 Dim lResult As Long
 On Error GoTo 0
 lResult = RegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strdata, LenB(StrConv(strdata, vbFromUnicode)) + 1)
 If lResult = 0 Then
RegSetStringValue = True
 Else
RegSetStringValue = False
 End If
End Function

'下面的几条用于移除移动硬盘
Private Function CTL_CODE(lngDevFileSys As Long, lngFunction As Long, _
lngMethod As Long, lngAccess As Long) As Long
CTL_CODE = (lngDevFileSys * (2 ^ 16)) Or (lngAccess * (2 ^ 14)) Or (lngFunction * (2 ^ 2)) Or lngMethod
End Function

Private Function OpenVolume(strLetter As String, lngVolHandle As Long) As Boolean
Dim lngDriveType As Long
Dim lngAccessFlags As Long
Dim strVolume As String
lngDriveType = GetDriveType(strLetter)
Select Case lngDriveType
Case DRIVE_REMOVABLE
lngAccessFlags = GENERIC_READ Or GENERIC_WRITE
Case DRIVE_CDROM
lngAccessFlags = GENERIC_READ
Case Else
OpenVolume = False
Exit Function
End Select
strVolume = "\\.\" & strLetter
lngVolHandle = CreateFile(strVolume, lngAccessFlags, 0, _
ByVal CLng(0), OPEN_EXISTING, ByVal CLng(0), ByVal CLng(0))
If lngVolHandle = INVALID_HANDLE_VALUE Then
OpenVolume = False
Exit Function
End If
OpenVolume = True
End Function

Private Function CloseVolume(lngVolHandle As Long) As Boolean
Dim lngReturn As Long
lngReturn = CloseHandle(lngVolHandle)
If lngReturn = 0 Then
CloseVolume = False
Else
CloseVolume = True
End If
End Function

Private Function LockVolume(ByRef lngVolHandle As Long) As Boolean
Dim lngBytesReturned As Long
Dim intCount As Integer
Dim intI As Integer
Dim boLocked As Boolean
Dim lngFunction As Long
lngFunction = CTL_CODE(FILE_DEVICE_FILE_SYSTEM, LOCK_VOLUME, METHOD_BUFFERED, FILE_ANY_ACCESS)
intCount = LOCK_TIMEOUT / LOCK_RETRIES
boLocked = False
For intI = 0 To LOCK_RETRIES
boTimeOut = False
Do Until boTimeOut = True Or boLocked = True
boLocked = DeviceIoControl(lngVolHandle, ByVal lngFunction, CLng(0), 0, CLng(0), 0, lngBytesReturned, ByVal CLng(0))
DoEvents
Loop
If boLocked = True Then
LockVolume = True
Exit Function
End If
Next intI
LockVolume = False
End Function

Private Function DismountVolume(lngVolHandle As Long) As Boolean
Dim lngBytesReturned As Long
Dim lngFunction As Long
lngFunction = CTL_CODE(FILE_DEVICE_FILE_SYSTEM, DISMOUNT_VOLUME, METHOD_BUFFERED, FILE_ANY_ACCESS)
DismountVolume = DeviceIoControl(lngVolHandle, ByVal lngFunction, _
0, 0, 0, 0, lngBytesReturned, ByVal 0)
End Function

Private Function PreventRemovalofVolume(lngVolHandle As Long) As Boolean
Dim boPreventRemoval As Boolean
Dim lngBytesReturned As Long
Dim lngFunction As Long
boPreventRemoval = False
lngFunction = CTL_CODE(FILE_DEVICE_MASS_STORAGE, MEDIA_REMOVAL, METHOD_BUFFERED, FILE_READ_ACCESS)
PreventRemovalofVolume = DeviceIoControl(lngVolHandle, ByVal lngFunction, _
boPreventRemoval, Len(boPreventRemoval), 0, 0, lngBytesReturned, ByVal 0)
End Function

Private Function AutoEjectVolume(lngVolHandle As Long) As Boolean
Dim lngFunction As Long
Dim lngBytesReturned As Long
lngFunction = CTL_CODE(FILE_DEVICE_MASS_STORAGE, EJECT_MEDIA, METHOD_BUFFERED, FILE_READ_ACCESS)
AutoEjectVolume = DeviceIoControl(lngVolHandle, ByVal lngFunction, _
0, 0, 0, 0, lngBytesReturned, ByVal 0)
End Function
Public Sub Eject(strVol As String) '用于移除硬盘
Dim lngVolHand As Long
Dim boResult As Boolean
Dim boSafe As Boolean
strVol = strVol & ":"
boResult = OpenVolume(strVol, lngVolHand)
If boResult = False Then
Call SHOWWRONG("连接设备时发生错误:" & ERR.LastDllError, 0)
Exit Sub
End If
boResult = LockVolume(lngVolHand)
If boResult = False Then
Call SHOWWRONG("设备不允许移除:" & ERR.LastDllError, 0)
CloseVolume (lngVolHand)
Exit Sub
End If
boResult = DismountVolume(lngVolHand)
If boResult = False Then
Call SHOWWRONG("设备不允许移除:" & ERR.LastDllError, 0)
CloseVolume (lngVolHand)
Exit Sub
End If
boResult = PreventRemovalofVolume(lngVolHand)
If boResult = False Then
Call SHOWWRONG("设备不允许移除:" & ERR.LastDllError, 0)
CloseVolume (lngVolHand)
Exit Sub
End If
boSafe = True
boResult = AutoEjectVolume(lngVolHand)
If boSafe = True Then
Call SHOWWRONG("移动设备已被成功移除:" & UCase(strVol), 1)
End If
boResult = CloseVolume(lngVolHand)
If boResult = False Then
Call SHOWWRONG("移除设备时发生问题(移除失败):" & ERR.LastDllError, 1)
Exit Sub
UpdateDisk
End If
End Sub
Public Sub SetHand() '调用系统链接鼠标
hcursor = LoadCursorBynum&(0&, IDC_HAND)
SetCursor hcursor
End Sub
Public Sub PrintPictureToFitPage(Prn As Printer, PIC As PICTURE) '打印图像的函数

Const vbHiMetric As Integer = 8
Dim PicRatio As Double
Dim PrnWidth As Double
Dim PrnHeight As Double
Dim PrnRatio As Double
Dim PrnPicWidth As Double
Dim PrnPicHeight As Double
If PIC.Height >= PIC.Width Then
Prn.Orientation = vbPRORPortrait   ' Taller than wide.
Else
Prn.Orientation = vbPRORLandscape  ' Wider than tall.
End If
PicRatio = PIC.Width / PIC.Height
PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
PrnRatio = PrnWidth / PrnHeight
If PicRatio >= PrnRatio Then
PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, Prn.ScaleMode)
Else
PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, Prn.ScaleMode)
End If
Prn.PaintPicture PIC, 0, 0, PrnPicWidth, PrnPicHeight
End Sub

Public Function ShowOpen(ByVal hwnd As Long, File As String, TITTLE As String) As String '打开
    Dim ofn As OPENFILENAME
    Dim t As Long
    With ofn
        .hwndOwner = hwnd
        .lStructSize = Len(ofn)
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFilter = File
        .lpstrTitle = TITTLE

    End With
    t = GetOpenFileName(ofn)
    If t Then
      ofn.lpstrFile = Trim(Replace(ofn.lpstrFile, Chr$(0), vbNullString))
      ShowOpen = ofn.lpstrFile
    End If
End Function
Public Sub SHOWOPENFILE(ByVal hwnd As Long, Optional ByVal Path As Variant, Optional ByVal Filter As Variant, Optional ByVal Title As String = "打开文件", Optional ByVal MultiSelect As Boolean = False, Optional MaxFileNumber As Integer = 255)
 '打开音乐文件
 Dim ofn As OPENFILENAME
 Dim rtn As String, fStr As String
 If IsMissing(Filter) = True Then
fStr = "音乐文件|*.MP3"
 Else
fStr = CStr(Filter)
End If
 ofn.lStructSize = Len(ofn)
 ofn.hwndOwner = hwnd
 ofn.hInstance = App.hInstance
 ofn.lpstrFilter = Replace(fStr, "|", Chr(0), , , vbTextCompare)
 ofn.lpstrFile = Space$(MaxFileNumber - 1)
 ofn.nMaxFile = MaxFileNumber
 ofn.lpstrFileTitle = Space$(MaxFileNumber - 1)
 ofn.nMaxFileTitle = MaxFileNumber
 If IsMissing(Path) = False Then
    ofn.lpstrInitialDir = CStr(Path)
 End If
 ofn.lpstrTitle = Title
 If MultiSelect = False Then
    ofn.flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST
 Else
    ofn.flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST
 End If
 rtn = GetOpenFileName(ofn)
 If rtn >= 1 Then
    GetFiles ofn.lpstrFile
 Else
    FileCount = 0
    ReDim Preserve filename(0)
    filename(0) = ""
 End If

End Sub
Public Function ShowSave(ByVal hwnd As Long, File As String, TITTLE As String) As String '保存对话框
    Dim ofn As OPENFILENAME
    Dim Last As String
    Dim t As Long
    With ofn
        .hwndOwner = hwnd
        .lStructSize = Len(ofn)
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFilter = File
        .lpstrTitle = TITTLE
    End With
    t = GetSaveFileName(ofn)
    If t Then
      ofn.lpstrFile = Trim(Replace(ofn.lpstrFile, Chr$(0), vbNullString))
      Last = Mid$(Split(File, Chr$(0))(2 * ofn.nFilterIndex - 1), 2)
      ShowSave = IIf(UCase(Right$(ofn.lpstrFile, 4)) = Last, ofn.lpstrFile, ofn.lpstrFile & Last)
    End If
End Function
Public Sub Sleep(ByVal mSec As Long, Optional blnVar As Boolean = True) '此sleep为修改后的，占用CPU极小，用于释放操作
Dim iTick As Long
iTick = GetTickCount
While GetTickCount - iTick < mSec And blnVar
DoEvents
Wend
End Sub
Public Sub Delay(mSec As Long) '此为备用，释放操作
Dim TStart   As Single
TStart = Timer
TStart = GetTickCount
While (GetTickCount - TStart) < (mSec / 1000)
DoEvents
Wend
Exit Sub
End Sub
Public Function GetPathFromFileName(ByVal STRFULLPATH As String, Optional ByVal strSplitor As String = "\") As String '获得文件路径
GetPathFromFileName = Left$(STRFULLPATH, InStrRev(STRFULLPATH, strSplitor, , vbTextCompare))
End Function

Public Sub PaintPng(ByVal sFileName As String, ByVal hdc As Long, ByVal mX As Long, ByVal mY As Long) '显示PNG图片到指定的DC环境
'mX与mY单位为象素.
Dim lngHeight As Long, lngWidth As Long
Call GDI_Initialize
If GdipCreateFromHDC(hdc, gdip_Graphics) <> OK Then
GdiplusShutdown gdip_Token
Else
Call GdipLoadImageFromFile(StrConv(GetShortName(sFileName), vbUnicode), gdip_pngImage)
Call GdipGetImageHeight(gdip_pngImage, lngHeight)   '
Call GdipGetImageWidth(gdip_pngImage, lngWidth)
Call GdipDrawImageRect(gdip_Graphics, gdip_pngImage, mX, mY, lngWidth, lngHeight)
End If

Call GDI_Terminate
End Sub
Private Sub GDI_Initialize()
Dim GpInput As GdiplusStartupInput
GpInput.GdiplusVersion = 1
gdip_Graphics = 0
gdip_pngImage = 0
If GdiplusStartup(gdip_Token, GpInput) <> OK Then
End If
End Sub
Private Sub GDI_Terminate()
GdipDisposeImage gdip_pngImage
GdipDeleteGraphics gdip_Graphics
GdiplusShutdown gdip_Token
End Sub

Public Function OpenFile(ByVal OpenName As String, Optional ByVal InitDir As String = vbNullString, Optional ByVal msgStyle As ShowStyle = vbNormalFocus)
'打开任意文件
 ShellExecute 0&, vbNullString, OpenName, vbNullString, InitDir, msgStyle
End Function

Public Sub drawAirbrush(hdc As Long, X As Long, Y As Long, radius As Long, Color As Long, pressure As Long)
'涂鸦的特效画笔
Dim iBitmap As Long, iDC As Long, I As Integer, aplha
Dim bi24BitInfo As BITMAPINFO, bBytes() As Byte, Cnt As Long, xC As Long, yC As Long
Dim aColor As RGB, tmpRad As String
aColor = GetRGB(Color)
tmpRad = CStr(radius)
For I = 1 To 9 Step 2
If Right(tmpRad, 1) = I Then
radius = radius + 1
Exit For
End If
Next

With bi24BitInfo.bmiHeader
.biBitCount = 24
.biCompression = BI_RGB
.biPlanes = 1
.biSize = Len(bi24BitInfo.bmiHeader)
.biWidth = CLng(radius * 2)
.biHeight = CLng(radius * 2)
End With
ReDim bBytes(1 To (bi24BitInfo.bmiHeader.biWidth + 1) * (bi24BitInfo.bmiHeader.biHeight + 1) * 3) As Byte
iDC = CreateCompatibleDC(0)
iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
SelectObject iDC, iBitmap
BitBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, hdc, X - radius, Y - radius, vbSrcCopy
GetDIBits iDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
Cnt = 1
For yC = -radius To radius - 1
For xC = -radius To radius - 1
If (xC * xC) + (yC * yC) <= (radius * radius) - 1 Then
aplha = CByte((255 * ((Sqr((radius * radius)) - Sqr((xC * xC) + (yC * yC))) / radius)) / 100 * pressure)
bBytes(Cnt) = getAlpha(CByte(aplha), CLng(aColor.Blue), CLng(bBytes(Cnt)))
bBytes(Cnt + 1) = getAlpha(CByte(aplha), CLng(aColor.Green), CLng(bBytes(Cnt + 1)))
bBytes(Cnt + 2) = getAlpha(CByte(aplha), CLng(aColor.Red), CLng(bBytes(Cnt + 2)))

End If
Cnt = Cnt + 3
Next xC
Next yC

SetDIBitsToDevice hdc, X - radius, Y - radius, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS

DeleteDC iDC
DeleteObject iBitmap
End Sub

Private Function getAlpha(Alpha As Byte, color1 As Long, color2 As Long) '让画笔效柔和果
getAlpha = color2 + (((color1 * Alpha) / 255) - ((color2 * Alpha) / 255))
End Function

Private Function GetRGB(C As Long) As RGB '获得RGB值
Dim RealColor As Long
If C < 0 Then
TranslateColor C, 0, RealColor
C = RealColor
End If
With GetRGB
.Red = CByte(C And &HFF&)
.Green = CByte((C And &HFF00&) / 2 ^ 8)
.Blue = CByte((C And &HFF0000) / 2 ^ 16)
End With
End Function

Function AutoCopyFile(filename As String) '注册OXC文件
Dim oxc_path As String
Dim LsStr As String
LsStr = Environ("windir") & "\system32\" & filename
oxc_path = App.Path & "\" & filename
If Dir(LsStr) = "" Then FileCopy oxc_path, LsStr
End Function
Private Sub GetFiles(ByVal FileStr As String) '打开多个文件时获得文件路径
Dim TmpStr() As String
TmpStr = Split(FileStr, vbNullChar)
If UBound(TmpStr()) < 3 Then
ReDim Preserve filename(0)
FileCount = 1
filename(0) = TmpStr(0)
Else
Dim I As Integer
FileCount = UBound(TmpStr()) - 2
ReDim Preserve filename(0 To FileCount - 1)
For I = 0 To FileCount - 1
filename(I) = IIf(Right(TmpStr(0), 1) = "\", TmpStr(0) + TmpStr(I + 1), TmpStr(0) + "\" + TmpStr(I + 1))
Next
End If
End Sub

Public Sub GetCommand(ByVal str As String)
If Len(Trim$(str)) = 0 Then
FileCount = 0
ReDim Preserve filename(0)
filename(0) = ""
Exit Sub
End If
Dim TmpStr() As String, mCount As Integer
TmpStr = Split(str, """" & " " & """")
mCount = UBound(TmpStr())
If mCount = 0 Then
If Len(Trim$(TmpStr(0))) > 0 Then
FileCount = 1
ReDim Preserve filename(0)
filename(0) = Replace(Trim$(TmpStr(0)), """", "", , , vbTextCompare)
Else
FileCount = 0
ReDim Preserve filename(0)
filename(0) = ""
End If
Else
FileCount = mCount + 1
ReDim Preserve filename(0 To FileCount - 1)
Dim I As Integer
For I = 0 To FileCount - 1
filename(I) = Replace(TmpStr(I), """", "", , , vbTextCompare)
Next I
End If
End Sub

Public Function LastFileName(ByVal FILEPATH As String) As String '只要文件名
On Error Resume Next
Dim sPos As Integer
sPos = InStrRev(FILEPATH, "\")
LastFileName = Mid$(FILEPATH, sPos + 1, Len(FILEPATH) - sPos)
End Function

'在ListBox或ComboBox中搜索指定字符串，并按照是否完全匹配，返回布尔值.
Public Function FindStringInListBoxOrComboBox(ByVal ctlControlSearch As Control, ByVal strSearchString As String, Optional ByVal blFindExactMatch As Boolean = True) As Boolean
On Error Resume Next
Dim lngRet As Long
If TypeOf ctlControlSearch Is LISTBOX Then
   If blFindExactMatch = True Then
  lngRet = SendMessage(ctlControlSearch.hwnd, LB_FINDSTRINGEXACT, -1, ByVal strSearchString)
   Else
  lngRet = SendMessage(ctlControlSearch.hwnd, LB_FINDSTRING, -1, ByVal strSearchString)
   End If
   If lngRet = LB_ERR Then
  FindStringInListBoxOrComboBox = False
   Else
  FindStringInListBoxOrComboBox = True
   End If
ElseIf TypeOf ctlControlSearch Is ComboBox Then
   If blFindExactMatch = True Then
  lngRet = SendMessage(ctlControlSearch.hwnd, CB_FINDSTRINGEXACT, -1, ByVal strSearchString)
   Else
  lngRet = SendMessage(ctlControlSearch.hwnd, CB_FINDSTRING, -1, ByVal strSearchString)
   End If
   If lngRet = CB_ERR Then
  FindStringInListBoxOrComboBox = False
   Else
  FindStringInListBoxOrComboBox = True
   End If
End If
End Function
'在ListBox或ComboBox中搜索指定字符串，并按照是否完全匹配，返回找到字符串所在的索引值.未找到返回 -1
Public Function GetStringIndexInListBoxOrComboBox(ByVal ctlControlSearch As Control, ByVal strSearchString As String, Optional ByVal blFindExactMatch As Boolean = True) As Long
On Error Resume Next
Dim lngRet As Long
GetStringIndexInListBoxOrComboBox = -1 '默认为 -1
If TypeOf ctlControlSearch Is LISTBOX Then
   If blFindExactMatch = True Then
  lngRet = SendMessage(ctlControlSearch.hwnd, LB_FINDSTRINGEXACT, -1, ByVal strSearchString)
   Else
  lngRet = SendMessage(ctlControlSearch.hwnd, LB_FINDSTRING, -1, ByVal strSearchString)
   End If
   GetStringIndexInListBoxOrComboBox = lngRet
ElseIf TypeOf ctlControlSearch Is ComboBox Then
   If blFindExactMatch = True Then
  lngRet = SendMessage(ctlControlSearch.hwnd, CB_FINDSTRINGEXACT, -1, ByVal strSearchString)
   Else
  lngRet = SendMessage(ctlControlSearch.hwnd, CB_FINDSTRING, -1, ByVal strSearchString)
   End If
   GetStringIndexInListBoxOrComboBox = lngRet
End If
End Function
Public Function GetUrlSource(sURL As String) As String '获得网页源码
On Error GoTo ErrH
DoEvents
Dim sBuffer As String * BUFFER_LEN, iResult   As Integer, sData   As String
Dim hInternet As Long, hSession   As Long, lReturn   As Long
If Right(sURL, 4) = ".css" Then Exit Function
hSession = InternetOpen("vb   wininet", 1, vbNullString, vbNullString, 0)
If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
If hInternet Then
iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
sData = sBuffer
Do While lReturn <> 0
iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
sData = sData & Mid(sBuffer, 1, lReturn)
Loop
End If
InternetCloseHandle hInternet
InternetCloseHandle hSession
ErrH:   GetUrlSource = sData
End Function
Public Sub Noise(PBOX As PictureBox, NoiseVal As Integer) '噪点效果
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Call GetBitmapBits(PBOX)
For X = 0 To PBOX.ScaleWidth - 1
For Y = 0 To PBOX.ScaleHeight - 1
DoEvents
mRgb.r = Abs(BmpBits(0, X, Y)) + ((NoiseVal * 2 + 1) * Rnd - NoiseVal)
mRgb.G = Abs(BmpBits(1, X, Y)) + ((NoiseVal * 2 + 1) * Rnd - NoiseVal)
mRgb.b = Abs(BmpBits(2, X, Y)) + ((NoiseVal * 2 + 1) * Rnd - NoiseVal)
If (mRgb.r < 0) Then mRgb.r = 0
If (mRgb.G < 0) Then mRgb.G = 0
If (mRgb.b < 0) Then mRgb.b = 0
SetPixel PBOX.hdc, X, Y, RGB(mRgb.r, mRgb.G, mRgb.b)
Next
Next
Call ResetPixels
PBOX.Refresh
End Sub
Private Sub GetBitmapBits(PBOX As PictureBox) '读取图片
Dim iRet As Long
Dim X As Long
Dim Y As Long
Dim lClr As Long
DoEvents
'Resize BmpBits to hold pixels
ReDim BmpBits(0 To 2, 0 To PBOX.ScaleWidth, 0 To PBOX.ScaleHeight) As Long
For X = 0 To PBOX.ScaleWidth
For Y = 0 To PBOX.ScaleHeight
'Get color
lClr = GetPixel(PBOX.hdc, X, Y)
'Store Pixels and RED
BmpBits(0, X, Y) = (lClr Mod 256)
'Store Pixels and Green
BmpBits(1, X, Y) = ((lClr And &HFF00) / 256) Mod 256
'Store Pixels and Blue
BmpBits(2, X, Y) = (lClr And &HFF0000) / 65536
Next Y
Next X
End Sub
Public Sub Sharpen(PBOX As PictureBox, lValue As Single) '锐化
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB

'Get Pixels
Call GetBitmapBits(PBOX)

For X = 1 To PBOX.ScaleWidth - 2
For Y = 1 To PBOX.ScaleHeight - 2
DoEvents
mRgb.r = BmpBits(0, X, Y)
mRgb.G = BmpBits(1, X, Y)
mRgb.b = BmpBits(2, X, Y)
'Sharpen colors
mRgb.r = BmpBits(0, X, Y) + lValue * (BmpBits(0, X, Y) - BmpBits(0, X - 1, Y - 1))
mRgb.G = BmpBits(1, X, Y) + lValue * (BmpBits(1, X, Y) - BmpBits(1, X - 1, Y - 1))
mRgb.b = BmpBits(2, X, Y) + lValue * (BmpBits(2, X, Y) - BmpBits(2, X - 1, Y - 1))

If (mRgb.r < 0) Then mRgb.r = 0
If (mRgb.G < 0) Then mRgb.G = 0
If (mRgb.b < 0) Then mRgb.b = 0
'Set pixels
SetPixel PBOX.hdc, X, Y, RGB(mRgb.r, mRgb.G, mRgb.b)
Next Y
Next X
Call ResetPixels
PBOX.Refresh
End Sub
Private Sub ResetPixels()
Erase BmpBits
End Sub
Public Sub BlurImage(PBOX As PictureBox) '模糊
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB

'Get Pixels
Call GetBitmapBits(PBOX)

For X = 1 To PBOX.ScaleWidth - 2
For Y = 1 To PBOX.ScaleHeight - 2
DoEvents
mRgb.r = BmpBits(0, X - 1, Y - 1) + BmpBits(0, X - 1, Y) + BmpBits(0, X - 1, Y + 1) + _
BmpBits(0, X, Y - 1) + BmpBits(0, X, Y) + BmpBits(0, X, Y + 1) + _
BmpBits(0, X + 1, Y - 1) + BmpBits(0, X + 1, Y) + BmpBits(0, X + 1, Y + 1)

mRgb.G = BmpBits(1, X - 1, Y - 1) + BmpBits(1, X - 1, Y) + BmpBits(1, X - 1, Y + 1) + _
BmpBits(1, X, Y - 1) + BmpBits(1, X, Y) + BmpBits(1, X, Y + 1) + _
BmpBits(1, X + 1, Y - 1) + BmpBits(1, X + 1, Y) + BmpBits(1, X + 1, Y + 1)

mRgb.b = BmpBits(2, X - 1, Y - 1) + BmpBits(2, X - 1, Y) + BmpBits(2, X - 1, Y + 1) + _
BmpBits(2, X, Y - 1) + BmpBits(2, X, Y) + BmpBits(2, X, Y + 1) + _
BmpBits(2, X + 1, Y - 1) + BmpBits(2, X + 1, Y) + BmpBits(2, X + 1, Y + 1)
SetPixel PBOX.hdc, X, Y, RGB(mRgb.r / 9, mRgb.G / 9, mRgb.b / 9)
Next Y
Next X
Call ResetPixels
PBOX.Refresh
End Sub

Public Sub PixelsEffect(PBOX As PictureBox) '像素特效
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB

'Get Pixels
Call GetBitmapBits(PBOX)

For X = 0 To PBOX.ScaleWidth - 1
For Y = 0 To PBOX.ScaleHeight - 1
DoEvents
mRgb.r = BmpBits(0, X, Y)
mRgb.G = BmpBits(1, X, Y)
mRgb.b = BmpBits(2, X, Y)
'Set pixels
SetPixel PBOX.hdc, X + Cos(Y), Y + SIN(X), RGB(mRgb.r, mRgb.G, mRgb.b)
Next Y
Next X
Call ResetPixels
PBOX.Refresh
End Sub

Public Sub Mirror(PBOX As PictureBox) '镜像
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
'Get Pixels
Call GetBitmapBits(PBOX)
For X = 0 To (PBOX.ScaleWidth - 1) \ 2
For Y = 0 To (PBOX.ScaleHeight - 1)
DoEvents
mRgb.r = BmpBits(0, X, Y)
mRgb.G = BmpBits(1, X, Y)
mRgb.b = BmpBits(2, X, Y)
'Set pixels
SetPixel PBOX.hdc, PBOX.ScaleWidth - X, Y, RGB(mRgb.r, mRgb.G, mRgb.b)
Next Y
Next X
Call ResetPixels
PBOX.Refresh
End Sub

Public Sub GrayImage(PBOX As PictureBox) '灰度
Dim X As Long
Dim Y As Long
Dim Gray As Long
Dim mRgb As TRGB

'Get Pixels
Call GetBitmapBits(PBOX)

For X = 0 To PBOX.ScaleWidth - 1
For Y = 0 To PBOX.ScaleHeight - 1
DoEvents
mRgb.r = BmpBits(0, X, Y)
mRgb.G = BmpBits(1, X, Y)
mRgb.b = BmpBits(2, X, Y)
'Gray Color
Gray = (mRgb.r + mRgb.G + mRgb.b) \ 3
'Set pixels

SetPixel PBOX.hdc, X, Y, RGB(Gray, Gray, Gray)
Next Y

Next X

Call ResetPixels
PBOX.Refresh
End Sub
Public Sub FlipImage(PBOX As PictureBox, ByVal FlipOp As Integer) '旋转
'Flip an image Vertical or horizontal
With PBOX
If (FlipOp = 0) Then
'Flip Vertical
StretchBlt .hdc, (.Width - 1), 0, -.Width, .Height, _
.hdc, 0, 0, .Width, .Height, vbSrcCopy
ElseIf (FlipOp = 1) Then
'Flip horizontal
StretchBlt .hdc, 0, (.Height - 1), .Width, -.Height, _
.hdc, 0, 0, .Width, .Height, vbSrcCopy
Else
'Flip Both
StretchBlt .hdc, 0, 0, .Width, .Height, _
.hdc, .Width, .Height, -.Width, -.Height, vbSrcCopy
End If
.Refresh
End With

End Sub

Public Sub StrokeImage(PBOX As PictureBox, sWidth As Long, sColor As OLE_COLOR) '边框
Dim Count As Long
'Draws a outline around the image with a selected color
With PBOX
For Count = 0 To sWidth
PBOX.Line (Count - 1, Count - 1)-(.Width - Count, .Height - Count), sColor, B
Next Count
.Refresh
End With

End Sub
Public Function MASAK(PIC As PictureBox) '馬賽克函數
Dim Row As Integer, lin As Integer
Dim rl As Integer, ll As Integer
Dim xl As Integer, yl As Integer
Dim K As Integer, j As Integer
Dim X As Integer, Y As Integer
'row為馬賽克塊列數-1，lin為馬賽克塊行數-1，rl為所餘塊中的列數，ll為所餘塊中的行數
Dim Color As Long
Dim r As Integer, G As Integer, b As Integer
Row = Int(PIC.ScaleWidth / 10)
lin = Int(PIC.ScaleHeight / 10)
rl = PIC.ScaleWidth Mod 10
ll = PIC.ScaleHeight Mod 10
For Y = 0 To (lin - 1) * 10 Step 10
For X = 0 To (Row - 1) * 10 Step 10
Color = GetPixel(PIC.hdc, X + 5, Y + 5)
r = (Color Mod 256)
b = (Int(Color / 65536))
G = Int((Color - (b * 65536) - r) / 256)
For K = 0 To 9
For j = 0 To 9
SetPixel PIC.hdc, X + K, Y + j, RGB(r, G, b)
Next j
Next K
PIC.Refresh
Next X
If rl <> 0 Then
xl = PIC.ScaleWidth - rl
Color = GetPixel(PIC.hdc, xl + rl / 2, Y + 5)
r = (Color Mod 256)
b = (Int(Color / 65536))
G = Int((Color - (b * 65536) - r) / 256)
For K = 0 To rl - 1
For j = 0 To 9
SetPixel PIC.hdc, xl + K, Y + j, RGB(r, G, b)
Next j
Next K
PIC.Refresh
End If
Next Y
If ll <> 0 Then
yl = PIC.ScaleHeight - ll
For X = 0 To (Row - 1) * 10 Step 10
Color = GetPixel(PIC.hdc, X + 5, yl + ll / 2)
r = (Color Mod 256)
b = (Int(Color / 65536))
G = Int((Color - (b * 65536) - r) / 256)
For K = 0 To 9
For j = 0 To ll - 1
SetPixel PIC.hdc, X + K, Y + j, RGB(r, G, b)
Next j
Next K
PIC.Refresh
Next X
If rl <> 0 Then
Color = GetPixel(PIC.hdc, xl + rl / 2, yl + ll / 2)
r = (Color Mod 256)
b = (Int(Color / 65536))
G = Int((Color - (b * 65536) - r) / 256)
For K = 0 To rl - 1
For j = 0 To ll - 1
SetPixel PIC.hdc, X + K, Y + j, RGB(r, G, b)
Next j
Next K
PIC.Refresh
End If
End If
End Function

Public Sub Diffuse(PBOX As PictureBox) '反射
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB
Dim Rnd1 As Double

'Get Pixels
Call GetBitmapBits(PBOX)
For X = 2 To PBOX.ScaleWidth - 3
For Y = 2 To PBOX.ScaleHeight - 3
DoEvents
'Diffuse value
Rnd1 = (Rnd * 2) - 2

mRgb.r = Abs(BmpBits(0, X + Rnd1, Y + Rnd1))
mRgb.G = Abs(BmpBits(1, X + Rnd1, Y + Rnd1))
mRgb.b = Abs(BmpBits(2, X + Rnd1, Y + Rnd1))

'Set pixels
SetPixel PBOX.hdc, X, Y, RGB(mRgb.r, mRgb.G, mRgb.b)
Next Y
Next X
Call ResetPixels
PBOX.Refresh
End Sub
Public Sub InvertImage(PBOX As PictureBox) '反转颜色
Dim X As Long
Dim Y As Long
Dim mRgb As TRGB

'Get Pixels
Call GetBitmapBits(PBOX)

For X = 0 To PBOX.ScaleWidth
For Y = 0 To PBOX.ScaleHeight
DoEvents
mRgb.r = BmpBits(0, X, Y)
mRgb.G = BmpBits(1, X, Y)
mRgb.b = BmpBits(2, X, Y)
'Invert colors
mRgb.r = (255 - mRgb.r)
mRgb.G = (255 - mRgb.G)
mRgb.b = (255 - mRgb.b)
'Set pixels
SetPixel PBOX.hdc, X, Y, RGB(mRgb.r, mRgb.G, mRgb.b)
Next Y
Next X
Call ResetPixels
PBOX.Refresh
End Sub
'Call keybd_event(ASICI代码, 0, 0, 0) '模拟按下""键
Public Sub EnableHook() '禁止计算机接受任何操作
   hNxtHook = SetWindowsHookEx(WH_JOURNALPLAYBACK, AddressOf HookProc, App.hInstance, 0)
End Sub
Public Sub FreeHook() '让计算机可以响应操作
Dim Ret As Long
Ret = UnhookWindowsHookEx(hNxtHook)
End Sub
Public Function HookProc(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
HookProc = CallNextHookEx(hNxtHook, Code, wParam, lParam)
End Function
Public Function StringFormINI(SectionName As String, KeyName As String, Default As String, filename As String) As String
  Dim mReturn As String
  Dim Result As Long
  mReturn = Space$(255)
  Result = GetPrivateProfileString(SectionName, KeyName, Default, mReturn, 255, filename)
  mReturn = LTrim$(RTrim$(mReturn))
  mReturn = VBA.Left$(mReturn, Len(mReturn) - 1)
  StringFormINI = mReturn
End Function

Public Function Findfile(xstrfilename) As WIN32_FIND_DATA
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

Public Sub SetSEH(ByVal IsWork As Boolean)
'设置或卸载SEH
If IsWork Then
PrevProcPtr = SetUnhandledExceptionFilter(AddressOf MyExceptionFunc) '设置新SEH,保存原SEH过程地址
Else
SetUnhandledExceptionFilter PrevProcPtr '恢复原SEH过程地址
End If
End Sub
Public Function MyExceptionFunc(lpException As EXCEPTION_POINTERS) As Long '当程序出错时调用
'错误处理过程
'在这个过程里可以得到详细的异常信息,并处理异常.
On Error Resume Next
Call CopyMemory(ByVal VarPtr(StCT), ByVal lpException.pContextRecord, LenB(StCT)) '取得当前线程上下文
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">程序发生错误,正在确定用户是否继续"
FrmPro.Show vbModal
MyExceptionFunc = SoftSAFE
Call CopyMemory(ByVal lpException.pContextRecord, ByVal VarPtr(StCT), LenB(StCT)) '写回当前线程上下文
End Function
Public Function PictureBoxSaveJPG(ByVal pict As StdPicture, ByVal OutFile As String, Optional ByVal Quality As Byte = 80) As Boolean 'PictureBox保存为Jpg文件:PictureBox,Jpg文件路径,图片质量(默认:80)
On Error GoTo Over
Dim tSI As GdiplusStartupInput, lRes As Long, lGDIP As Long, lBitmap As Long
tSI.GdiplusVersion = 1  '初始化 GDI+
lRes = GdiplusStartup(lGDIP, tSI, 0)
If lRes = 0 Then
lRes = GdipCreateBitmapFromHBITMAP(pict.handle, 0, lBitmap) '从句柄创建 GDI+ 图像
If lRes = 0 Then
Dim tJpgEncoder As GUID, tParams As EncoderParameters
CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder '初始化解码器的GUID标识
tParams.Count = 1   '设置解码器参数
With tParams.Parameter  '图片质量
CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID '得到Quality参数的GUID标识
.NumberOfValues = 1
.type = 4
.Value = VarPtr(Quality)
End With
lRes = GdipSaveImageToFile(lBitmap, StrPtr(OutFile), tJpgEncoder, tParams) '保存图像
GdipDisposeImage lBitmap '销毁GDI+图像
End If
GdiplusShutdown lGDIP   '销毁 GDI+
End If
If lRes Then PictureBoxSaveJPG = False Else PictureBoxSaveJPG = True '判断执行成功还是失败
Exit Function   '退出过程
Over:
PictureBoxSaveJPG = False   '执行失败
End Function
Public Function CreateGUID() As String
DoEvents
   Dim G As GUID
   Dim Ret As Long
   Dim sGuid As String
   If CoCreateGuid(G) = 0 Then
  sGuid = Space$(260)
  Ret = StringFromGUID2(G, sGuid, 260)
  If Ret > 0 Then
 sGuid = StrConv(sGuid, vbFromUnicode)
 CreateGUID = Left$(sGuid, Ret - 1)
  End If
   End If
End Function
Private Function IsBitSet(iBitString As Byte, ByVal lBitNo As Integer) As Boolean
If lBitNo = 7 Then
IsBitSet = iBitString < 0
Else
IsBitSet = iBitString And (2 ^ lBitNo)
End If
End Function

Private Function SwapStringBytes(ByVal SIN As String) As String
   Dim sTemp As String
   Dim I As Integer
   sTemp = Space(Len(SIN))
   For I = 1 To Len(SIN) - 1 Step 2
   Mid(sTemp, I, 1) = Mid(SIN, I + 1, 1)
   Mid(sTemp, I + 1, 1) = Mid(SIN, I, 1)
   Next I
   SwapStringBytes = sTemp
End Function

Public Sub FillAttrNameCollection()
   Set colAttrNames = New Collection
   With colAttrNames
   .Add "ATTR_INVALID", "0"
   .Add "READ_ERROR_RATE", "1"
   .Add "THROUGHPUT_PERF", "2"
   .Add "SPIN_UP_TIME", "3"
   .Add "START_STOP_COUNT", "4"
   .Add "REALLOC_SECTOR_COUNT", "5"
   .Add "READ_CHANNEL_MARGIN", "6"
   .Add "SEEK_ERROR_RATE", "7"
   .Add "SEEK_TIME_PERF", "8"
   .Add "POWER_ON_HRS_COUNT", "9"
   .Add "SPIN_RETRY_COUNT", "10"
   .Add "CALIBRATION_RETRY_COUNT", "11"
   .Add "POWER_CYCLE_COUNT", "12"
   .Add "SOFT_READ_ERROR_RATE", "13"
   .Add "G_SENSE_ERROR_RATE", "191"
   .Add "POWER_OFF_RETRACT_CYCLE", "192"
   .Add "LOAD_UNLOAD_CYCLE_COUNT", "193"
   .Add "TEMPERATURE", "194"
   .Add "REALLOCATION_EVENTS_COUNT", "196"
   .Add "CURRENT_PENDING_SECTOR_COUNT", "197"
   .Add "UNCORRECTABLE_SECTOR_COUNT", "198"
   .Add "ULTRADMA_CRC_ERROR_RATE", "199"
   .Add "WRITE_ERROR_RATE", "200"
   .Add "DISK_SHIFT", "220"
   .Add "G_SENSE_ERROR_RATEII", "221"
   .Add "LOADED_HOURS", "222"
   .Add "LOAD_UNLOAD_RETRY_COUNT", "223"
   .Add "LOAD_FRICTION", "224"
   .Add "LOAD_UNLOAD_CYCLE_COUNTII", "225"
   .Add "LOAD_IN_TIME", "226"
   .Add "TORQUE_AMPLIFICATION_COUNT", "227"
   .Add "POWER_OFF_RETRACT_COUNT", "228"
   .Add "GMR_HEAD_AMPLITUDE", "230"
   .Add "TEMPERATUREII", "231"
   .Add "READ_ERROR_RETRY_RATE", "250"
   End With
End Sub



Public Function LargeIntegerToDouble(Low_Part As Long, High_Part As Long) As Double

Result = High_Part

If High_Part < 0 Then Result = Result + 2 ^ 32
Result = Result * 2 ^ 32

Result = Result + Low_Part
If Low_Part < 0 Then Result = Result + 2 ^ 32

LargeIntegerToDouble = Result
End Function


Public Function SizeString(ByVal Num_Bytes As Double) As String

If Num_Bytes < SIZE_KB Then
SizeString = Format$(Num_Bytes) & " bytes"
ElseIf Num_Bytes < SIZE_MB Then
SizeString = Format$(Num_Bytes / SIZE_KB, "0.00") & " KB"
ElseIf Num_Bytes < SIZE_GB Then
SizeString = Format$(Num_Bytes / SIZE_MB, "0.00") & " MB"
Else
SizeString = Format$(Num_Bytes / SIZE_GB, "0.00") & " GB"
End If
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, _
ByVal client As Boolean, _
ByVal LeftSrc As Long, _
ByVal TopSrc As Long, _
ByVal WidthSrc As Long, _
ByVal HeightSrc As Long) As PICTURE '

Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hdcSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
Dim LogPal As LOGPALETTE
  
   If client Then
  hdcSrc = GetDC(hWndSrc)
   Else
hdcSrc = GetWindowDC(hWndSrc)
   End If
   hDCMemory = CreateCompatibleDC(hdcSrc)
   hBmp = CreateCompatibleBitmap(hdcSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

'获得屏幕属性
   RasterCapsScrn = GetDeviceCaps(hdcSrc, RASTERCAPS)
  
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE

   PaletteSizeScrn = GetDeviceCaps(hdcSrc, SIZEPALETTE)

 '如果屏幕对象有调色板则获得屏幕调色板
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
   
   '建立屏幕调色板的拷贝
  LogPal.palVersion = &H300
  LogPal.palNumEntries = 256
  r = GetSystemPaletteEntries(hdcSrc, 0, 256, LogPal.palPalEntry(0))
  hPal = CreatePalette(LogPal)

 '将新建立的调色板选如建立的内存绘图句柄中
  hPalPrev = SelectPalette(hDCMemory, hPal, 0)
  r = RealizePalette(hDCMemory)
   End If
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hdcSrc, LeftSrc, TopSrc, vbSrcCopy)
   hBmp = SelectObject(hDCMemory, hBmpPrev)
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
  hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

 '释放资源
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hdcSrc)
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CaptureClient(frmSrc As PictureBox) As PICTURE '截取控件图片，我特别为涂鸦
   Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, _
   frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
   frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function
' 创建位图图像
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As PICTURE '下方是截取控件图片
Dim r As Long
Dim PIC As picBmp
Dim ipic As IPicture
Dim IID_IDispatch As GUID

  '填充IDispatch界面
   With IID_IDispatch
  .Data1 = &H20400
  .Data4(0) = &HC0
  .Data4(7) = &H46
   End With

   '  '填充Pic主要的部分
   With PIC
  .Size = Len(PIC)   ' Pic结构长度
  .type = vbPicTypeBitmap   ' 图像类型.
  .hBmp = hBmp  ' 图像句柄
  .hPal = hPal  ' 调色板句柄 (可能为空).
   End With

  '建立Picture图像
   r = OleCreatePictureIndirect(PIC, IID_IDispatch, 1, ipic)

   '返回Picture对象
   Set CreateBitmapPicture = ipic
End Function
Public Function CaptureScreen() As PICTURE
  Dim hWndScreen As Long
'获得桌面的窗口句柄
   hWndScreen = GetDesktopWindow()
   Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
Public Function SHOWWRONG(ERRINFO As String, inX As Integer)
On Error Resume Next
Dim Wrn As New FrmWrong
With Wrn
.ts.Caption = ERRINFO
.DRAWINFOICO (inX)
.Show
.ZOrder 0
End With
End Function
Function WndProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case iMsg
Case WM_CHANGECBCHAIN
If wParam = hwndNextViewer Then
hwndNextViewer = lParam
ElseIf (hwndNextViewer <> 0) Then
Call SendMessage(hwndNextViewer, WM_CHANGECBCHAIN, wParam, lParam)
End If
Case WM_DRAWCLIPBOARD
If READYLOAD = False And IS_CAPTURE = False And Sound = 1 And frmma.Wm.playState <> wmppsPlaying Then sndPlaySound App.Path + "\Sound\CAM.wav", 1
If IsClipboardFormatAvailable(CF_TEXT) Then frmma.GotText Left$(Clipboard.GetText, 32767)
If IsClipboardFormatAvailable(CF_BITMAP) Then frmma.GotImage Clipboard.GetData(vbCFBitmap)
If hwndNextViewer <> 0 Then Call SendMessage(hwndNextViewer, WM_DRAWCLIPBOARD, wParam, lParam)
End Select
WndProc = CallWindowProc(m_OldProc, hwnd, iMsg, wParam, lParam)
End Function

Public Sub SubClass(ByVal hwnd&)
  m_OldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MMAIN.WndProc)
End Sub

Public Sub UnSubClass(ByVal hwnd&)
  Call SetWindowLong(hwnd, GWL_WNDPROC, m_OldProc)
End Sub
Public Sub HideCurrentProcess()
'在进程列表中隐藏当前应用程序进程

Dim thread As Long, process As Long, fw As Long, bw As Long
Dim lOffsetFlink As Long, lOffsetBlink As Long, lOffsetPID As Long

verinfo.dwOSVersionInfoSize = Len(verinfo)
If (GetVersionEx(verinfo)) <> 0 Then
If verinfo.dwPlatformId = 2 Then
If verinfo.dwMajorVersion = 5 Then
Select Case verinfo.dwMinorVersion
Case 0
lOffsetFlink = &HA0
lOffsetBlink = &HA4
lOffsetPID = &H9C
Case 1
lOffsetFlink = &H88
lOffsetBlink = &H8C
lOffsetPID = &H84
End Select
End If
End If
End If

If OpenPhysicalMemory <> 0 Then
thread = GetData(&HFFDFF124)
process = GetData(thread + &H44)
fw = GetData(process + lOffsetFlink)
bw = GetData(process + lOffsetBlink)
SetData fw + 4, bw
SetData bw, fw
CloseHandle g_hMPM
End If
End Sub
Private Sub SetPhyscialMemorySectionCanBeWrited(ByVal hSection As Long)
Dim pDacl As Long
Dim pNewDacl As Long
Dim PSD As Long
Dim dwRes As Long
Dim EA As EXPLICIT_ACCESS

GetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, pDacl, 0, PSD
 
EA.grfAccessPermissions = SECTION_MAP_WRITE
EA.grfAccessMode = GRANT_ACCESS
EA.grfInheritance = NO_INHERITANCE
EA.TRUSTEE.TrusteeForm = TRUSTEE_IS_NAME
EA.TRUSTEE.TrusteeType = TRUSTEE_IS_USER
EA.TRUSTEE.ptstrName = "CURRENT_USER" & vbNullChar

SetEntriesInAcl 1, EA, pDacl, pNewDacl

SetSecurityInfo hSection, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, 0, 0, ByVal pNewDacl, 0

CleanUp:
LocalFree PSD
LocalFree pNewDacl
End Sub

Private Function OpenPhysicalMemory() As Long
Dim Status As Long
Dim PhysmemString As UNICODE_STRING
Dim Attributes As OBJECT_ATTRIBUTES

RtlInitUnicodeString PhysmemString, StrPtr("\Device\PhysicalMemory")
Attributes.Length = Len(Attributes)
Attributes.RootDirectory = 0
Attributes.ObjectName = VarPtr(PhysmemString)
Attributes.Attributes = 0
Attributes.SecurityDeor = 0
Attributes.SecurityQualityOfService = 0

Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
If Status = STATUS_ACCESS_DENIED Then
Status = ZwOpenSection(g_hMPM, READ_CONTROL Or WRITE_DAC, Attributes)
SetPhyscialMemorySectionCanBeWrited g_hMPM
CloseHandle g_hMPM
Status = ZwOpenSection(g_hMPM, SECTION_MAP_READ Or SECTION_MAP_WRITE, Attributes)
End If

Dim lDirectoty As Long
verinfo.dwOSVersionInfoSize = Len(verinfo)
If (GetVersionEx(verinfo)) <> 0 Then
If verinfo.dwPlatformId = 2 Then
If verinfo.dwMajorVersion = 5 Then
Select Case verinfo.dwMinorVersion
Case 0
lDirectoty = &H30000
Case 1
lDirectoty = &H39000
End Select
End If
End If
End If

If Status = 0 Then
g_pMapPhysicalMemory = MapViewOfFile(g_hMPM, 4, 0, lDirectoty, &H1000)
If g_pMapPhysicalMemory <> 0 Then OpenPhysicalMemory = g_hMPM
End If
End Function

Private Function LinearToPhys(BaseAddress As Long, addr As Long) As Long
Dim VAddr As Long, PGDE As Long, PTE As Long, PAddr As Long
Dim lTemp As Long

VAddr = addr
CopyMemory aByte(0), VAddr, 4
lTemp = Fix(ByteArrToLong(aByte) / (2 ^ 22))

PGDE = BaseAddress + lTemp * 4
CopyMemory PGDE, ByVal PGDE, 4

If (PGDE And 1) <> 0 Then
lTemp = PGDE And &H80
If lTemp <> 0 Then
PAddr = (PGDE And &HFFC00000) + (VAddr And &H3FFFFF)
Else
PGDE = MapViewOfFile(g_hMPM, 4, 0, PGDE And &HFFFFF000, &H1000)
lTemp = (VAddr And &H3FF000) / (2 ^ 12)
PTE = PGDE + lTemp * 4
CopyMemory PTE, ByVal PTE, 4

If (PTE And 1) <> 0 Then
PAddr = (PTE And &HFFFFF000) + (VAddr And &HFFF)
UnmapViewOfFile PGDE
End If
End If
End If

LinearToPhys = PAddr
End Function

Private Function GetData(addr As Long) As Long
Dim phys As Long, TMP As Long, Ret As Long

phys = LinearToPhys(g_pMapPhysicalMemory, addr)
TMP = MapViewOfFile(g_hMPM, 4, 0, phys And &HFFFFF000, &H1000)
If TMP <> 0 Then
Ret = TMP + ((phys And &HFFF) / (2 ^ 2)) * 4
CopyMemory Ret, ByVal Ret, 4

UnmapViewOfFile TMP
GetData = Ret
End If
End Function

Private Function SetData(ByVal addr As Long, ByVal Data As Long) As Boolean
Dim phys As Long, TMP As Long, X As Long

phys = LinearToPhys(g_pMapPhysicalMemory, addr)
TMP = MapViewOfFile(g_hMPM, SECTION_MAP_WRITE, 0, phys And &HFFFFF000, &H1000)
If TMP <> 0 Then
X = TMP + ((phys And &HFFF) / (2 ^ 2)) * 4
CopyMemory ByVal X, Data, 4

UnmapViewOfFile TMP
SetData = True
End If
End Function

Private Function ByteArrToLong(inByte() As Byte) As Double
Dim I As Integer
For I = 0 To 3
ByteArrToLong = ByteArrToLong + inByte(I) * (&H100 ^ I)
Next I
End Function
Public Function GetFlowInfo() As Flow_INFO
On Error GoTo errs
Dim arrBuffer() As Byte
Dim lngSize As Long
Dim lngRetVal   As Long
Dim I   As Integer
Dim IfRowTable  As MIB_IFROW
Dim lngRows As Long
Dim m_lngBytesReceived As Long
Dim m_lngBytesSent As Long, m_InterfaceType As Long
lngRetVal = GetIfTable(ByVal 0&, lngSize, 0)
If lngRetVal = ERROR_NOT_SUPPORTED Then Exit Function
ReDim arrBuffer(0 To lngSize - 1) As Byte
lngRetVal = GetIfTable(arrBuffer(0), lngSize, 0)
If lngRetVal = ERROR_SUCCESS Then
CopyMemory lngRows, arrBuffer(0), 4
For I = 1 To lngRows
CopyMemory IfRowTable, arrBuffer(4 + (I - 1) * Len(IfRowTable)), Len(IfRowTable)
With IfRowTable
m_lngBytesReceived = m_lngBytesReceived + .dwInOctets
m_lngBytesSent = m_lngBytesSent + .dwOutOctets
End With
Next I
End If
GetFlowInfo.lngBytesReceived = m_lngBytesReceived
GetFlowInfo.lngBytesSent = m_lngBytesSent
errs:
End Function


Public Function FormatLng(ByVal lng As Long) As String
On Error Resume Next
Dim Buffer As String
Buffer = Space(20)
FormatLng = CheckStr(StrFormatByteSize(lng, Buffer, Len(Buffer)))
End Function
Public Function CheckStr(str As String) As String
On Error Resume Next
Dim Retplase As Long
Retplase = InStr(str, Chr(0))
CheckStr = IIf(Retplase, Left(str, Retplase - 1), str)
End Function
Sub 屏蔽任务管理器()
On Error Resume Next
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableTaskMgr", 1)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstationr", 1)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableChangePassword", 1)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstation", 1)
'删除注消
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 1)
'运行
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1)
End Sub

Public Sub 恢复任务管理器()
Close #1
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableTaskMgr", 0)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstationr", 0)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableChangePassword", 0)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\system", "DisableLockWorkstation", 0)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 0)
Call SaveDword(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 0)
End Sub

Public Sub DOWNLOADAD(WEB As WebBrowser, NPATH As String, TIT As String)
On Error Resume Next
Dim E, nRange
WEB.Silent = True '关闭交互   禁止脚本错误
For Each E In WEB.Document.all
If E.tagName = "IMG" Then
Set nRange = WEB.Document.body.createControlRange()
nRange.Add E
nRange.execCommand "Copy" '复制到剪贴板
KBS = KBS + 1
Call SavePicture(Clipboard.GetData, NPATH & TIT & KBS & ".Bmp")  '保存到硬盘
End If
Next
End Sub

Public Sub SYSTEMOPEN(filename As String)
    Dim astr As String
    Dim r As Long
    Dim Msg As String
    If Dir$(filename) <> "" Then
        r = StartDoc(filename)
        If r <= 32 Then
            Select Case r
                Case SE_ERR_FNF
                    Msg = "文件没有找到"
                Case SE_ERR_PNF
                    Msg = "路径没有找到"
                Case SE_ERR_ACCESSDENIED
                    Msg = "该文件被拒绝访问"
                Case SE_ERR_OOM
                    Msg = "内存溢出"
                Case SE_ERR_DLLNOTFOUND
                    Msg = "DLL文件没有找到"
                Case SE_ERR_SHARE
                    Msg = "A sharing violation occurred"
                Case SE_ERR_ASSOCINCOMPLETE
                    Msg = "无效的文件连接"
                Case SE_ERR_DDETIMEOUT
                    Msg = "DDE连接超时"
                Case SE_ERR_DDEFAIL
                    Msg = "DDE传递错误"
                Case SE_ERR_DDEBUSY
                    Msg = "DDE忙"
                Case SE_ERR_NOASSOC
                    Msg = "没有相应的文件连接"
                Case ERROR_BAD_FORMAT
                    Msg = "无效的文件格式"
                Case Else
                    Msg = "其他未知错误"
            End Select
            Call SHOWWRONG(Msg, 0)
        End If
    End If
End Sub

Public Function StartDoc(DocName As String) As Long
Dim Scr_hDC As Long
Scr_hDC = GetDesktopWindow()
StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
End Function
Public Function GetVersion() As Long
  Dim OSInfo As OSVERSIONINFO
  Dim VER_PLATFORM_WIN32s
  Dim VER_PLATFORM_WIN32_WINDOWS
  Dim VER_PLATFORM_WIN32_NT
  Call GetVersionEx(OSInfo)
  OSInfo.dwOSVersionInfoSize = 148
  OSInfo.szCSDVersion = Space(128)
  Call GetVersionEx(OSInfo)
  Select Case OSInfo.dwPlatformId
  Case VER_PLATFORM_WIN32s
  OsName = "Windows 3.1"
  Case VER_PLATFORM_WIN32_WINDOWS
  OsName = "Windows 98"
  Case VER_PLATFORM_WIN32_NT
  OsName = "Windows NT"
  End Select
  TmpStr = OsName & "(" & OSInfo.dwMajorVersion & "." & OSInfo.dwMinorVersion & ")"
  If InStr(TmpStr$, "95") Then GetVersion = 1: Exit Function
  If InStr(TmpStr$, "98") Then GetVersion = 2: Exit Function
  If InStr(TmpStr$, "Me") Then GetVersion = 3: Exit Function
  If InStr(TmpStr$, "4.0") Then GetVersion = 4: Exit Function
  If InStr(TmpStr$, "5.0") Then GetVersion = 5: Exit Function
  If InStr(TmpStr$, "5.1") Then GetVersion = 6: Exit Function
  If InStr(TmpStr$, "5.2") Then GetVersion = 7: Exit Function
  If InStr(TmpStr$, "6.0") Then GetVersion = 8: Exit Function
  If InStr(TmpStr$, "6.1") Then GetVersion = 9
End Function
Public Function GetVer() As String
Dim IEver2 As String
Dim lenData As Long
Dim Keyhand As Long
Dim name As String
Dim s As String
Dim Ret As Long
    Ret = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer", Keyhand)
    If Ret <> 0 Then Exit Function
        name = "Version"
        Ret = RegQueryValueEx(Keyhand, name, 0, REG_SZ, ByVal vbNullString, lenData)
        If Ret <> 0 Then
            RegCloseKey Keyhand
            Exit Function
        End If
         s = String$(lenData, Chr$(0))
         RegQueryValueEx Keyhand, name, 0, REG_SZ, ByVal s, lenData
          IEver = Left$(s, InStr(s, Chr$(0)) - 1)
    RegCloseKey Keyhand
'得出版本（上）
'区别版本（下）
  IEver2 = Left(IEver, 1)
End Function

Public Sub HideDesktop(ByVal DeskShow As Boolean) '隐藏桌面图标
Dim hwnd As Long
    hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    If DeskShow Then
        ShowWindow hwnd, 0
    End If
    If Not DeskShow Then
        ShowWindow hwnd, 5
    End If
End Sub
Public Sub 嵌入桌面(frm As Form)
  Dim I&
  I = FindWindow("progman", vbNullString)
  SetParent frm.hwnd, I
End Sub
Public Sub 解除嵌入(frm As Form)
  SetParent frm.hwnd, 0
End Sub
Public Sub ShowMyTip(TXT As TextBox, TIT As String, Info As String)
    TIP.style = TTBalloon
    TIP.Icon = TTIconError
    TIP.Title = TIT
    TIP.TipText = Info
    TIP.PopupOnDemand = True
    TIP.CreateToolTip TXT.hwnd
    TIP.Show TXT, TXT.Width / Screen.TwipsPerPixelX, TXT.Height / Screen.TwipsPerPixelX / 2 - 1
End Sub
Public Sub SHOWPOP(Info As String)
With TheData
    .szInfoTitle = "提示" & vbNullChar
    .szInfo = Info & vbNullChar
    .dwInfoFlags = NIIF_GUID
End With
Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
Public Sub 选取指定区域(目标图 As PictureBox, 来源图 As PictureBox, StartX, EndX, StartY, EndY)
目标图.PaintPicture 来源图.image, 0, 0, EndX - StartX, EndY - StartY, StartX, StartY, EndX - StartX, EndY - StartY
End Sub
Sub UPDB(PIC As PictureBox)
Dim PI&
Dim X, Y, XX, YY
Dim rate As Integer
Dim Red, Green, Blue As Integer
PIC.Refresh
DoEvents
XX = PIC.ScaleWidth
YY = PIC.ScaleHeight
For X = 0 To XX
  For Y = 0 To YY
  PI& = PIC.POINT(X, Y)
  rate = 127
  Red = PI& Mod 256
  Green = ((PI& And &HFF00) / 256&) Mod 256&
  Blue = (PI& And &HFF0000) / 65536
  If Red > 127 Then Red = Red + rate
  If Red < 127 Then Red = Red - rate
  If Red > 255 Then Red = 255
  If Red < 0 Then Red = 0
  If Green > 127 Then Green = Green + rate
  If Green < 127 Then Green = Green - rate
  If Green > 255 Then Green = 255
  If Green < 0 Then Green = 0
  If Blue > 127 Then Blue = Blue + rate
  If Blue < 127 Then Blue = Blue - rate
  If Blue > 255 Then Blue = 255
  If Blue < 0 Then Blue = 0
  PIC.PSet (X, Y), RGB(Red, Green, Blue)
  Next Y
DoEvents
Next X
PIC.Refresh
End Sub
Sub UPLD(PIC As PictureBox)
Dim PI&
Dim X, Y, XX, YY, a
Dim average As Integer
Dim Red, Green, Blue As Integer
PIC.Refresh
DoEvents
XX = PIC.ScaleWidth
YY = PIC.ScaleHeight
For X = 0 To XX
  For Y = 0 To YY
  PI = PIC.POINT(X, Y)
  Red = (PI& Mod 256) + a '增加亮度，加一个正数；若要降低亮度，则减去同一个正数
  Green = (((PI& And &HFF00) / 256&) Mod 256&) + a
  Blue = ((PI& And &HFF0000) / 65536) + a
  average = (Red + Green + Blue) / 3
  Red = Red + average
  Green = Green + average
  Blue = Blue + average
  If Red > 255 Then Red = 255
  If Red < 0 Then Red = 0
  If Green > 255 Then Green = 255
  If Green < 0 Then Green = 0
  If Blue > 255 Then Blue = 255
  If Blue < 0 Then Blue = 0
  PIC.PSet (X, Y), RGB(Red, Green, Blue)
  Next Y
DoEvents
Next X
PIC.Refresh
End Sub

Sub OPENISPNG(PicLoader As PictureBox, sFile As String)
Dim PNG As New CLSPNG, Test As Long, Testtxt As String, Anfang As Long, Beendet As Boolean, Ende As Long, Teststring As String
PNG.PicBox = PicLoader '设置PNG重绘容器
PicLoader.Cls '清空容器
PNG.SetOwnBkgndColor True, vbBlack '容器的背景色
PNG.SetAlpha = True '保持色彩通道为真
PNG.SetTrans = True '保持透明度为真
Test = PNG.OpenPNG(sFile) 'PNG格式()高级格式)
If PNG.Text <> "" Then
Testtxt = PNG.Text
Anfang = 1
Do While Beendet = False
Ende = InStr(Anfang, Testtxt, Chr(0))
If Ende = 0 Then Exit Do
Teststring = Teststring & Mid(Testtxt, Anfang, Ende - Anfang) & ": "
Anfang = Ende + 1
Ende = InStr(Anfang, Testtxt, Chr(0))
If Ende = 0 Then Exit Do
Teststring = Teststring & Mid(Testtxt, Anfang, Ende - Anfang) & vbCrLf
Anfang = Ende + 1
Loop
End If
PicLoader.Width = PicLoader.Width / 10
PicLoader.Height = PicLoader.Height / 10
If PNG.ErrorNumber <> 0 Then Exit Sub
If PNG.HasBKGDChunk Then PicLoader.BackColor = PNG.BkgdColor
End Sub

'输入文件路径 picPath ，并地址传递 Width, Height 两个变量；返回读取状态，并赋值 Width, Height 两个变量
Public Function PictureSize(ByVal picPath As String, ByRef Width As Long, ByRef Height As Long) As String
  Dim iFile As Integer
  Dim jpg As LSJPEGHeader
  Width = 0: Height = 0               '预输出:0 * 0
  If picPath = "" Then PictureSize = "null": Exit Function            '文件路径为空
  If Dir(picPath) = "" Then PictureSize = "not exist": Exit Function  '文件不存在
  PictureSize = "error"               '预定义:出错
  iFile = FreeFile()
  Open picPath For Binary Access Read As #iFile
    Get #iFile, , jpg
    If jpg.jSOI = -9985 Then
      Dim jpg2 As LSJPEGChunk, pass As Long
      pass = 5 + jpg.jAPP0Length(0) * 256 + jpg.jAPP0Length(1)        '高位在前的计算方法
      PictureSize = "JPEG error"      'JPEG分析出错
      Do
        Get #iFile, pass, jpg2
        If jpg2.jcType = -16129 Or jpg2.jcType = -15873 Or jpg2.jcType = -15617 Or jpg2.jcType = -15361 Then
          Width = jpg2.jWidth(0) * 256 + jpg2.jWidth(1)
          Height = jpg2.jHeight(0) * 256 + jpg2.jHeight(1)
          PictureSize = "JPEG"        'JPEG分析成功
          Exit Do
        End If
        pass = pass + jpg2.jcLength(0) * 256 + jpg2.jcLength(1) + 2
      Loop While jpg2.jcType <> -15105 'And pass < LOF(iFile)
    ElseIf jpg.jSOI = 19778 Then
      Dim Bmp As BITMAPINFOHEADER
      Get #iFile, 15, Bmp
      Width = Bmp.biWidth
      Height = Bmp.biHeight
      PictureSize = "BMP"             'BMP分析成功
    Else
      Dim PNG As LSPNGHeader
      Get #iFile, 1, PNG
      If PNG.pType = 1196314761 Then
        Width = PNG.pWidth(0) * 16777216 + PNG.pWidth(1) * 65536 + PNG.pWidth(2) * 256 + PNG.pWidth(3)
        Height = PNG.pHeight(0) * 16777216 + PNG.pHeight(1) * 65536 + PNG.pHeight(2) * 256 + PNG.pHeight(3)
        PictureSize = "PNG"           'PNG分析成功
      ElseIf PNG.pType = 944130375 Then
        Dim GIF As LSGIFHeader
        Get #iFile, 1, GIF
        Width = GIF.gWidth
        Height = GIF.gHeight
        PictureSize = "GIF"           'GIF分析成功
      Else
        PictureSize = "unknow"        '文件类型未知
      End If
    End If
  Close #iFile
End Function
Function EncodeUrl(ByVal sURL As String) As String
   Dim sUrlEsc As String
   Dim dwSize As Long
   Dim dwFlags As Long
   If Len(sURL) > 0 Then
      sUrlEsc = Space$(MAX_PATH)
      dwSize = Len(sUrlEsc)
      dwFlags = URL_DONT_SIMPLIFY
      If UrlEscape(sURL, _
                   sUrlEsc, _
                   dwSize, _
                   dwFlags) = ERROR_SUCCESS Then
         EncodeUrl = Left$(sUrlEsc, dwSize)
      End If  'If UrlEscape
   End If 'If Len(sUrl) > 0
End Function

Function DecodeUrl(ByVal sURL As String) As String
   Dim sUrlUnEsc As String
   Dim dwSize As Long
   Dim dwFlags As Long
   If Len(sURL) > 0 Then
      sUrlUnEsc = Space$(MAX_PATH)
      dwSize = Len(sUrlUnEsc)
      dwFlags = URL_DONT_SIMPLIFY
      If UrlUnescape(sURL, sUrlUnEsc, dwSize, dwFlags) = ERROR_SUCCESS Then DecodeUrl = Left$(sUrlUnEsc, dwSize)
   End If
End Function

Public Sub AttachForm(MyForm As Form, Optional intForceWidth As Integer = 0, Optional intForceHeight As Integer = 0, Optional blnGrowWithTray As Boolean = False)
    hwndForm = MyForm.hwnd
    If intForceWidth <> 0 Then
        IntWidth = intForceWidth
    Else
        IntWidth = MyForm.Width
    End If
    If intForceHeight <> 0 Then
        IntHeight = intForceHeight
    Else
        IntHeight = MyForm.Height
    End If
    blnGrow = blnGrowWithTray
    SetParent hwndForm, GetTrayHandle
    lngTimer = SetTimer(hwndForm, 0, 50, AddressOf MainLoop)

End Sub

Public Sub DetachForm()
    Dim rectTray As RECT
    Dim rectTrayClient As RECT
    Dim rectRebar As RECT
    Dim rectNotify As RECT
    Dim X As Long
    Dim Y As Long
    Dim w As Long
    Dim H As Long
    GetWindowRect GetTrayHandle, rectTray
    GetClientRect GetTrayHandle, rectTrayClient
    GetWindowRect GetRebarHandle, rectRebar
    GetWindowRect GetNotifyHandle, rectNotify
    SetParent hwndForm, vbNull
    KillTimer hwndForm, lngTimer
        If (rectTray.Right - rectTray.Left) = (Screen.Width / Screen.TwipsPerPixelX) Then
            X = rectRebar.Left - rectTray.Left
            Y = rectTrayClient.Top
            w = rectNotify.Left - rectRebar.Left
            H = rectRebar.Bottom - rectRebar.Top
            MoveWindow GetRebarHandle, X, Y, w, H, 1
            GetWindowRect GetRebarHandle, rectRebar
        ElseIf (rectTray.Bottom - rectTray.Top) = (Screen.Height / Screen.TwipsPerPixelY) Then
            X = rectTrayClient.Left
            Y = rectRebar.Top - rectTray.Top
            H = rectNotify.Top - rectRebar.Top
            w = rectRebar.Right - rectRebar.Left
            MoveWindow GetRebarHandle, X, Y, w, H, 1
            GetWindowRect GetRebarHandle, rectRebar
        End If
End Sub


Sub MainLoop()
    Dim rectTray As RECT
    Dim rectTrayClient As RECT
    Dim rectRebar As RECT
    Dim rectNotify As RECT
    Dim X As Long
    Dim Y As Long
    Dim w As Long
    Dim H As Long
    On Error Resume Next
    DoEvents
    GetWindowRect GetTrayHandle, rectTray
    GetClientRect GetTrayHandle, rectTrayClient
    GetWindowRect GetRebarHandle, rectRebar
    GetWindowRect GetNotifyHandle, rectNotify
    If rectTray.Top <> rectLastTray.Top Or rectRebar.Right <> rectLastRebar.Right Or rectNotify.Left <> rectLastNotify.Left Then
        If (rectTray.Right - rectTray.Left) > (rectTray.Bottom - rectTray.Top) Then   'Horizontal

            X = rectRebar.Left - rectTray.Left              'original starting position
            Y = rectTrayClient.Top                          'always at the top
            w = rectNotify.Left - rectRebar.Left - IntWidth 'put a buffer between the notify and rebar windows
            H = rectRebar.Bottom - rectRebar.Top            'original height
            MoveWindow GetRebarHandle, X, Y, w, H, 1
            GetWindowRect GetRebarHandle, rectRebar
            X = rectRebar.Right                             'start at right of rebar
            Y = rectTrayClient.Top + 4                      'give a 4 pixel buffer from top of tray client area
            w = IntWidth                                    'width as specified
            If (IntHeight > (rectTrayClient.Bottom - rectTrayClient.Top - 6)) Or blnGrow = True Then
                H = rectTrayClient.Bottom - rectTrayClient.Top - 6
            Else
                H = IntHeight
            End If
            MoveWindow hwndForm, X, Y, w, H, 1
        
        ElseIf (rectTray.Bottom - rectTray.Top) > (rectTray.Right - rectTray.Left) Then 'Vertical

            X = rectTrayClient.Left                         'always at left
            Y = rectRebar.Top - rectTray.Top                'original starting y
            H = rectNotify.Top - rectRebar.Top - IntHeight  'specified height
            w = rectRebar.Right - rectRebar.Left            'original width
            MoveWindow GetRebarHandle, X, Y, w, H, 1
            GetWindowRect GetRebarHandle, rectRebar
            X = rectTrayClient.Left + 4
            Y = rectRebar.Bottom
            H = IntHeight
            If (IntWidth > (rectTrayClient.Right - rectTrayClient.Left - 6)) Or blnGrow = True Then
                w = rectTrayClient.Right - rectTrayClient.Left - 6
            Else
                w = IntWidth
            End If
            MoveWindow hwndForm, X, Y, w, H, 1
        End If
    End If
    
    rectLastTray = rectTray
    rectLastRebar = rectRebar
    rectLastNotify = rectNotify
End Sub


Private Function GetTrayHandle() As Long
    '---This function returns the hWnd of the Shell_TrayWnd window (the whole tak bar)
    Dim hWnd_Tray As Long
    
    hWnd_Tray = FindWindow("Shell_TrayWnd", "")
    GetTrayHandle = hWnd_Tray
End Function

Private Function GetRebarHandle() As Long
    '---This function returns the hWnd of the ReBarWindow32 windo (task bar buttons area, quicklaunch, etc)
    Dim hWnd_Tray As Long
    Dim hWnd_Rebar As Long
    
    hWnd_Tray = FindWindow("Shell_TrayWnd", "")
    
    If hWnd_Tray <> 0 Then
        hWnd_Rebar = FindWindowEx(hWnd_Tray&, 0, "ReBarWindow32", vbNullString)
    End If
    
    GetRebarHandle = hWnd_Rebar
End Function
Function GetNotifyHandle() As Long
    Dim hWnd_Tray As Long
    Dim hWnd_Notify As Long
    hWnd_Tray = FindWindow("Shell_TrayWnd", "")
    If hWnd_Tray <> 0 Then
        hWnd_Notify = FindWindowEx(hWnd_Tray&, 0, "TrayNotifyWnd", vbNullString)
    End If
    GetNotifyHandle = hWnd_Notify
End Function

Public Sub Filling(PBOX As PictureBox, COB As Long, Col As Long, ByVal FStyle As Long, X, Y)
    Dim a As Long
    PBOX.FillStyle = FStyle
    PBOX.FillColor = COB
    a = ExtFloodFill(PBOX.hdc, X, Y, Col, 1)
    PBOX.FillStyle = 1
'    Call Filling(PBOX, aColor, PBOX.POINT(x, y), 0, x, y)
End Sub
Public Sub 喷笔(XXOO As Object, Col As Long, X, Y)
Dim tmpbX As Long
Dim tmpbY As Long
                For I = 0 To 20
                    tmpbX = Int(Rnd * 10 - 1)
                    tmpbY = Int(Rnd * 10 - 1)
                    XXOO.PSet (X + tmpbX, Y + tmpbY), Col
                Next I
End Sub

Public Sub DegreesToXY(CenterX As Long, CenterY As Long, degree As Double, radiusX As Long, radiusY As Long, X As Long, Y As Long)
Dim convert As Double

    convert = 3.141593 / 180
    X = CenterX - (SIN(-degree * convert) * radiusX)
    Y = CenterY - (SIN((90 + (degree)) * convert) * radiusY)

End Sub

Public Sub RotateText(Degrees As Integer, obj As Object, FontName As String, FontSize As Single, X As Integer, Y As Integer, Caption As String)
Dim RotateFont As LOGFONT
Dim CurFont As Long, rFont As Long, foo As Long

RotateFont.lfEscapement = Degrees * 10
RotateFont.lfFaceName = FontName & Chr$(0)
If obj.FontBold Then
    RotateFont.lfWeight = 800
Else
    RotateFont.lfWeight = 400
End If
RotateFont.lfHeight = (FontSize * -20) / Screen.TwipsPerPixelY
rFont = CreateFontIndirect(RotateFont)
CurFont = SelectObject(obj.hdc, rFont)

obj.CurrentX = X
obj.CurrentY = Y
obj.Print Caption

'Restore
foo = SelectObject(obj.hdc, CurFont)
foo = DeleteObject(rFont)

End Sub
Public Sub TextCircle(obj As Object, TXT As String, X As Long, Y As Long, radius As Long, startdegree As Double)
Dim foo As Integer, TXTX As Long, TXTY As Long, checkit As Integer
Dim twipsperdegree As Long, wrktxt As String, wrklet As String, degreexy As Double, degree As Double
twipsperdegree = (radius * 3.14159 * 2) / 360
If startdegree < 0 Then
    Select Case startdegree
    Case -1
        startdegree = Int(360 - (((obj.TextWidth(TXT)) / twipsperdegree) / 2))
    Case -2
        radius = (obj.TextWidth(TXT) / 2) / 3.14159
        twipsperdegree = (radius * 3.14159 * 2) / 360
    End Select
End If


For foo = 1 To Len(TXT)
    wrklet = Mid$(TXT, foo, 1)
    degreexy = (obj.TextWidth(wrktxt)) / twipsperdegree + startdegree
    DegreesToXY X, Y, degreexy, radius, radius, TXTX, TXTY
    degree = (obj.TextWidth(wrktxt) + 0.5 * obj.TextWidth(wrklet)) / twipsperdegree + startdegree
    RotateText 360 - degree, obj, obj.FontName, obj.FontSize, (TXTX), (TXTY), wrklet
    wrktxt = wrktxt & wrklet
Next foo
End Sub

Sub 旋转文本(Picture1 As Object, style As Integer, Text As String)
On Error Resume Next
Select Case style
Case 0 'center on top: degree = -1
    Picture1.FontName = "arial"
    Picture1.FontSize = 40
    Picture1.FontBold = True
    TextCircle Picture1, Text, Picture1.ScaleWidth / 2, Picture1.ScaleHeight, Picture1.ScaleHeight * 0.8, -1
Case 1 'adjust circle size to fit text length: degree = -2
    Picture1.FontName = "arial"
    Picture1.FontSize = 12
    Picture1.FontBold = True
    TextCircle Picture1, Text, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, Picture1.ScaleHeight * 0.3, -2
Case 2 'start at point: degree = 0 to 360
    Picture1.FontName = "arial"
    Picture1.FontSize = 12
    Picture1.FontBold = True
    TextCircle Picture1, Text, Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2, Picture1.ScaleHeight * 0.5, 90
Case 3
Dim foo As Integer
For foo = 0 To 360 Step 45
  Picture1.Refresh
  RotateText foo, Picture1, "Arial", 24, 2400, 2400, Text
  DoEvents
Next foo
Case 4
Picture1.FontName = "arial"
Picture1.FontSize = 8
For foo = 0 To 3
RotateText 270, Picture1, "Arial", 8, Picture1.ScaleWidth, foo * Picture1.TextWidth(Text), Text
Next foo
End Select
End Sub
Public Function sNT(ByVal sStr As String) As String
  Dim iNL As Integer
  
  iNL = InStr(sStr, Chr(0))
  If iNL > 0 Then
    sNT = Left(sStr, iNL - 1)
  Else
    sNT = sStr
  End If
End Function
'GBK编码函数
Function GBKEncode(szInput As String) As String
    Dim I As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim X() As Byte
    
    X = StrConv(szInput, vbFromUnicode)
    
    startIndex = LBound(X)
    endIndex = UBound(X)
    For I = startIndex To endIndex
        GBKEncode = GBKEncode & "%" & Hex(X(I))
    Next
End Function

'网络搜歌
Public Function FindLic(ByVal SINGER As String, ByVal Music As String) As String
Dim IStart As Long, IEnd As Long, strCode As String, strlrc As String
strlrc = "http://www.cnlyric.com/search.php.k=" + GBKEncode(SINGER) + " " + GBKEncode(Music) + "&t=s"
strCode = ReadinteFile(strlrc)
IStart = InStr(1, strCode, "LrcDown")
If IStart <> 0 Then
IEnd = InStr(IStart, strCode, "lrc")
FindLic = "http://www.cnlyric.com/" + Mid$(strCode, IStart, IEnd - IStart + 3)
Else
FindLic = ""
FRMLRC.LBLRC.SETTXT "没有找到歌词"
End If

End Function

'将汉字转化为百度URL编码
 Public Function UTF8EncodeURI(szInput)
        Dim wch, uch, szRet
        Dim X
        Dim nAsc, nAsc2, nAsc3

        If szInput = "" Then
            UTF8EncodeURI = szInput
            Exit Function
        End If

        For X = 1 To Len(szInput)
            wch = Mid(szInput, X, 1)
            nAsc = AscW(wch)

            If nAsc < 0 Then nAsc = nAsc + 65536

            If (nAsc And &HFF80) = 0 Then
                szRet = szRet & wch
            Else
                If (nAsc And &HF000) = 0 Then
                    uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                    szRet = szRet & uch
                Else
                    uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                    Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                    Hex(nAsc And &H3F Or &H80)
                    szRet = szRet & uch
                End If
            End If
        Next

        UTF8EncodeURI = szRet
    End Function
Public Function ReadinteFile(ByVal sURL As String) As String
Dim xmlHTTP1 As Object
Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
xmlHTTP1.Open "get", sURL, True
xmlHTTP1.Send
While xmlHTTP1.ReadyState <> 4
DoEvents
Wend
ReadinteFile = xmlHTTP1.responseText
Set xmlHTTP1 = Nothing
End Function
Public Function FindMp3URL(ByVal MusicName As String, ByVal Artic As String) As String
On Error Resume Next
Dim IStart As Long, IEnd As Long, strCode As String, sTemp As String, strURL As String
strURL = "http://box.zhangmen.baidu.com/x.op=12&count=1&title=" & MusicName & "$$" & Artic & "$$$$"
strCode = ReadinteFile(strURL)
If Len(strCode) > 100 Then
IStart = InStr(1, strCode, "<encode>")
IEnd = InStr(1, strCode, "</encode>")
sTemp = Mid$(strCode, IStart + 17, IEnd - IStart - 20)
IStart = InStr(1, strCode, "<decode>")
IEnd = InStr(1, strCode, "</decode>")
FindMp3URL = sTemp & "/" & Mid$(strCode, IStart + 17, IEnd - IStart - 20)
Else
If IS_CHK_LIST = False Then FindMp3URL = "": Exit Function ' Call SHOWWRONG("歌曲不存在或歌手名或歌名有误!", 0)
End If
If strCode = "" Then
Call SHOWWRONG("网络未连接,请稍后重试!", 0)
FindMp3URL = ""
End If
End Function

Public Sub CalculateEntry()
'On Error GoTo ErrorHandler:
Dim Answer As String
Dim BinAnswer As String
Dim DecimalCheck As Long
Dim I As Integer
Dim LenAfterDecimal As Long
Dim NumOfDecimals As Integer
Dim Remainder As String
Dim Tag As String

    'Set default values
    CurrentEntryIndex = 1
    Help = False
    InError = False
    InputString = frmma.txtEntry.Text
    PrevEntry = frmma.txtEntry.Text
    SetVariable = False

    'Extract the first token
    ExtractToken

    'Evaluate the entire expression
    Answer = CStr(GetE)

    'If we "finished" the evaluation prematurely, an
    'error occured
    If Not InError And OutputString <> "EOS" Then
        TrapErrors 0
    End If

    'Set error message if error occurred
    If InError Then
        Answer = ">> " + ErrorMessage + vbNewLine + frmma.txtAnswer.Text

    Else

        'Set previous answer
        PrevAnswer = Answer
        Tag = ""
        If frmma.optBaseMode(1).Value = True Then

            'Convert to binary if necessary
            If CDbl(Answer) <= 32767 Then
                BinAnswer = ""
                DecimalCheck = InStr(1, CStr(Answer), ".")
                If DecimalCheck <> 0 Then
                    If CInt(Mid(CStr(Answer), DecimalCheck + 1, 1)) < 5 Then
                        Answer = CDbl(Left(Answer, DecimalCheck - 1))
                    Else
                        Answer = CDbl(Left(Answer, DecimalCheck - 1)) + 1
                    End If
                End If
                Do
                    Answer = Answer / 2
                    DecimalCheck = InStr(1, CStr(Answer), ".")
                    If DecimalCheck = 0 Then
                        Remainder = "0"
                    Else
                        
                        Answer = CDbl(Left(Answer, DecimalCheck - 1))
                        Remainder = "1"
                    End If
                    BinAnswer = Remainder + BinAnswer
                Loop Until Answer < 1
                Answer = CDbl(BinAnswer)
                Tag = " (bin)"
            End If
        ElseIf frmma.optBaseMode(2).Value = True Then

            'Convert to hexadecimal if necessary
            Answer = Hex(Answer)
            Tag = " (hex)"
        ElseIf frmma.optBaseMode(3).Value = True Then

            'Convert to octadecimal if necessary
            Answer = Oct(Answer)
            Tag = " (oct)"
        Else

            'If in decimal mode, convert to set
            'number of decimal places
            If frmma.txtDecimal.Text <> "F" Then

                'Check for decimal
                NumOfDecimals = Val(frmma.txtDecimal.Text)
                DecimalCheck = InStr(1, CStr(Answer), ".")

                'If decimal does not exist, tag on the number
                'of zeroes that the user specified
                If DecimalCheck = 0 Then
                    If NumOfDecimals <> "0" Then
                        Answer = Answer + "."
                        For I = 1 To NumOfDecimals
                            Answer = Answer + "0"
                        Next I
                    End If

                'If decimal does exist, adjust the answer to
                'the number of decimal places that the user
                'specified
                Else
                    LenAfterDecimal = Len(Answer) - DecimalCheck
                    If LenAfterDecimal > NumOfDecimals Then
                        If NumOfDecimals = "0" Then
                            DecimalCheck = DecimalCheck - 1
                        End If
                        Answer = Mid(Answer, 1, DecimalCheck + NumOfDecimals)
                    Else
                        For I = 1 To (NumOfDecimals - LenAfterDecimal)
                            Answer = Answer + "0"
                        Next I
                    End If
                End If
            End If
        End If
        Answer = ">> " + Answer + Tag + vbNewLine + frmma.txtAnswer.Text
    End If

    'Display final answer
    frmma.txtAnswer.Text = Answer

    Exit Sub

ErrorHandler:

    'Trap errors
    TrapErrors ERR.Number

End Sub


Public Sub ExtractToken()
Dim I As Integer
    OutputString = ""
    OutputValue = 0
    ValueString = ""
    If CurrentEntryIndex > Len(InputString) Then
        OutputString = "EOS"
        Exit Sub
    End If
    Char = Mid(InputString, CurrentEntryIndex, 1)
    If Char = " " Then
        CurrentEntryIndex = CurrentEntryIndex + 1
        ExtractToken
        Exit Sub
    End If
    If Char = "+" Or Char = "-" Or Char = "*" Or Char = "/" Or Char = "^" Or Char = "(" Or Char = ")" Or Char = "!" Or Char = "=" Then
        CurrentEntryIndex = CurrentEntryIndex + 1
        OutputString = Char
        Exit Sub
    End If
    If (Char >= "0" And Char <= "9") Or Char = "." Then
        While Char >= "0" And Char <= "9"
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Decimal
        While Char = "."
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Digits after decimal
        While Char >= "0" And Char <= "9"
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Set return values
        OutputString = "Number"
        OutputValue = CDbl(ValueString)
        Exit Sub
    End If

    'Return text language identifiers
    If LCase(Char) >= "a" And LCase(Char) <= "z" Then
        While (LCase(Char) >= "a" And LCase(Char) <= "z")
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Pi or e
        If LCase(ValueString) = "pi" Or LCase(ValueString) = "e" Then
            OutputString = "Number"
            If LCase(ValueString) = "pi" Then
                OutputValue = PI
            Else
                OutputValue = Exp(1)
            End If
            Exit Sub
        End If

        'Set return value
        OutputString = LCase(ValueString)
        Exit Sub
    End If

End Sub

Public Function GetE()
On Error GoTo ErrorHandler

    '**********************************
    '* PARSING ROUTINE (Expression E) *
    '* E ::= T + T | T - T | T        *
    '**********************************

    'Get the lower value (T)
    Value = GetT()

    'Exit function if error or help call returned
    If InError Or Help Then
        Exit Function
    End If

    'User set a value to a variable
    If SetVariable Then
        GetE = Value
        Exit Function
    End If

    'Allow for multiple operators of the same precedence
    'level occuring immediately after each other
    While OutputString = "+" Or OutputString = "-"

        Select Case OutputString
    
            'Addition operator
            Case "+"
                ExtractToken
                Value = Value + GetT()
    
            'Subraction operator
            Case "-"
                ExtractToken
                Value = Value - GetT()

        End Select

    Wend

    'Return value for E
    GetE = Value

    'Exit function before error handler
    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors ERR.Number

End Function

Public Function GetT()
On Error GoTo ErrorHandler
Dim Exponent As Double

    '****************************
    '* PARSING ROUTINE (Term T) *
    '* T ::= F * F | F / F | F  *
    '****************************

    'Get the lower value (F)
    Value = GetF

    'Exit function if error or help call returned
    If InError Or Help Then
        Exit Function
    End If

    'User set a value to a variable
    If SetVariable Then
        GetT = Value
        Exit Function
    End If

    'Allow for multiple operators of the same precedence
    'level occuring immediately after each other
    While OutputString = "*" Or OutputString = "/"

        Select Case OutputString
    
            'Multiplication operator
            Case "*"
                ExtractToken
                Value = Value * GetF()
    
            'Division operator
            Case "/"
                ExtractToken
                Value = Value / GetF()
    
        End Select

    Wend

    'Return value for T
    GetT = Value

    'Exit function before error handler
    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors ERR.Number

End Function

Public Function GetF()
On Error GoTo ErrorHandler
Dim ArrayIndex As Long
Dim ArrayItemExists As Boolean
Dim ArrayString As String
Dim ArrayValue As Double
Dim base As Double
Dim Constant As Double
Dim ConstantExists As Boolean
Dim FileNumber As Long
Dim LogBase As String
Dim LogIndex As Long
Dim I As Long
Dim temp As String
Dim Temp2 As String
    Select Case OutputString
        Case "Number"
            Value = OutputValue
            ExtractToken
            GetF = PostToken
        Case "-"
            ExtractToken
            GetF = -(GetF())
        Case "rnd"
            Randomize
            Value = Rnd
            ExtractToken
            GetF = PostToken

        'Parenthesis
        Case "("
            ExtractToken
            Value = GetE
            If OutputString <> ")" And OutputString <> "EOS" Then
                TrapErrors 0
                Exit Function
            End If
            If OutputString = "EOS" Then
                GetF = Value
            Else
                ExtractToken
                GetF = PostToken
            End If
        Case "ans"
            Value = PrevAnswer
            ExtractToken
            GetF = PostToken
        Case "abs"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                GetF = Abs(Value)
            End If

        'Help
        Case "help"
            Help = True

        'Square Root
        Case "sr"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                GetF = Sqr(Value)
            End If
        Case "log"

            'Get logarithm base
            LogBase = frmma.txtLogBase.Text

            'If the box is empty, set it with the default 10
            If LogBase = "" Then
                frmma.txtLogBase.Text = "10"
                base = 10

            'Retrieve logarithm base
            Else
                base = Val(LogBase)
            End If

            'Get number
            ExtractToken
            GetF = Log(GetF()) / Log(base)

        'Natural logarithm
        Case "ln"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                GetF = Log(Value)
            End If

        'Cosine
        Case "cos"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = Cos(Value)
            End If

        'Cotangent
        Case "cot"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 1 / Tan(Value)
            End If

        'Cosecant
        Case "csc"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 1 / SIN(Value)
            End If

        'Hyperbolic cosecant
        Case "hcsc"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 2 / (Exp(Value) - Exp(-Value))
            End If
            Exit Function

        'Hyperbolic cosine
        Case "hcos"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) + Exp(-Value)) / 2
            End If

        'Hyperbolic cotangent
        Case "hcot"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) + Exp(-Value)) / (Exp(Value) - Exp(-Value))
            End If

        'Hyperbolic secant
        Case "hsec"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 2 / (Exp(Value) + Exp(-Value))
            End If

        'Hyperbolic sine
        Case "hsin"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) - Exp(-Value)) / 2
            End If

        'Hyperbolic tangent
        Case "htan"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) - Exp(-Value)) / (Exp(Value) + Exp(-Value))
            End If

        'Inverse hyperbolic cosine
        Case "ihcos"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log(Value + Sqr(Value * Value - 1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic cosecant
        Case "ihcsc"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((Sgn(Value) * Sqr(Value * Value + 1) + 1) / Value)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic cotangent
        Case "ihcot"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((Value + 1) / (Value - 1)) / 2
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic sine
        Case "ihsin"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log(Value + Sqr(Value * Value + 1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic secant
        Case "ihsec"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((Sqr(-Value * Value + 1) + 1) / Value)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic tangent
        Case "ihtan"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((1 + Value) / (1 - Value)) / 2
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse cosecant
        Case "icsc"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(Value * Value - 1)) + (Sgn(Value) - 1) * (2 * Atn(1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse cosine
        Case "icos"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse cotangent
        Case "icot"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value) + 2 * Atn(1)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse secant
        Case "isec"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(Value * Value - 1)) + Sgn((Value) - 1) * (2 * Atn(1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse sine
        Case "isin"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(-Value * Value + 1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse tangent
        Case "itan"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value)
                ConvertToDegrees
                GetF = Value
            End If

        'Secant
        Case "sec"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 1 / Cos(Value)
            End If

        'Sine
        Case "sin"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = SIN(Value)
            End If

        'Tangent
        Case "tan"
            ExtractToken
            Value = GetF()
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = Tan(Value)
            End If

        'Equals sign
        Case "="

            'A string that starts with an equals sign is not
            'a reserved keyword
            If CurrentEntryIndex = 2 Then
                TrapErrors 0
            Else
                TrapErrors -10
            End If

        'Check for stored variables or variables that
        'should be stored - anything else is an error
        Case Else

            'Check to see if variable exists with the
            'entered name
            ArrayItemExists = False
            ArrayString = OutputString
            For I = LBound(VariableArray) To UBound(VariableArray)
                If VariableArray(I) = ArrayString Then
                    ArrayItemExists = True
                    ArrayIndex = I
                    Exit For
                End If
            Next I

            'Check to see if variable exists as a constant
            FileNumber = FreeFile
            Open "constants.csd" For Input As #FileNumber
                Do While Not EOF(FileNumber)
                    Input #FileNumber, temp
                    Input #FileNumber, Temp2
                    If temp = ArrayString Then
                        ConstantExists = True
                        Constant = Val(Temp2)
                        Exit Do
                    End If
                Loop
            Close #FileNumber

            'Check to see if the user wishes to store a
            'variable
            ExtractToken
            If OutputString = "=" Then

                'Get the value for the string to be stored
                ExtractToken
                ArrayValue = GetE()
                If InError Then
                    Exit Function
                End If

                If ArrayItemExists Then

                    'Replace the variable's value only,
                    'instead of created a whole new array
                    'item
                    ValueArray(ArrayIndex) = ArrayValue

                Else

                    'User cannot assign a value to a preset
                    'constant value
                    If ConstantExists Then
                        TrapErrors -10
                        Exit Function
                    End If

                    'Give an extra space for the new variable
                    ReDim Preserve ValueArray(UBound(ValueArray) + 1)
                    ReDim Preserve VariableArray(UBound(VariableArray) + 1)

                    'Store the new variable in the array
                    ValueArray(UBound(ValueArray)) = ArrayValue
                    VariableArray(UBound(VariableArray)) = ArrayString

                End If

                'Displayed answer is the stored value
                SetVariable = True
                GetF = ArrayValue
            Else

                'If the user has stored a variable with the
                'entered name, return the stored value
                If ArrayItemExists Then
                    Value = ValueArray(ArrayIndex)
                    GetF = PostToken
                    Exit Function
                End If

                'If the user has stored a constant with the
                'entered name, return the stored value
                If ConstantExists Then
                    Value = Constant
                    GetF = PostToken
                    Exit Function
                End If

                'If variable does not exist, then error
                TrapErrors 0
            End If

    End Select

    'Exit function before error handler
    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors ERR.Number

End Function


Public Sub TrapErrors(ErrNumber As Long)

    'Set trapped error message
    If ErrNumber = -10 Then

        'Reserved keyword
        ErrorMessage = "错误: 保留关键字"

    ElseIf ErrNumber = 6 Then

        'Overflow
        ErrorMessage = "错误: 溢出"

    ElseIf ErrNumber = 11 Then

        'Division By Zero
        ErrorMessage = "错误: 除数不能为零"

    Else

        'Unknown error
        ErrorMessage = "表达式错误"

    End If

    'Set return values
    InError = True
    OutputString = "TError"

End Sub

Public Sub ConvertToDegrees()

    'Convert to degrees
    If frmma.optAngMode(0).Value = True Then
        Value = Value * (180 / PI)
    End If

End Sub

Public Sub ConvertToRadians()

    'Convert to radians
    If frmma.optAngMode(0).Value = True Then
        Value = Value * (PI / 180)
    End If

End Sub

Public Function PostToken()
On Error GoTo ErrorHandler
Dim Factorial As Double
Dim I As Integer
Dim Value2 As Double

    'Ignore operators, EOS strings, right parentheses, and
    'equals signs
    If OutputString = "+" Or OutputString = "-" Or OutputString = "*" Or OutputString = "/" Or OutputString = "EOS" Or OutputString = ")" Or OutputString = "=" Or OutputString = "," Then
        PostToken = Value

    'Handle special tokens that come after the value
    Else
        Select Case OutputString

            'Factorial
            Case "!"
                If (CDbl(Value) <> CLng(Value)) Or Value < 0 Then
                    TrapErrors 0
                    Exit Function
                End If
                Factorial = 1
                For I = Value To 1 Step -1
                    Factorial = Factorial * I
                Next I
                ExtractToken

                'Ignore operators, EOS strings, right
                'parentheses, and equals signs
                If OutputString = "+" Or OutputString = "-" Or OutputString = "*" Or OutputString = "/" Or OutputString = "EOS" Or OutputString = ")" Or OutputString = "=" Then
                    PostToken = Factorial
                    ExtractToken

                'Handle special tokens that come after a
                'factorial
                Else

                    Select Case OutputString

                        'Factorial
                        Case "!"
                            TrapErrors 0
                            Exit Function

                        'Exponent
                        Case "^"
                            ExtractToken
                            PostToken = Factorial ^ GetF

                        'Other "post" tokens multiply
                        Case Else
                            PostToken = Factorial * GetF
                    End Select
                End If

            'Exponent
            Case "^"
                ExtractToken
                PostToken = Value ^ GetF

            'Left parenthesis
            Case "("
                PostToken = Value * GetF

            'Other "post" tokens multiply
            Case Else
                PostToken = Value * GetF
        End Select
    End If

    Exit Function

ErrorHandler:

    TrapErrors ERR.Number

End Function
Public Function FileDel(Str1 As String) As Long
    Dim Result As Long, fileop As SHFILEOPSTRUCT
    With fileop
        .hwnd = 0
        .wFunc = FO_DELETE
        .pFrom = Str1 & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
    Result = SHFileOperation(fileop)
End Function
Public Sub UnHook()
 Dim temp As Long
 temp = SetWindowLong(gHW, GWL_WNDPROCB, lpPrevWndProc)
End Sub

Public Sub Hook()
lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROCB, AddressOf WindowProc)
End Sub
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
 If uMsg = WM_MOUSEWHEEL Then
ProcMouseWheel wParam, lParam
 Else
WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
 End If
End Function

Public Function GetCPUTemp() As Double '获得CPU温度
  Dim I As Integer
  Dim mCPU As Variant
  Dim U As Variant
  Dim s As String
  Set mCPU = GetObject("WINMGMTS:{impersonationLevel=impersonate}!root\wmi").ExecQuery("SELECT   CurrentTemperature   From   MSAcpi_ThermalZoneTemperature")
  For Each U In mCPU
  s = s & U.CurrentTemperature
  Next
  Set mCPU = Nothing
  GetCPUTemp = (s - 2732) / 10
End Function


Public Function GetString(hKey As Long, strpath As String, strValue As String)

Dim Keyhand As Long
Dim lValueType As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(hKey, strpath, Keyhand)
lResult = RegQueryValueEx(Keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(Keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If
End Function


Public Sub GetKeyNames(ByVal hKey As Long, ByVal strpath As String)
Dim Cnt As Long, StrBuff As String, StrKey As String, TKey As Long
    RegOpenKey hKey, strpath, TKey
    Do
        StrBuff = String(255, vbNullChar)
        If RegEnumKeyEx(TKey, Cnt, StrBuff, 255, 0, vbNullString, 0, ByVal 0&) <> 0 Then Exit Do
        Cnt = Cnt + 1
        StrKey = Left(StrBuff, InStr(StrBuff, vbNullChar) - 1)
        sKeys.Add StrKey
    Loop
End Sub


Public Sub ProcMouseWheel(wParam As Long, lParam As Long)
On Error Resume Next
Dim fwKeys As Long
Dim zDelta As Long
Dim XPos As Long
Dim YPos As Long
Dim Shift16 As Long
Dim lIdx As Long
Shift16 = 65536
If wParam < 0 Then
zDelta = ((CLng(wParam) And &HFFFF0000) \ Shift16) And &HFFFF&
'注: 第二个&一定要加
zDelta = zDelta - Shift16
Else
zDelta = ((CLng(wParam) And &HFFFF0000) \ Shift16) And &HFFFF&
End If
fwKeys = (CLng(wParam) And &HFFFF&)
YPos = ((CLng(lParam) And &HFFFF0000) \ Shift16) And &HFFFF&
XPos = (CLng(lParam) And &HFFFF&)
If H_DOS = 0 Then
frmGraphic.vsbSlide.Value = frmGraphic.vsbSlide.Value - 0.5 * zDelta
ElseIf H_DOS = 1 Then

ElseIf H_DOS = 2 Then
Exit Sub
ElseIf H_DOS = 3 Then
FRMBOARD.SCRO.Value = FRMBOARD.SCRO.Value - 0.5 * zDelta
ElseIf H_DOS = 4 Then
frmset.SCRO.Value = frmset.SCRO.Value - 0.5 * zDelta
ElseIf H_DOS = 5 Then

ElseIf H_DOS = 6 Then
FRMPIC_MAN.vsbSlide.Value = FRMPIC_MAN.vsbSlide.Value - 0.5 * zDelta
ElseIf H_DOS = 7 Then
If MAINSTYLE = 3 And frmma.PP.Visible = False And frmma.PicUse.Left = 0 And frmma.PF(4).Visible = True Then frmma.SHRO.Value = frmma.SHRO.Value - 0.5 * zDelta
ElseIf H_DOS = 8 And FrmNetMusic.PO(2).Left = 0 And FrmNetMusic.PICFRAME.Visible = True Then
FrmNetMusic.vsbSlide.Value = FrmNetMusic.vsbSlide.Value - 0.5 * zDelta
End If
End Sub
Function existproc(ByVal exefile As String) As Boolean '检测进程
    existproc = False
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPall, 0&)
    uProcess.dwSize = Len(uProcess)
    Dim r As Long
    r = Process32First(hSnapShot, uProcess)
    Do While r
        If Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0)) = exefile Then
            existproc = True
            Exit Do
        End If
        r = Process32Next(hSnapShot, uProcess)
    Loop
End Function
'示例:existproc("ICEE.exe")=true表示ICEE已运行
Public Function RamUsage(Optional strProcess As String = "") As Double
    If strProcess = "" Then strProcess = UCase(App.exename) & ".EXE" 'Will count the current application as the process if no arguments given
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & strProcess & "'")
    For Each objProcess In colProcessList
           RamUsage = objProcess.WorkingSetSize / 1024
    Next
End Function
Public Function FormatUsage(tUsage As Double)
    If Int(tUsage) = tUsage Then
        If tUsage = 0 Then
            FormatUsage = 0
        Else
            FormatUsage = Format(tUsage, "###,###")
        End If
    Else
        FormatUsage = Format(tUsage, "###,###.#")
    End If
End Function

Public Function PFUsage(Optional strProcess As String = "") As Double
    If strProcess = "" Then strProcess = UCase(App.exename) & ".EXE" 'Will count the current application as the process if no arguments given
    
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & strProcess & "'")
    For Each objProcess In colProcessList
           PFUsage = objProcess.PagefileUsage / 1024
    Next
End Function
