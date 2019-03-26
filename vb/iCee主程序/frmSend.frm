VERSION 5.00
Begin VB.Form frmFileChoose 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00231C09&
   BorderStyle     =   0  'None
   Caption         =   "选择文件"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   FillColor       =   &H00383537&
   ForeColor       =   &H00383537&
   Icon            =   "frmSend.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox X2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5820
      Picture         =   "frmSend.frx":038A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   25
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5820
      Picture         =   "frmSend.frx":046E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   24
      Top             =   15
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox X1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   5820
      Picture         =   "frmSend.frx":0552
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   23
      Top             =   15
      Width           =   750
   End
   Begin VB.PictureBox PINFO 
      AutoRedraw      =   -1  'True
      BackColor       =   &H005C6105&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   3240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXE/DLL/OCX文件信息"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   120
         Width           =   1710
      End
      Begin VB.Shape DSB 
         BackColor       =   &H0030F1F1&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   1
         Left            =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.PictureBox PINFO 
      AutoRedraw      =   -1  'True
      BackColor       =   &H005C6105&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   7
      Top             =   4320
      Width           =   3135
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件信息"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   11
         Top             =   120
         Width           =   720
      End
      Begin VB.Shape DSB 
         BackColor       =   &H0030F1F1&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   0
         Left            =   0
         Top             =   360
         Width           =   3135
      End
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   4
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   480
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1920
      Width           =   5805
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   3375
      _ExtentX        =   6376
      _ExtentY        =   873
   End
   Begin ICEE.ICEE_COMMAND ICM 
      Height          =   495
      Index           =   2
      Left            =   3480
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
      _ExtentX        =   4260
      _ExtentY        =   873
   End
   Begin VB.PictureBox PO 
      BackColor       =   &H0047491F&
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   1
      Left            =   120
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   10
      Top             =   4800
      Width           =   6375
      Begin VB.PictureBox picLarge 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0047491F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   5640
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   21
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picSmall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0047491F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5400
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   20
         Top             =   960
         Width           =   240
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
         Left            =   2640
         TabIndex        =   22
         Top             =   360
         Width           =   360
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
         Left            =   2670
         TabIndex        =   19
         Top             =   600
         Width           =   1710
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
         TabIndex        =   18
         Top             =   600
         Width           =   810
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
         TabIndex        =   17
         Top             =   840
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
         Left            =   2670
         TabIndex        =   16
         Top             =   840
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
         TabIndex        =   15
         Top             =   1050
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
         Index           =   2
         Left            =   2670
         TabIndex        =   14
         Top             =   1050
         Width           =   1710
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
         TabIndex        =   13
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.PictureBox PO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00554513&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Index           =   0
      Left            =   120
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   425
      TabIndex        =   9
      Top             =   4800
      Width           =   6375
   End
   Begin VB.Image IA 
      Enabled         =   0   'False
      Height          =   240
      Index           =   4
      Left            =   240
      Picture         =   "frmSend.frx":0636
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择文件"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   720
   End
   Begin VB.Label LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "附加信息(最多200字)"
      ForeColor       =   &H00C0C0C0&
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1710
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   960
      Width           =   5175
   End
   Begin VB.Shape SB 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1935
      Index           =   1
      Left            =   120
      Top             =   1800
      Width           =   6375
   End
End
Attribute VB_Name = "frmFileChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Dim MYID        As Long
Dim SendClicked As Boolean
     Private filename As String
     Private Directory As String
     Private FullFileName As String

     Private StrucVer As String
     Private FileVer As String
     Private ProdVer As String
     Private FileFlags As String
     Private FileOS As String
     Private FileType As String
     Private FileSubType As String
    
       Private Type VS_NEWINFO
        astr As String * 1024
     End Type
     
       Private Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersionl As Integer
        dwStrucVersionh As Integer
        dwFileVersionMSl As Integer
        dwFileVersionMSh As Integer
        dwFileVersionLSl As Integer
        dwFileVersionLSh As Integer
        dwProductVersionMSl As Integer
        dwProductVersionMSh As Integer
        dwProductVersionLSl As Integer
        dwProductVersionLSh As Integer
        dwFileFlagsMask As Long
        dwFileFlags As Long
        dwFileOS As Long
        dwFileType As Long
        dwFileSubtype As Long
        dwFileDateMS As Long
        dwFileDateLS As Long
     End Type

      Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
        "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal _
        dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
      Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
        "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
        lpdwHandle As Long) As Long
       Private Declare Function VerQueryValue Lib "Version.dll" Alias _
        "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, _
        lplpBuffer As Any, puLen As Long) As Long
      Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Dest As Any, ByVal Source As Long, ByVal Length As Long)
       Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
        "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long
     

     Private Const VS_FFI_SIGNATURE = &HFEEF04BD
     Private Const VS_FFI_STRUCVERSION = &H10000
     Private Const VS_FFI_FILEFLAGSMASK = &H3F&


     Private Const VS_FF_DEBUG = &H1
     Private Const VS_FF_PRERELEASE = &H2
     Private Const VS_FF_PATCHED = &H4
     Private Const VS_FF_PRIVATEBUILD = &H8
     Private Const VS_FF_INFOINFERRED = &H10
     Private Const VS_FF_SPECIALBUILD = &H20


     Private Const VOS_UNKNOWN = &H0
     Private Const VOS_DOS = &H10000
     Private Const VOS_OS216 = &H20000
     Private Const VOS_OS232 = &H30000
     Private Const VOS_NT = &H40000

     Private Const VOS_BASE = &H0
     Private Const VOS_WINDOWS16 = &H1
     Private Const VOS_PM16 = &H2
     Private Const VOS_PM32 = &H3
     Private Const VOS_WINDOWS32 = &H4

     Private Const VOS_DOS_WINDOWS16 = &H10001
     Private Const VOS_DOS_WINDOWS32 = &H10004
     Private Const VOS_OS216_PM16 = &H20002
     Private Const VOS_OS232_PM32 = &H30003
     Private Const VOS_NT_WINDOWS32 = &H40004
     

     Private Const VFT_UNKNOWN = &H0
     Private Const VFT_APP = &H1
     Private Const VFT_DLL = &H2
     Private Const VFT_DRV = &H3
     Private Const VFT_FONT = &H4
     Private Const VFT_VXD = &H5
     Private Const VFT_STATIC_LIB = &H7


     Private Const VFT2_UNKNOWN = &H0
     Private Const VFT2_DRV_PRINTER = &H1
     Private Const VFT2_DRV_KEYBOARD = &H2
     Private Const VFT2_DRV_LANGUAGE = &H3
     Private Const VFT2_DRV_DISPLAY = &H4
     Private Const VFT2_DRV_MOUSE = &H5
     Private Const VFT2_DRV_NETWORK = &H6
     Private Const VFT2_DRV_SYSTEM = &H7
     Private Const VFT2_DRV_INSTALLABLE = &H8
     Private Const VFT2_DRV_SOUND = &H9
     Private Const VFT2_DRV_COMM = &HA

'这是一个获取文件信息的程序

Private Sub DisplayVerInfo()
        '*** 这个子程序获取文件的版本信息 ****
        Dim rc                As Long
        Dim lDummy            As Long
        Dim sBuffer()         As Byte
        Dim lBufferLen        As Long
        Dim lVerPointer       As Long
        Dim udtVerBuffer      As VS_FIXEDFILEINFO
        Dim lVerbufferLen     As Long
        Dim aBuffer()         As Byte
        Dim lAdd              As Long
        Dim astr              As String
        Dim lTran             As Long

        '*** Get size ****
        lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
        If lBufferLen < 1 Then Call SHOWWRONG("无法获取文件版本信息!", 2): Exit Sub
        '**** 获取文件信息并且保存到udtVerBuffer结构中 ****
        ReDim sBuffer(lBufferLen)
        rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
        rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
        MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
                
        StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & _
           Format$(udtVerBuffer.dwStrucVersionl)

        '**** 获得文件版本 ****
        FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & _
           Format$(udtVerBuffer.dwFileVersionMSl) & "." & _
           Format$(udtVerBuffer.dwFileVersionLSh) & "." & _
           Format$(udtVerBuffer.dwFileVersionLSl)

        '**** 获取产品版本 ****
        ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & _
           Format$(udtVerBuffer.dwProductVersionMSl) & "." & _
           Format$(udtVerBuffer.dwProductVersionLSh) & "." & _
           Format$(udtVerBuffer.dwProductVersionLSl)

        '**** 获取文件类型 ****
        FileFlags = ""
        If udtVerBuffer.dwFileFlags And VS_FF_DEBUG _
           Then FileFlags = "Debug "
        If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE _
           Then FileFlags = FileFlags & "PreRel "
        If udtVerBuffer.dwFileFlags And VS_FF_PATCHED _
           Then FileFlags = FileFlags & "Patched "
        If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD _
           Then FileFlags = FileFlags & "Private "
        If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRED _
           Then FileFlags = FileFlags & "Info "
        If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD _
           Then FileFlags = FileFlags & "Special "
        If udtVerBuffer.dwFileFlags And VFT2_UNKNOWN _
           Then FileFlags = FileFlags + "Unknown "

        '**** 获取文件所适应的操作系统 ****
        Select Case udtVerBuffer.dwFileOS
           Case VOS_WINDOWS32
             FileOS = "Win32位操作系统"
           Case VOS_WINDOWS16
             FileOS = "Win16位操作系统"
           Case VOS_DOS
             FileOS = "DOS操作系统"
           Case VOS_DOS_WINDOWS16
             FileOS = "DOS-Win16操作系统"
           Case VOS_DOS_WINDOWS32
             FileOS = "DOS-Win32操作系统"
           Case VOS_OS216_PM16
             FileOS = "OS/2-16 PM-16操作系统"
           Case VOS_OS232_PM32
             FileOS = "OS/2-16 PM-32操作系统"
           Case VOS_NT_WINDOWS32
             FileOS = "NT-Win32操作系统"
           Case Else
             FileOS = "未知操作系统"
        End Select
        Select Case udtVerBuffer.dwFileType
           Case VFT_APP
              FileType = "应用程序"
           Case VFT_DLL
              FileType = "动态连接库"
           Case VFT_DRV
              FileType = "驱动程序"
              Select Case udtVerBuffer.dwFileSubtype
                 Case VFT2_DRV_PRINTER
                    FileSubType = "打印驱动程序"
                 Case VFT2_DRV_KEYBOARD
                    FileSubType = "键盘驱动程序"
                 Case VFT2_DRV_LANGUAGE
                    FileSubType = "语言模块"
                 Case VFT2_DRV_DISPLAY
                    FileSubType = "显示驱动程序"
                 Case VFT2_DRV_MOUSE
                    FileSubType = "鼠标驱动程序"
                 Case VFT2_DRV_NETWORK
                    FileSubType = "网络驱动程序"
                 Case VFT2_DRV_SYSTEM
                    FileSubType = "系统驱动程序"
                 Case VFT2_DRV_INSTALLABLE
                    FileSubType = "Installable"
                 Case VFT2_DRV_SOUND
                    FileSubType = "声音驱动程序"
                 Case VFT2_DRV_COMM
                    FileSubType = "串行驱动程序"
                 Case VFT2_UNKNOWN
                    FileSubType = "未知驱动程序"
              End Select
           'Case VFT_FONT
              'FileType = "字体"
             ' Select Case udtVerBuffer.dwFileSubtype
                 'Case VFT_FONT_RASTER
                  '  FileSubType = "光栅字体"
              '   Case VFT_FONT_VECTOR
              '      FileSubType = "矢量字体"
              '   Case VFT_FONT_TRUETYPE
              '      FileSubType = "TrueType字体"
              'End Select
           Case VFT_VXD
              FileType = "VxD"
           Case VFT_STATIC_LIB
              FileType = "Lib"
           Case Else
              FileType = "未知"
        End Select
        PO(0).CurrentX = 4
        PO(0).CurrentY = 4
        PO(0).Print "文件全路径:"
        PO(0).CurrentX = 4
        PO(0).Print "文件版本:"
        PO(0).CurrentX = 4
        PO(0).Print "产品版本:"
        PO(0).CurrentX = 4
        PO(0).Print "文件标志:"
        PO(0).CurrentX = 4
        PO(0).Print "操作系统:"
        PO(0).CurrentX = 4
        PO(0).Print "文件类型:"
        PO(0).CurrentX = 4
        PO(0).Print "文件子类型:"
        PO(0).CurrentX = 60
        PO(0).CurrentY = 4
        PO(0).Print FullFileName
        PO(0).CurrentX = 60
        PO(0).Print FileVer
        PO(0).CurrentX = 60
        PO(0).Print ProdVer
        PO(0).CurrentX = 60
        PO(0).Print FileFlags
        PO(0).CurrentX = 60
        PO(0).Print FileOS
        PO(0).CurrentX = 60
        PO(0).Print FileType
        PO(0).CurrentX = 60
        PO(0).Print FileSubType
        '清除上一次保存的信息
        FullFileName = ""
        FileVer = ""
        ProdVer = ""
        FileFlags = ""
        FileOS = ""
        FileType = ""
        FileSubType = ""
        
        
        ReDim aBuffer(lBufferLen)
        Dim ab As VS_NEWINFO
        
        lVerPointer = 0
        rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
        rc = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lVerbufferLen)
        MoveMemory lTran, lVerPointer, 4&
        astr = "0" + Hex$(lTran)
        astr = Right$(astr, 4) + Left$(astr, 4)
        rc = VerQueryValue(sBuffer(0), "\StringFileInfo\" + astr + "\FileDescription", lVerPointer, lVerbufferLen)
        MoveMemory ab, lVerPointer, Len(ab)
        PO(0).CurrentX = 4
        PO(0).Print "文件描述";
        PO(0).CurrentX = 60
        PO(0).Print Left$(ab.astr, (InStr(ab.astr, Chr$(0)) - 1))
        
        rc = VerQueryValue(sBuffer(0), "\StringFileInfo\" + astr + "\ProductName", lVerPointer, lVerbufferLen)
        If rc Then
          MoveMemory ab, lVerPointer, Len(ab)
          PO(0).CurrentX = 4
          PO(0).Print "产品名称";
          PO(0).CurrentX = 60
          PO(0).Print Left$(ab.astr, (InStr(ab.astr, Chr$(0)) - 1))
        End If
        
        rc = VerQueryValue(sBuffer(0), "\StringFileInfo\" + astr + "\OriginalFilename", lVerPointer, lVerbufferLen)
        If rc Then
          MoveMemory ab, lVerPointer, Len(ab)
          PO(0).CurrentX = 4
          PO(0).Print "文件原始名";
          PO(0).CurrentX = 60
          PO(0).Print Left$(ab.astr, (InStr(ab.astr, Chr$(0)) - 1))
        End If
        
        rc = VerQueryValue(sBuffer(0), "\StringFileInfo\" + astr + "\InternalName", lVerPointer, lVerbufferLen)
        If rc Then
          MoveMemory ab, lVerPointer, Len(ab)
          PO(0).CurrentX = 4
          PO(0).Print "文件内部名";
          PO(0).CurrentX = 60
          PO(0).Print Left$(ab.astr, (InStr(ab.astr, Chr$(0)) - 1))
        End If
        
        rc = VerQueryValue(sBuffer(0), "\StringFileInfo\" + astr + "\CompanyName", lVerPointer, lVerbufferLen)
        If rc Then
          MoveMemory ab, lVerPointer, Len(ab)
          PO(0).CurrentX = 4
          PO(0).Print "公司名称";
          PO(0).CurrentX = 60
          PO(0).Print Left$(ab.astr, (InStr(ab.astr, Chr$(0)) - 1))
        End If
        
        rc = VerQueryValue(sBuffer(0), "\StringFileInfo\" + astr + "\LegalCopyright", lVerPointer, lVerbufferLen)
        If rc Then
          MoveMemory ab, lVerPointer, Len(ab)
          PO(0).CurrentX = 4
          PO(0).Print "版权所有";
          PO(0).CurrentX = 100
          PO(0).Print Left$(ab.astr, (InStr(ab.astr, Chr$(0)) - 1))
        End If
End Sub

Private Sub Form_Load()
Me.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), COLOR_HIGH, B
Call PaintPng(App.Path & "\SKIN\FC_T.PNG", Me.hdc, 8, 8)
ICM(0).HASLINE = False
ICM(1).HASLINE = False
ICM(2).HASLINE = False
Me.Move frmma.Left + (frmma.Width - Me.Width) / 2, frmma.Top + (frmma.Height - Me.Height) / 2
ICM(0).SETTXT "浏    览"
ICM(1).SETTXT "发    送"
ICM(2).SETTXT "取    消"

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strpath As String
If Data.files.Count > 0 Then
strpath = Data.files(1)
txtFile.Text = strpath
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If (SendClicked = False) Then
    Set frmFileChoose = Nothing
    Set ftSend(MYID).frmChoose = Nothing
  End If
End Sub
Private Sub ICM_Click(Index As Integer)
Select Case Index
Case 0
On Error GoTo Err_DetermineErr
txtFile = ShowOpen(Me.hwnd, "所有文件" & Chr$(0) & "*.*", "选择文件")
FullFileName = UCase(txtFile.Text)
If Right(FullFileName, 3) = "EXE" Or Right(FullFileName, 3) = "DLL" Or Right(FullFileName, 3) = "OCX" Then PINFO(1).Visible = True: Call DisplayVerInfo Else PINFO(1).Visible = False
Exit Sub
Err_DetermineErr:
Case 1
If Dir(txtFile.Text) = "" Then Call SHOWWRONG("文件不存在!", 0): txtFile.Text = "": Exit Sub
  With ftSend(MYID)
    .Comment = txtComments
    .FileSize = CDbl(FileLen(txtFile))
    .FileToSend = txtFile
    .frmSend.InitTransfer MYID
  End With
  SendClicked = True
  Unload Me
Case 2
  Unload Me
End Select
End Sub

Private Sub LA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub LB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call PINFO_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub PINFO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
DSB(1).Visible = False
DSB(0).Visible = True
PO(1).Visible = True
PO(0).Visible = False
Case 1
DSB(1).Visible = True
DSB(0).Visible = False
PO(1).Visible = False
PO(0).Visible = True
End Select

End Sub

Private Sub PO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call CMV(Me)
End Sub

Private Sub PO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = True
X2.Visible = False
End Sub

Private Sub txtFile_Change()
  On Error GoTo ErrHandler
  Call Me.updatestats(txtFile.Text)
  Exit Sub
ErrHandler:
End Sub

Public Function ChooseSend(ByVal id As Long)
  MYID = id
  Me.Visible = True
End Function

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub x1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1.Visible = False
X2.Visible = True
End Sub
Private Sub x2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
X2.Visible = False
X3.Visible = True
End If
End Sub
Private Sub x3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X3.Visible = False
X1.Visible = True
If X3.Visible = False Then Unload Me
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

