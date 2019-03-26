VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'É¨Ãè¾ÖÓòÍø
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Long, ReplyBuffer As Icmp_Echo_Reply, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, addrlen As Long, addrType As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal ipaddress$) As Long
Private Declare Function gethostbyname& Lib "wsock32.dll" (ByVal hostname$)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const AF_INET As Integer = 2
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const PING_TIMEOUT = 150

Private Type Inet_address
  Byte4 As Byte
  Byte3 As Byte
  Byte2 As Byte
  Byte1 As Byte
End Type


Private Type Icmp_Options
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Private Type Icmp_Echo_Reply
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As Icmp_Options
    Data            As String * 250
End Type

Private IPLong As Inet_address
Private IcmPopt As Icmp_Options

Private Function HiByte(ByVal wParam As Integer)
    Let HiByte = wParam \ &H100 And &HFF&
End Function

Private Function LoByte(ByVal wParam As Integer)
    Let LoByte = wParam And &HFF&
End Function

Public Function Ping(szAddress As String) As Boolean
    Dim hPort As Long
    Dim dwAddress As Long
    Dim sDataToSend As String
    Dim Echo As Icmp_Echo_Reply
    Dim iOpt As Long
    Let sDataToSend = "Data"
    Call Me.SocketsInitialize
    Let dwAddress = AddressStringToLong(szAddress)
    Let hPort = IcmpCreateFile()
    Let Ping = False
    If (IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, Echo, Len(Echo), PING_TIMEOUT)) Then
        Let Ping = (Echo.status = 0)
    End If
    Call Me.SocketsCleanup
    Call IcmpCloseHandle(hPort)
End Function
   
Private Function AddressStringToLong(ByVal TMP As String) As Long
    Dim i As Integer
    Dim parts(1 To 4) As String
    Let i = 0
    Let AddressStringToLong = 0
    While (InStr(TMP, ".") > 0)
        Let i = i + 1
        Let parts(i) = Mid(TMP, 1, InStr(TMP, ".") - 1)
        Let TMP = Mid(TMP, InStr(TMP, ".") + 1)
    Wend
    Let i = i + 1
    Let parts(i) = TMP
    If (i <> 4) Then
        Exit Function
    End If
    Let AddressStringToLong = Val("&H" & Right("00" & Hex(parts(4)), 2) & Right("00" & Hex(parts(3)), 2) & Right("00" & Hex(parts(2)), 2) & Right("00" & Hex(parts(1)), 2))
End Function

Public Function SocketsCleanup() As Boolean
    Dim X As Long
    Let X = WSACleanup()
    Let SocketsCleanup = True
    If (X <> 0) Then
        Call SHOWWRONG("Windows Sockets error " & Trim$(str$(X)) & " occurred in Cleanup.", 0)
        Let SocketsCleanup = False
    End If
End Function

Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSAData
    Dim X As Integer
    Dim szLoByte As String, szHiByte As String, szBuf As String
    Let X = WSAStartup(WS_VERSION_REQD, WSAD)
    Let SocketsInitialize = False
    If (X <> 0) Then
        Call SHOWWRONG("Windows Sockets for 32 bit Windows " & "environments is not successfully responding.", 2)
        Exit Function
    End If
    If ((LoByte(WSAD.wVersion) < WS_VERSION_MAJOR) Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR)) Then
        Let szHiByte = Trim$(str$(HiByte(WSAD.wVersion)))
        Let szLoByte = Trim$(str$(LoByte(WSAD.wVersion)))
        Let szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        Let szBuf = szBuf & " is not supported by Windows " & "Sockets for 32 bit Windows environments."
        Call SHOWWRONG(szBuf, 2)
        Exit Function
    End If
    If (WSAD.wMaxSockets < MIN_SOCKETS_REQD) Then
        Let szBuf = "This application requires a minimum of " & Trim$(str$(MIN_SOCKETS_REQD)) & " supported sockets."
        Call SHOWWRONG(szBuf, 2)
        Exit Function
    End If
    Let SocketsInitialize = True
End Function

Public Function ResolveHostname(ByVal ipaddress As String) As String
    Dim hostip_addr As Long
    Dim hostent_addr As Long
    Dim newAddr As Long
    Dim Host As HOSTENT
    Dim strTemp As String
    Dim strHost As String * 255
    If (SocketsInitialize()) Then
        Let newAddr = inet_addr(ipaddress)
        Let hostent_addr = gethostbyaddr(newAddr, Len(newAddr), AF_INET)
        If (hostent_addr = 0) Then
            Call SocketsCleanup
            Exit Function
        End If
        Call RtlMoveMemory(Host, hostent_addr, Len(Host))
        Call RtlMoveMemory(ByVal strHost, Host.hName, 255)
        Let strTemp = strHost
        If InStr(strTemp, Chr(0)) <> 0 Then strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
        Let strTemp = Trim(strTemp)
        Let ResolveHostname = strTemp
        Call SocketsCleanup
    End If
End Function




