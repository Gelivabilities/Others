VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl ICEE_NET 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin MSWinsockLib.Winsock WskItem 
      Index           =   0
      Left            =   600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ICEE_NET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To &HFF) As RGBQUAD
End Type

Private Const BI_RGB As Long = 0&


Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long


Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Const DIB_RGB_COLORS As Long = 0
Private Const DIB_PAL_COLORS As Long = 1


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Private Declare Function SetDIBitsToDevice Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nDestWidth As Long, ByVal nDestHeight As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long


Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, lpbmi As Any, ByVal iUsage As Long, ByRef ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long



Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const BDR_RAISEDOUTER As Long = &H1
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_OUTER As Long = &H3
Private Const BDR_RAISEDINNER As Long = &H4
'private Const BDR_RAISED As Long = &H5
Private Const BDR_SUNKENINNER As Long = &H8
'private Const BDR_SUNKEN As Long = &HA
Private Const BDR_INNER As Long = &HC
Private Const EDGE_RAISED As Long = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_ETCHED As Long = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP As Long = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_SUNKEN As Long = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_DIAGONAL As Long = &H10
Private Const BF_MIDDLE As Long = &H800
Private Const BF_SOFT As Long = &H1000
Private Const BF_ADJUST As Long = &H2000
Private Const BF_FLAT As Long = &H4000
Private Const BF_MONO As Long = &H8000
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_TOPLEFT As Long = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT As Long = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT As Long = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT As Long = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDBOTTOMLEFT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT As Long = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT As Long = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDTOPRIGHT As Long = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)




'////////////////////////////////////////////////
'################################################
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Const SizeofMyCommandHeader As Long = 4
Private Type MyCommandHeader
    Sign As Byte '识别标记（=MyCommandSign）
    Code As Byte '命令（eMyCommandID）
    Size As Integer '数据的长度（单位字节）.不可>=&H4000
End Type
Private Const MyCommandSign As Byte = &HFF

'[S/C]:服务器/客户机发送的（客户机/服务器处理指令）
'[S>C]:服务器发送的（客户机处理指令）
'[C>S]:客户机发送的（服务器处理指令）
Private Enum eMyCommandID
    MyCID_Null = 0  '[...]（保留）
    MyCID_Stop      '[S/C]结束传输                  （附加数据:0）
    MyCID_QVer      '[C>S]查询版本号                （附加数据:0）
    MyCID_Ver       '[S>C]得到版本号                （附加数据:2）
    MyCID_Next      '[C>S]提示服务器发送下一幅图片  （附加数据:0）
    MyCID_Info      '[S>C]图像数据信息              （附加数据:8）
    MyCID_QData     '[C>S]请求数据                  （附加数据:0）
    MyCID_Send      '[S>C]服务器发来图像数据        （附加数据:4+x）
End Enum

Private Const SoftVer As Integer = &H100
Private mCurVer As Integer


'MyCID_Info
Private Type MyImageInfo '8Byte
    SizeImage As Long
    Width As Integer
    Height As Integer
End Type



'~~ 处理流程 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'[C]Connect事件: 发送MyCID_QVer
'[S]发送MyCID_Ver
'[C]If 版本号正确 Then
'    [C]发送MyCID_Next
'    [S]触发OnQueryPicture事件
'    [S]压缩图像
'    [S]发送MyCID_Info
'    Do
'        [C]发送MyCID_QData
'        [S]发送MyCID_Send
'    While Until 图像压缩数据接收完毕
'    [C]发送MyCID_Next（这样可以实现并行处理）
'    [C]解压图像数据
'    [C]触发OnPictureArrival事件
'    ……
'~~ 另一种表示法 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'01.[C]>>MyCID_QVer>>[S]
'02.[C]<<MyCID_Ver<<[S]
'03.[C]判断版本号，若不能匹配则中断连结
'04.[C]>>MyCID_Next>>[S]
'06.                 [S]触发OnQueryPicture事件
'07.                 [S]压缩图像
'08.[C]<<MyCID_Info<<[S]
'09.[C]>>MyCID_QData>>[S]
'10.[C]<<MyCID_Send<<[S]
'11.[C]若数据没有接收完，则转到9
'12.[C]>>MyCID_Next>>[S]（这样可以实现并行处理）
'13.[C]解压图像数据
'14.[C]触发OnPictureArrival事件
'15.因12发送的指令，转到8


Private DitherTable(0 To &HFF) As Byte '抖动模板

Private palWeb216(0 To &HFF) As RGBQUAD

'Private Num8Bto6(0 To &HFF) As Long '\&H33
'Private Num6to8B(0 To 5) As Long    '*&H33
Private Diff8Bto6(0 To &HFF) As Long '8位数据转为6种后的误差（格式化到[0,256]区间）

Private mBI As BITMAPINFO  '位图信息
Private mScanBytes As Long '扫描行字节数
Private mMapData() As Byte '位图数据
Private mIsChangeBitmap As Boolean '图片是否改变（是否需要重新编码）

Private Const MaxFrameSize As Integer = &HF00 '最大封包大小

'Private LZWStream() As Byte 'LZW数据流
'Private LZWStreamSize As Long 'LZW数据流长度
'Private LZWStreamPos As Long '当前位置
Private mLZWS As New CByteStream
Private mImgInfo As MyImageInfo

Private Const PicColorBits As Byte = 8 '图片颜色位数
Private Const LZW_MinCodeLen As Byte = PicColorBits '最小编码单元
Private Const LZW_MaxCodeBits As Long = 12 'GIF-LZW最大编码长度

Private mInited As Boolean

Private bClosing As Boolean '正在关闭
Private bDecode As Boolean '准备解码

Private mCurClients As Long

'Private DataStream() As Byte '数据流
'Private DataStreamSize As Long '数据流长度

Private mCmdS As New CByteStream '命令数据流

Private Type ServerData
    CmdS As CByteStream
    LZWS As CByteStream
End Type
Private mServers() As ServerData

Public Event CloseConnect() '关闭连结

Public Event OnQueryPicture()   '[S]请求新的图片
Public Event OnPictureArrival() '[C]图片已经接收

'缺省属性值:
Const m_def_MaxClient = 100
Const m_def_IsServer = False
'属性变量:
Dim m_MaxClient As Long
Private m_IsServer As Boolean
Private Sub pInit()
    If mInited Then Exit Sub
    mInited = True
    
    Debug.Print String(60, "=")
    
    Call 桌面传输.Init
    
    mCurVer = SoftVer
    mCurClients = 0
    ReDim mServers(WskItem.LBound To WskItem.UBound)
    'Debug.Print "Init"
    
    Dim TempArr As Variant
    Dim i As Long, j As Long, K As Long
    Dim Idx As Long
    
    TempArr = Array(0, 235, 59, 219, 15, 231, 55, 215, 2, 232, 56, 217, 12, 229, 52, 213, _
            128, 64, 187, 123, 143, 79, 183, 119, 130, 66, 184, 120, 140, 76, 180, 116, _
            33, 192, 16, 251, 47, 207, 31, 247, 34, 194, 18, 248, 44, 204, 28, 244, _
            161, 97, 144, 80, 175, 111, 159, 95, 162, 98, 146, 82, 172, 108, 156, 92, _
            8, 225, 48, 208, 5, 239, 63, 223, 10, 226, 50, 210, 6, 236, 60, 220, _
            136, 72, 176, 112, 133, 69, 191, 127, 138, 74, 178, 114, 134, 70, 188, 124, _
            41, 200, 24, 240, 36, 197, 20, 255, 42, 202, 26, 242, 38, 198, 22, 252, _
            169, 105, 152, 88, 164, 100, 148, 84, 170, 106, 154, 90, 166, 102, 150, 86, _
            3, 233, 57, 216, 13, 228, 53, 212, 1, 234, 58, 218, 14, 230, 54, 214, _
            131, 67, 185, 121, 141, 77, 181, 117, 129, 65, 186, 122, 142, 78, 182, 118, _
            35, 195, 19, 249, 45, 205, 29, 245, 32, 193, 17, 250, 46, 206, 30, 246, _
            163, 99, 147, 83, 173, 109, 157, 93, 160, 96, 145, 81, 174, 110, 158, 94, _
            11, 227, 51, 211, 7, 237, 61, 221, 9, 224, 49, 209, 4, 238, 62, 222, _
            139, 75, 179, 115, 135, 71, 189, 125, 137, 73, 177, 113, 132, 68, 190, 126, _
            43, 203, 27, 243, 39, 199, 23, 253, 40, 201, 25, 241, 37, 196, 21, 254, _
            171, 107, 155, 91, 167, 103, 151, 87, 168, 104, 153, 89, 165, 101, 149, 85)
    For i = 0 To &HFF
        DitherTable(i) = TempArr(i)
    Next i
    
    For i = 0 To 5 'Blue
        For j = 0 To 5 'Green
            For K = 0 To 5 'Red
                Idx = (i * 6 + j) * 6 + K
                palWeb216(Idx).rgbRed = K * &H33
                palWeb216(Idx).rgbGreen = j * &H33
                palWeb216(Idx).rgbBlue = i * &H33
            Next K
        Next j
    Next i
    
    With mBI.bmiHeader
        .biSize = Len(mBI.bmiHeader)
        .biWidth = 0
        .biHeight = 0
        .biBitCount = PicColorBits
        .biPlanes = 1
        .biCompression = BI_RGB
        mScanBytes = 0
        .biSizeImage = mScanBytes * .biHeight
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = 0
        .biClrImportant = 0
    End With
    CopyMemory mBI.bmiColors(0), palWeb216(0), &H100 * 4
    
    For i = 0 To &HFF
        Diff8Bto6(i) = ((i - (i \ &H33) * &H33) * &H100 + (&H33 \ 2)) \ &H33
    Next i
    
End Sub

'设置位图
Private Function pSetBitmap(ByVal hdc As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal Width As Long, ByVal Height As Long) As Boolean
    If Width <= 0 Or Height <= 0 Then Exit Function
    
    Dim rc As Boolean
    
    Dim SrcBI As BITMAPINFOHEADER
    Dim ScanBytes As Long
    Dim pSrcDIB As Long
    Dim hSrcDIB As Long
    Dim hDCDIB As Long
    Dim hOldMap As Long
    
    hDCDIB = CreateCompatibleDC(hdc)
    If hDCDIB Then
        With SrcBI
            .biSize = Len(SrcBI)
            .biWidth = Width
            .biHeight = Height
            .biBitCount = 24
            .biPlanes = 1
            .biCompression = BI_RGB
            ScanBytes = (.biWidth * 3 + 3) And &H7FFFFFFC
            .biSizeImage = ScanBytes * .biHeight
            .biXPelsPerMeter = 0
            .biYPelsPerMeter = 0
            .biClrUsed = 0
            .biClrImportant = 0
        End With
        hSrcDIB = CreateDIBSection(hDCDIB, SrcBI, DIB_RGB_COLORS, pSrcDIB, 0, 0)
        If hSrcDIB Then
            hOldMap = SelectObject(hDCDIB, hSrcDIB)
            
            Call BitBlt(hDCDIB, 0, 0, Width, Height, hdc, X, Y, vbSrcCopy)
            
            '// 进行色彩量化
            Dim pByteS() As Byte, pBytePtrS As SAFEARRAY1D
            Dim pByteD() As Byte, pBytePtrD As SAFEARRAY1D
            Dim ScanAddS As Long
            Dim ScanAddD As Long
            Dim i As Long, j As Long
            Dim CurDither As Long
            
            '初始化DIB
            With mBI.bmiHeader
                .biSize = Len(mBI.bmiHeader)
                .biWidth = Width
                .biHeight = Height
                .biBitCount = PicColorBits
                .biPlanes = 1
                .biCompression = BI_RGB
                mScanBytes = (.biWidth * 1 + 3) And &H7FFFFFFC
                .biSizeImage = mScanBytes * .biHeight
                .biXPelsPerMeter = 0
                .biYPelsPerMeter = 0
                .biClrUsed = 0
                .biClrImportant = 0
                
                If .biSizeImage > 0 Then
                    ReDim mMapData(0 To .biSizeImage - 1)
                Else
                    .biSizeImage = 0
                    mScanBytes = 0
                End If
                
            End With
            If mScanBytes > 0 Then
                CopyMemory mBI.bmiColors(0), palWeb216(0), &H100 * 4
                
                MakePoint VarPtrArray(pByteS), pBytePtrS, 1
                MakePoint VarPtrArray(pByteD), pBytePtrD, 1
                
                ScanAddS = ScanBytes - SrcBI.biWidth * 3
                ScanAddD = mScanBytes - mBI.bmiHeader.biWidth * 1
                Ptr(pBytePtrS) = pSrcDIB
                Ptr(pBytePtrD) = VarPtr(mMapData(0))
                
                For i = 0 To SrcBI.biHeight - 1
                    For j = 0 To SrcBI.biWidth - 1
                        CurDither = DitherTable((i And &HF) * &H10 + (j And &HF))
                        pByteD(0) = (pByteS(0) \ &H33 + ((Diff8Bto6(pByteS(0)) > CurDither) And 1)) * 36 _
                                  + (pByteS(1) \ &H33 + ((Diff8Bto6(pByteS(1)) > CurDither) And 1)) * 6 _
                                  + (pByteS(2) \ &H33 + ((Diff8Bto6(pByteS(2)) > CurDither) And 1)) * 1
                        
                        pBytePtrS.pvData = pBytePtrS.pvData + 3
                        pBytePtrD.pvData = pBytePtrD.pvData + 1
                        
                    Next j
                    
                    pBytePtrS.pvData = pBytePtrS.pvData + ScanAddS
                    pBytePtrD.pvData = pBytePtrD.pvData + ScanAddD
                    
                Next i
                
                FreePoint VarPtrArray(pByteS)
                FreePoint VarPtrArray(pByteD)
                
            End If
            '\\ 进行色彩量化
            
            Call SelectObject(hDCDIB, hOldMap)
            DeleteObject hSrcDIB
            
            rc = True
            
        End If
        
        DeleteDC hDCDIB
        
    End If
    
    If rc Then
        mIsChangeBitmap = True
    End If
    
    pSetBitmap = rc
    
End Function

'图像编码
Private Sub pEncode()
    If mIsChangeBitmap = False Then Exit Sub
    'Debug.Assert False
    
    If mScanBytes > 0 Then
        'GIF-LZW编码
        Dim NextNode(0 To &H1000) As Integer '第一个下层节点的索引
        Dim SubNode(0 To &H1000) As Integer '下一个同层节点的索引
        Dim StrAdd(0 To &H1000) As Byte '新增加的那个字节（比上层节点多的那个字节）
        Dim TableSize As Long
        Dim TableMaxSize As Long
        Dim CurBits As Long
        Dim LZW_CLEAR As Integer
        Dim LZW_EOI As Integer
        Dim OldCode As Integer
        Dim CurByte As Byte
        Dim TempIdx As Integer
        Dim f As Boolean
        '模拟指针
        Dim pByteS() As Byte, pBytePtrS As SAFEARRAY1D
        Dim ScanPtr As Long
        '缓冲区
        Dim BitBuff As Long
        Dim BitUsed As Long
        Dim BufLZW() As Byte
        Dim BufLZWPos As Long
        '其他
        Dim X As Long, Y As Long
        
        '分配缓冲区
        ReDim BufLZW(0 To mBI.bmiHeader.biSizeImage * 2 - 1)
        BufLZWPos = 0
        BitBuff = 0
        BitUsed = 0
        
        '建立模拟指针
        MakePoint VarPtrArray(pByteS), pBytePtrS, 1
        
        '初始化LZW字符串表
        LZW_CLEAR = BitPosMask(LZW_MinCodeLen) '2 ^ LZW_MinCodeLen '1<<LZW_MinCodeLen
        LZW_EOI = LZW_CLEAR + 1
        CurBits = LZW_MinCodeLen + 1
        'GoSub InitStrTable
            OldCode = LZW_CLEAR
            'GoSub WriteCode
                BitBuff = BitBuff Or OldCode * BitPosMask(BitUsed)
                BitUsed = BitUsed + CurBits
                'GoSub ShiftBit
                    Do While BitUsed >= 8
                        BufLZW(BufLZWPos) = BitBuff And &HFF
                        BufLZWPos = BufLZWPos + 1
                        BitBuff = BitBuff \ &H100 'BitBuff>>8
                        BitUsed = BitUsed - 8
                    Loop
            CurBits = LZW_MinCodeLen + 1
            TableSize = LZW_EOI + 1
            TableMaxSize = BitPosMask(CurBits) '2 ^ CurBits '1<<CurBits
            Call ZeroMemory(NextNode(0), &H2000)
            Call ZeroMemory(SubNode(0), &H2000)
            Call ZeroMemory(StrAdd(0), &H1000)
        
        '初始化位图
        Y = 0
        X = 0
        ScanPtr = VarPtr(mMapData(0)) + (mBI.bmiHeader.biHeight - 1) * mScanBytes  'DIB是逆序存储
        pBytePtrS.pvData = ScanPtr
        
        '正式开始
        OldCode = pByteS(0)
        pBytePtrS.pvData = pBytePtrS.pvData + 1
        X = X + 1
        Do While Y < mBI.bmiHeader.biHeight
            '得到数据
            CurByte = pByteS(0)
            
            '看编码是否在字符串表中
            f = SubNode(OldCode) '没有下级节点，就必然不在
            If f Then '进一步判断
                TempIdx = SubNode(OldCode) '得到当前层节点的索引
                Do Until StrAdd(TempIdx) = CurByte '判断是否是已存在的节点
                    If NextNode(TempIdx) Then '存在下一节点
                        TempIdx = NextNode(TempIdx) '指向下一节点
                    Else '不存在下一节点
                        NextNode(TempIdx) = TableSize '设置同层下一节点索引指针
                        f = False
                        Exit Do
                    End If
                Loop
            Else
                SubNode(OldCode) = TableSize '设置下层节点索引指针
            End If
            
            If f Then '在
                OldCode = TempIdx
            Else '不在
                '添加编码
                'GoSub WriteCode
                    BitBuff = BitBuff Or OldCode * BitPosMask(BitUsed)
                    BitUsed = BitUsed + CurBits
                    'GoSub ShiftBit
                        Do While BitUsed >= 8
                            BufLZW(BufLZWPos) = BitBuff And &HFF
                            BufLZWPos = BufLZWPos + 1
                            BitBuff = BitBuff \ &H100 'BitBuff>>8
                            BitUsed = BitUsed - 8
                        Loop
                StrAdd(TableSize) = CurByte
                TableSize = TableSize + 1
                
                '判断字符串表大小
                If TableSize > TableMaxSize Then
                    If CurBits < LZW_MaxCodeBits Then
                        CurBits = CurBits + 1
                        TableMaxSize = TableMaxSize * 2 'tablemaxsize<<=1
                    Else
                        'GoSub InitStrTable
                            OldCode = LZW_CLEAR
                            'GoSub WriteCode
                                BitBuff = BitBuff Or OldCode * BitPosMask(BitUsed)
                                BitUsed = BitUsed + CurBits
                                'GoSub ShiftBit
                                    Do While BitUsed >= 8
                                        BufLZW(BufLZWPos) = BitBuff And &HFF
                                        BufLZWPos = BufLZWPos + 1
                                        BitBuff = BitBuff \ &H100 'BitBuff>>8
                                        BitUsed = BitUsed - 8
                                    Loop
                            CurBits = LZW_MinCodeLen + 1
                            TableSize = LZW_EOI + 1
                            TableMaxSize = BitPosMask(CurBits) '2 ^ CurBits '1<<CurBits
                            Call ZeroMemory(NextNode(0), &H2000)
                            Call ZeroMemory(SubNode(0), &H2000)
                            Call ZeroMemory(StrAdd(0), &H1000)
                    End If
                End If
                OldCode = CurByte
                
            End If
            
            '移动到下一像素
            X = X + 1
            pBytePtrS.pvData = pBytePtrS.pvData + 1
            
            '判断是否处理完一行
            If X >= mBI.bmiHeader.biWidth Then
                Y = Y + 1
                ScanPtr = ScanPtr - mScanBytes
                pBytePtrS.pvData = ScanPtr
                X = 0
            End If
            
        Loop
        
        '输出最后一个编码
        'GoSub WriteCode
            BitBuff = BitBuff Or OldCode * BitPosMask(BitUsed)
            BitUsed = BitUsed + CurBits
            'GoSub ShiftBit
                Do While BitUsed >= 8
                    BufLZW(BufLZWPos) = BitBuff And &HFF
                    BufLZWPos = BufLZWPos + 1
                    BitBuff = BitBuff \ &H100 'BitBuff>>8
                    BitUsed = BitUsed - 8
                Loop
        
        '输出LZW_EOI
        OldCode = LZW_EOI
        'GoSub WriteCode
            BitBuff = BitBuff Or OldCode * BitPosMask(BitUsed)
            BitUsed = BitUsed + CurBits
            'GoSub ShiftBit
                Do While BitUsed >= 8
                    BufLZW(BufLZWPos) = BitBuff And &HFF
                    BufLZWPos = BufLZWPos + 1
                    BitBuff = BitBuff \ &H100 'BitBuff>>8
                    BitUsed = BitUsed - 8
                Loop
        
        '结束位流
        If BitUsed Then
            BitUsed = 8
            'GoSub ShiftBit
                BufLZW(BufLZWPos) = BitBuff And &HFF
                BufLZWPos = BufLZWPos + 1
                BitBuff = BitBuff \ &H100 'BitBuff>>8
                BitUsed = BitUsed - 8
        End If
        
        '释放模拟指针
        FreePoint VarPtrArray(pByteS)
        
        '复制位流
        Call mLZWS.Clear
        Call mLZWS.AddData4Ptr(VarPtr(BufLZW(0)), BufLZWPos)
        
    End If
    
    mIsChangeBitmap = False
    
End Sub

'图像解码
Private Sub pDecode()
    'Debug.Assert False
    
    With mBI.bmiHeader
        .biWidth = mImgInfo.Width
        .biHeight = mImgInfo.Height
        mScanBytes = (.biWidth + 3) And &H7FFFFFFC
        .biSizeImage = mScanBytes * .biWidth
        
        If .biWidth > 0 And .biHeight > 0 Then
            ReDim mMapData(0 To .biSizeImage - 1)
        Else
            mScanBytes = 0
        End If
        
    End With
    
    If mScanBytes > 0 Then
        'GIF-LZW解码
        Dim StrAdd(0 To &H1000) As Byte '新增加的那个字节（比上层节点多的那个字节）
        Dim Parent(0 To &H1000) As Integer '父节点的索引指针
        Dim Level(0 To &H1000) As Integer '当前节点共有多少层（当前节点有多少字节数据）
        Dim TableSize As Long
        Dim TableMaxSize As Long
        Dim BufCode(0 To &H1000) As Byte '单个编码解压的缓冲区
        Dim cbBufCode As Long
        Dim CurBits As Long
        Dim LZW_CLEAR As Integer
        Dim LZW_EOI As Integer
        Dim CurCode As Long
        Dim OldCode As Integer
        'Dim CurByte As Byte
        Dim TempIdx As Integer
        Dim f As Boolean
        '模拟指针
        Dim CurPtr As Long
        Dim ScanPtr As Long
        '缓冲区
        Dim BitBuff As Long
        Dim BitUsed As Long
        Dim BufLZW() As Byte
        Dim BufLZWPos As Long
        Dim BufLZWSize As Long
        '其他
        Dim X As Long, Y As Long
        Dim i As Long
        
        '分配缓冲区
        BufLZWSize = mLZWS.PeekData(BufLZW)
        BufLZWPos = 0
        BitBuff = 0
        BitUsed = 0
        If BufLZWSize > 0 Then
            '初始化LZW字符串表
            LZW_CLEAR = BitPosMask(LZW_MinCodeLen) '2 ^ LZW_MinCodeLen '1<<LZW_MinCodeLen
            LZW_EOI = LZW_CLEAR + 1
            CurBits = LZW_MinCodeLen + 1
            'GoSub GetNextCode
                Do While BitUsed < CurBits
                    BitBuff = BitBuff Or (BufLZW(BufLZWPos) * BitPosMask(BitUsed)) 'TempCode |= BufLZW(BufLZWPos)<<BitBuff
                    BufLZWPos = BufLZWPos + 1
                    If BufLZWPos >= BufLZWSize Then GoTo DecodeEnd
                    BitUsed = BitUsed + 8
                Loop
                CurCode = BitBuff And BitsMask(CurBits)
                BitBuff = BitBuff \ BitPosMask(CurBits)
                BitUsed = BitUsed - CurBits
            
            OldCode = CurCode
            If OldCode = LZW_CLEAR Then '正确的编码
                'GoSub InitStrTable
                    CurBits = LZW_MinCodeLen + 1
                    TableSize = LZW_EOI + 1
                    TableMaxSize = 2 ^ CurBits
                    Call ZeroMemory(StrAdd(0), &H1000)
                    Call ZeroMemory(Parent(0), &H2000)
                    Call ZeroMemory(Level(0), &H2000)
                
                '初始化位图
                Y = 0
                X = 0
                ScanPtr = VarPtr(mMapData(0)) + (mBI.bmiHeader.biHeight - 1) * mScanBytes  'DIB是逆序存储
                CurPtr = ScanPtr
                
                '正式开始
                Do
                    If CurCode = LZW_CLEAR Then
                        'GoSub InitStrTable
                            CurBits = LZW_MinCodeLen + 1
                            TableSize = LZW_EOI + 1
                            TableMaxSize = 2 ^ CurBits
                            Call ZeroMemory(StrAdd(0), &H1000)
                            Call ZeroMemory(Parent(0), &H2000)
                            Call ZeroMemory(Level(0), &H2000)
                        Do
                            'GoSub GetNextCode
                                Do While BitUsed < CurBits
                                    BitBuff = BitBuff Or (BufLZW(BufLZWPos) * BitPosMask(BitUsed)) 'TempCode |= BufLZW(BufLZWPos)<<BitBuff
                                    BufLZWPos = BufLZWPos + 1
                                    If BufLZWPos >= BufLZWSize Then GoTo DecodeEnd
                                    BitUsed = BitUsed + 8
                                Loop
                                CurCode = BitBuff And BitsMask(CurBits)
                                BitBuff = BitBuff \ BitPosMask(CurBits)
                                BitUsed = BitUsed - CurBits
                        Loop While CurCode = LZW_CLEAR
                        If CurCode >= LZW_EOI Then Debug.Print "过界," & Y: GoTo DecodeEnd
                        
                        '第一个编码
                        BufCode(0) = CurCode
                        cbBufCode = 1
                        
                    ElseIf CurCode = LZW_EOI Then
                        Exit Do
                        
                    Else
                        If OldCode = LZW_CLEAR Then
                            '不可能出现OldCode = LZW_CLEAR的情况，所以一定出错了
                            Debug.Assert False
                            GoTo DecodeEnd
                        End If
                        
                        If TableSize > TableMaxSize Then
                            '字符串表已达最大大小
                            '同时没有LZW_CLEAR
                            '无法解决字符串表问题
                            Debug.Print ">"
                            GoTo DecodeEnd
                        End If
                        
                        '解压数据
                        TempIdx = IIf(CurCode < TableSize, CurCode, OldCode)
                            'If CurCode < TableSize
                                '表示CurCode在字符串表中，所以使用CurCode
                            'Else
                                '表示CurCode不在字符串表中，只有使用OldCode
                        cbBufCode = Level(TempIdx)
                        For i = 0 To cbBufCode - 1
                            BufCode(cbBufCode - i) = StrAdd(TempIdx)
                            TempIdx = Parent(TempIdx)
                        Next i
                        If TempIdx > &HFF Then
                            GoSub DecodeEnd
                        End If
                        BufCode(0) = TempIdx '最后一个字节是节点索引本身（0~255的默认数据）
                        cbBufCode = cbBufCode + 1
                        If CurCode >= TableSize Then '不在字符串表中
                            '= +GetFirstChar(Code2Str(OldCode))
                            'GIF-LZW解码算法规定，还得加上OldCode的第一字节
                            BufCode(cbBufCode) = BufCode(0)
                            cbBufCode = cbBufCode + 1
                        End If
                        
                        '添加新节点
                        StrAdd(TableSize) = BufCode(0)
                        Parent(TableSize) = OldCode
                        Level(TableSize) = Level(OldCode) + 1
                        TableSize = TableSize + 1
                        If TableSize >= TableMaxSize Then
                            If CurBits < LZW_MaxCodeBits Then
                                CurBits = CurBits + 1
                                TableMaxSize = TableMaxSize * 2
                            End If
                        End If
                        
                    End If
                    OldCode = CurCode
                    
                    '取得下一节点
                    'GoSub GetNextCode
                        Do While BitUsed < CurBits
                            BitBuff = BitBuff Or (BufLZW(BufLZWPos) * BitPosMask(BitUsed)) 'TempCode |= BufLZW(BufLZWPos)<<BitBuff
                            BufLZWPos = BufLZWPos + 1
                            If BufLZWPos >= BufLZWSize Then GoTo DecodeEnd
                            BitUsed = BitUsed + 8
                        Loop
                        CurCode = BitBuff And BitsMask(CurBits)
                        BitBuff = BitBuff \ BitPosMask(CurBits)
                        BitUsed = BitUsed - CurBits
                    
                    '复制位图数据
                    i = 0
                    While i < cbBufCode
                        TempIdx = mBI.bmiHeader.biWidth - X
                        If cbBufCode - i >= TempIdx Then '数据不在一扫描行内
                            CopyMemory ByVal CurPtr, BufCode(i), TempIdx
                            i = i + TempIdx
                            
                            '非交错
                            Y = Y + 1
                            If Y >= mBI.bmiHeader.biHeight Then Exit Do '超过图像大小
                            ScanPtr = ScanPtr - mScanBytes
                            CurPtr = ScanPtr
                            
                            X = 0
                            
                        Else '数据在一扫描行内
                            TempIdx = cbBufCode - i
                            CopyMemory ByVal CurPtr, BufCode(i), TempIdx
                            i = i + TempIdx
                            X = X + TempIdx
                            CurPtr = CurPtr + TempIdx
                            
                            '退出“复制位图数据”循环
                            cbBufCode = 0
                            
                        End If
                    Wend
                    cbBufCode = 0
                    
                Loop
                
DecodeEnd: '编码结束
                
            End If
            
        End If
        
    End If
    
End Sub

'## 控件事件 ##############################################

Private Sub UserControl_InitProperties()
    Call pInit
    
    m_IsServer = m_def_IsServer
    m_MaxClient = m_def_MaxClient
End Sub

Private Sub UserControl_Paint()
    Dim rct As RECT
    
    With UserControl
        '-- 绘制边框
        rct.Left = 0
        rct.Top = 0
        rct.Right = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)
        rct.Bottom = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels)
        Call DrawEdge(.hdc, rct, EDGE_BUMP, BF_RECT)
        
    End With
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Call pInit
    
    wskMain.LocalPort = PropBag.ReadProperty("LocalPort", 0)
    wskMain.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    wskMain.RemotePort = PropBag.ReadProperty("RemotePort", 0)
    m_IsServer = PropBag.ReadProperty("IsServer", m_def_IsServer)
    m_MaxClient = PropBag.ReadProperty("MaxClient", m_def_MaxClient)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    Const SelfEdge = 4

    
End Sub

Private Sub UserControl_Terminate()
    Me.CloseConnect
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    Call PropBag.WriteProperty("LocalPort", wskMain.LocalPort, 0)
    Call PropBag.WriteProperty("RemoteHost", wskMain.RemoteHost, "")
    Call PropBag.WriteProperty("RemotePort", wskMain.RemotePort, 0)
    Call PropBag.WriteProperty("IsServer", m_IsServer, m_def_IsServer)
    Call PropBag.WriteProperty("MaxClient", m_MaxClient, m_def_MaxClient)
End Sub

Private Sub WskItem_Close(Index As Integer)
    'Debug.Print "WskItem_Close["; Index; "]:"
    
    mCurClients = mCurClients - 1
    
    '清空命令流
    Call mServers(Index).CmdS.Clear
    
    '释放LZW位流对象
    Set mServers(Index).LZWS = Nothing
    
End Sub

Private Sub WskItem_Connect(Index As Integer)
    'Debug.Print "WskItem_Connect["; Index; "]:"
    'mCurClients = mCurClients + 1
End Sub

Private Sub WskItem_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'Debug.Print "WskItem_ConnectionRequest["; Index; "]:"
    
End Sub

Private Sub WskItem_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Debug.Print "WskItem_DataArrival["; Index; "]: bytesTotal="; bytesTotal
    
    Dim tCmdR As MyCommandHeader, tCmdS As MyCommandHeader
    Dim TempBytes() As Byte
    
    '得到数据
    Call WskItem(Index).GetData(TempBytes, vbByte Or vbArray, bytesTotal)
    With mServers(Index)
        '联接位流
        Call .CmdS.AddData(TempBytes)
        
        Do While .CmdS.Count >= SizeofMyCommandHeader
            '查探数据流是否足够长度
            Call .CmdS.PeekData4Ptr(VarPtr(tCmdR), , SizeofMyCommandHeader)
            If .CmdS.Count < (SizeofMyCommandHeader + tCmdR.Size) Then Exit Do
            Call .CmdS.DeleteData(, SizeofMyCommandHeader)
            
            '处理数据
            '因客户机端都在wskMain处理，所以这里都是服务器端处理
            'Debug.Print "wskItem[" & Index & "]: " & "Command=" & tCmdR.Code & "," & vbTab & "Size=" & tCmdR.SIZE & "(&H" & Hex(tCmdR.SIZE) & ")"
            Select Case tCmdR.Code
            Case MyCID_Stop
                If sckClosed <> WskItem(Index).State Then WskItem(Index).Close
                
            Case MyCID_QVer
                With tCmdS
                    .Sign = MyCommandSign
                    .Code = MyCID_Ver
                    .Size = 2
                    ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
                End With
                CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
                
                mCurVer = SoftVer
                CopyMemory TempBytes(SizeofMyCommandHeader), mCurVer, tCmdS.Size
                
                Call WskItem(Index).SendData(TempBytes)
                
            Case MyCID_Ver
                '（无）
                
            Case MyCID_Next
                Dim tInfo As MyImageInfo
                
                '请求图像
                RaiseEvent OnQueryPicture
                
                '压缩图像
                Call pEncode
                
                '提交压缩数据
                'Call .LZWS.Clear
                'Call .LZWS.AddData(mLZWS.Data)
                Call .LZWS.CloneFrom(mLZWS)
                
                '发送MyCID_Info
                With tCmdS
                    .Sign = MyCommandSign
                    .Code = MyCID_Info
                    .Size = Len(tInfo)
                    ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
                End With
                CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
                
                With tInfo
                    .SizeImage = mServers(Index).LZWS.Count
                    .Width = mBI.bmiHeader.biWidth
                    .Height = mBI.bmiHeader.biHeight
                End With
                CopyMemory TempBytes(SizeofMyCommandHeader), tInfo, tCmdS.Size
                
                Call WskItem(Index).SendData(TempBytes)
                
            Case MyCID_Info
                '（无）
                
            Case MyCID_QData
                With tCmdS
                    .Sign = MyCommandSign
                    .Code = MyCID_Send
                    .Size = IIf(mServers(Index).LZWS.Count > MaxFrameSize, MaxFrameSize, mServers(Index).LZWS.Count)
                    'Debug.Print .SIZE
                    ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
                End With
                CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
                
                If tCmdS.Size > 0 Then
                    Call .LZWS.GetData4Ptr(VarPtr(TempBytes(SizeofMyCommandHeader)), tCmdS.Size)
                    'Debug.Print TempBytes(SizeofMyCommandHeader)
                End If
                
                Call WskItem(Index).SendData(TempBytes)
                
            Case MyCID_Send
                '（无）
                
            End Select
            
            '删除多余数据
            Call .CmdS.DeleteData(, tCmdR.Size)
            
        Loop
        
    End With
    
End Sub

Private Sub WskItem_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Debug.Print "WskItem_Error["; Index; "]:"
    
    If sckClosed <> WskItem(Index).State Then WskItem(Index).Close
    
    Call SHOWWRONG(Number & vbCrLf & Description, 2)
    
End Sub

Private Sub WskItem_SendComplete(Index As Integer)
    'Debug.Print "WskItem_SendComplete["; Index; "]:"
    
End Sub

Private Sub WskItem_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    'Debug.Print "WskItem_SendProgress["; Index; "]:"
    
End Sub

Private Sub wskMain_Close()
    'Debug.Print "wskMain_Close:"
    
    '清空命令流
    mCmdS.Clear
    
    RaiseEvent CloseConnect
    
End Sub

Private Sub wskMain_Connect()
    'Debug.Print "wskMain_Connect:"
    
    Dim tCmdS As MyCommandHeader
    Dim TempBytes() As Byte
    
    If Me.IsServer Then
    Else
        '发送MyCID_QVer
        With tCmdS
            .Sign = MyCommandSign
            .Code = MyCID_QVer
            .Size = 0
            ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
        End With
        CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
        Call wskMain.SendData(TempBytes)
        
    End If
    
End Sub

Private Sub wskMain_ConnectionRequest(ByVal requestID As Long)
    'Debug.Print "wskMain_ConnectionRequest:" & Hex(requestID)
    
    If Me.IsServer Then
        '## 单连结
        'If wskMain.State <> 0 Then wskMain.Close
        'wskMain.Accept requestID '允许连接请求
        
        '## 多连结
        Dim i As Long
        Dim Idx As Long
        Dim fFree As Boolean
        Dim fErr As Boolean
        
        If mCurClients >= Me.MaxClient Then Exit Sub
        
        Idx = -1
        For i = WskItem.LBound To WskItem.UBound
            On Error Resume Next
            If sckClosing = WskItem(i).State Then WskItem(i).Close
            fErr = ERR.Number
            On Error GoTo 0
            fFree = (sckClosed = WskItem(i).State) '空闲
            
            If fErr Then '控件不存在，创建
                Idx = i
                Load WskItem(Idx)
                Exit For
            Else
                If fFree Then '空闲的
                    Idx = i
                    Exit For
                Else '已连结
                End If
            End If
            
        Next i
        
        '仍没有找到，则创建
        If Idx = -1 Then
            Idx = WskItem.UBound + 1
            Load WskItem(Idx)
            ReDim Preserve mServers(0 To Idx)
        End If
        
        '接收连结请求
        WskItem(Idx).accept requestID
        mCurClients = mCurClients + 1
        Debug.Print "mCurClients:"; mCurClients
        
        '创建流对象
        If mServers(Idx).CmdS Is Nothing Then Set mServers(Idx).CmdS = New CByteStream
        If mServers(Idx).LZWS Is Nothing Then Set mServers(Idx).LZWS = New CByteStream
        
    Else
    End If
    
End Sub

Private Sub wskMain_DataArrival(ByVal bytesTotal As Long)
    'Debug.Print "wskMain_DataArrival: bytesTotal="; bytesTotal
    
    Dim tCmdR As MyCommandHeader, tCmdS As MyCommandHeader
    Dim TempBytes() As Byte
    
    '得到数据
    Call wskMain.GetData(TempBytes, vbByte Or vbArray, bytesTotal)
    
    '联接位流
    Call mCmdS.AddData(TempBytes)
    
    Do While mCmdS.Count >= SizeofMyCommandHeader
        '查探数据流是否足够长度
        Call mCmdS.PeekData4Ptr(VarPtr(tCmdR), , SizeofMyCommandHeader)
        'Debug.Print "wskMain: " & "Command=" & tCmdR.Code & "," & vbTab & "Size=" & tCmdR.SIZE & "(&H" & Hex(tCmdR.SIZE) & ")"
        'If tCmdR.Code = f Then Debug.Assert False
        If mCmdS.Count < (SizeofMyCommandHeader + tCmdR.Size) Then Exit Do
        Call mCmdS.DeleteData(, SizeofMyCommandHeader)
        
        '处理数据
        '因服务器端的wskMain只负责监听，所以这里都是客户机处理
        'Debug.Print "wskMain: " & "Command=" & tCmdR.Code & "," & vbTab & "Size=" & tCmdR.SIZE & "(&H" & Hex(tCmdR.SIZE) & ")"
        Select Case tCmdR.Code
        Case MyCID_Stop
            Me.CloseConnect
            
        Case MyCID_QVer
            '（无）
            
        Case MyCID_Ver
            '取得数据
            Call mCmdS.PeekData4Ptr(VarPtr(mCurVer), , 2)
            Debug.Print "Ver:"; Hex(mCurVer)
            
            '版本号判断
            If mCurVer <> SoftVer Then
                '不符合
                mCmdS.Clear
                Me.CloseConnect
                Debug.Print "ErrVer"
                Exit Do
            End If
            
            '发送MyCID_Next
            With tCmdS
                .Sign = MyCommandSign
                .Code = MyCID_Next
                .Size = 0
                ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
            End With
            CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
            Call wskMain.SendData(TempBytes)
            
        Case MyCID_Next
            '（无）
            
        Case MyCID_Info
            'Dim tInfo As MyImageInfo
            
            '取得数据
            Call mCmdS.PeekData4Ptr(VarPtr(mImgInfo), , Len(mImgInfo))
            'mLZWSSize = mImgInfo.SizeImage
            
            If mImgInfo.SizeImage <= 0 Or mImgInfo.Width <= 0 Or mImgInfo.Height <= 0 Then
                mImgInfo.SizeImage = 0
                Debug.Print "Error Image!"
            Else
                '清空LZW数据流，等待数据
                mLZWS.Clear
            End If
            
            '发送MyCID_QData（若图像数据错误，则发送MyCID_Next）
            With tCmdS
                .Sign = MyCommandSign
                .Code = IIf(mImgInfo.SizeImage > 0, MyCID_QData, MyCID_Next)
                .Size = 0
                ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
            End With
            CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
            Call wskMain.SendData(TempBytes)
            
        Case MyCID_QData
            '（无）
            
        Case MyCID_Send
            '合并数据流
            Call mCmdS.PeekData(TempBytes, , tCmdR.Size)
            'Debug.Print TempBytes(0)
            Call mLZWS.AddData(TempBytes)
            
            '发送标记
            With tCmdS
                .Sign = MyCommandSign
                bDecode = (mLZWS.Count >= mImgInfo.SizeImage)
                .Code = IIf(bDecode, MyCID_Next, MyCID_QData)
                .Size = 0
                ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
            End With
            CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
            Call wskMain.SendData(TempBytes)
            
        End Select
        
        '删除多余数据
        Call mCmdS.DeleteData(, tCmdR.Size)
        
    Loop
    
End Sub

Private Sub wskMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Debug.Print "wskMain_Error: Number=" & Number
    'Debug.Print wskMain.State
    
    Select Case Number
    Case sckConnectionRefused
        Me.CloseConnect
        
    Case Else
        Me.CloseConnect
        Call SHOWWRONG(Number & vbCrLf & Description, 2)
        
    End Select
    
    
End Sub

Private Sub wskMain_SendComplete()
    'Debug.Print "wskMain_SendComplete:"
    
    If bClosing Then
        'If sckClosed <> wskMain.State Then wskMain.Close
        bClosing = False
    End If
    
    '解码图片
    If bDecode Then
        If mImgInfo.SizeImage > 0 Then
            Call pDecode
            
            mLZWS.Clear
            
            RaiseEvent OnPictureArrival
            
        End If
        
        mImgInfo.SizeImage = 0
        bDecode = False
        
    End If
    
End Sub

Private Sub wskMain_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    'Debug.Print "wskMain_SendProgress:"
    
End Sub

'## 外部函数 ##############################################

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,LocalHostName
Public Property Get LocalHostName() As String
    LocalHostName = wskMain.LocalHostName
End Property

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,LocalIP
Public Property Get LocalIP() As String
    LocalIP = wskMain.LocalIP
End Property

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,LocalPort
Public Property Get LocalPort() As Long
    LocalPort = wskMain.LocalPort
End Property

Public Property Let LocalPort(ByVal New_LocalPort As Long)
    wskMain.LocalPort() = New_LocalPort
    PropertyChanged "LocalPort"
End Property

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,RemoteHost
Public Property Get RemoteHost() As String
    RemoteHost = wskMain.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    wskMain.RemoteHost() = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,RemoteHostIP
Public Property Get RemoteHostIP() As String
    RemoteHostIP = wskMain.RemoteHostIP
End Property

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,RemotePort
Public Property Get RemotePort() As Long
    RemotePort = wskMain.RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Long)
    wskMain.RemotePort() = New_RemotePort
    PropertyChanged "RemotePort"
End Property

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,SocketHandle
Public Property Get SocketHandle() As Long
    SocketHandle = wskMain.SocketHandle
End Property

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,State
Public Property Get State() As Integer
    State = wskMain.State
End Property
'注意!不要删除或修改下列被注释的行!
'MemberInfo=0,0,0,false
Public Property Get IsServer() As Boolean
    IsServer = m_IsServer
End Property

Public Property Let IsServer(ByVal New_IsServer As Boolean)
    If wskMain.State <> sckClosed Then Exit Property
    m_IsServer = New_IsServer
    PropertyChanged "IsServer"
End Property

'注意!不要删除或修改下列被注释的行!
'MemberInfo=5
Public Function Connect() As Boolean
    Dim rc As Boolean
    
    If Me.IsServer Then
        On Error Resume Next
        Call wskMain.Listen
        rc = (0 = ERR.Number)
        On Error GoTo 0
    Else
        On Error Resume Next
        Call wskMain.Connect
        rc = (0 = ERR.Number)
        On Error GoTo 0
    End If
    
    Connect = rc
    
End Function

'注意!不要删除或修改下列被注释的行!
'MappingInfo=wskMain,wskMain,-1,Close
Public Sub CloseConnect()
    If Me.IsServer Then
        Dim i As Long
        
        '关闭所有服务连结
        On Error Resume Next
        For i = WskItem.LBound To WskItem.UBound
            If sckClosed <> WskItem(i).State Then WskItem(i).Close
        Next i
        On Error GoTo 0
        
        If sckClosed <> wskMain.State Then wskMain.Close
        
        RaiseEvent CloseConnect
        
    ElseIf sckConnected <> wskMain.State Then
        If sckClosed <> wskMain.State Then wskMain.Close
        
    Else
        Dim tCmdS As MyCommandHeader
        Dim TempBytes() As Byte
        
        With tCmdS
            .Sign = MyCommandSign
            .Code = MyCID_Stop
            .Size = 0
            ReDim TempBytes(0 To SizeofMyCommandHeader + .Size - 1)
        End With
        CopyMemory TempBytes(0), tCmdS, SizeofMyCommandHeader
        Call wskMain.SendData(TempBytes)
        
        'If sckClosed <> wskMain.State Then wskMain.Close
        bClosing = True
        
    End If
    
End Sub

'注意!不要删除或修改下列被注释的行!
'MemberInfo=8,0,0,30
Public Property Get MaxClient() As Long
    MaxClient = m_MaxClient
End Property

Public Property Let MaxClient(ByVal New_MaxClient As Long)
    If New_MaxClient <= 0 Then
        Exit Property
    End If
    If wskMain.State <> sckClosed Then
        Exit Property
    End If
    m_MaxClient = New_MaxClient
    PropertyChanged "MaxClient"
End Property

Public Property Get CurClients() As Long
    CurClients = mCurClients
End Property

'设置位图
Public Function SetBitmap(ByVal hdc As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal Width As Long, ByVal Height As Long) As Boolean
    SetBitmap = pSetBitmap(hdc, X, Y, Width, Height)
End Function

Public Function Draw(ByVal hdc As Long, _
        Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, _
        Optional ByVal Width As Long = DefNum, Optional ByVal Height As Long = DefNum, _
        Optional ByVal SrcX As Long = 0, Optional ByVal SrcY As Long = 0, _
        Optional ByVal SrcWidth As Long = DefNum, Optional ByVal SrcHeight As Long = DefNum, _
        Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Long
    If mScanBytes <= 0 Then Exit Function
    
    If Width = DefNum Then Width = mBI.bmiHeader.biWidth
    If Height = DefNum Then Height = mBI.bmiHeader.biHeight
    If SrcWidth = DefNum Then SrcWidth = mBI.bmiHeader.biWidth
    If SrcHeight = DefNum Then SrcHeight = mBI.bmiHeader.biHeight
    
    Draw = StretchDIBits(hdc, X, Y, Width, Height, SrcX, SrcY, SrcWidth, SrcHeight, mMapData(0), mBI, DIB_RGB_COLORS, dwRop)
    
End Function






