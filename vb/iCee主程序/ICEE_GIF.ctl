VERSION 5.00
Begin VB.UserControl GIF 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  '透明
   CanGetFocus     =   0   'False
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   570
   ClipBehavior    =   0  '无
   ClipControls    =   0   'False
   DrawStyle       =   2  'Dot
   ForeColor       =   &H80000006&
   HasDC           =   0   'False
   HitBehavior     =   0  '无
   MaskColor       =   &H80000014&
   PaletteMode     =   4  'None
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   38
   Windowless      =   -1  'True
End
Attribute VB_Name = "GIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'

Option Explicit

Public Event Click()
Attribute Click.VB_Description = "发生在用户按下然后释放了一个对象鼠标按钮"
Attribute Click.VB_MemberFlags = "200"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "当用户按下鼠标按钮时，一个对象具有焦点"
Public Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MOUSEUP.VB_Description = "发生在用户释放鼠标按钮时，一个对象具有焦点"
Public Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MOUSEMOVE.VB_Description = "发生在用户移动鼠标"
Public Event LoopsEnded() ' fired when a looping gif terminates its loops naturally
Attribute LoopsEnded.VB_Description = "事件发生时，所有的GIF循环已经终止"
Public Event FrameChanged(ByVal FrameIndex As Long, viaTimer As Boolean) ' fired each time a frame is rendered
Public Event RemoteLoadComplete(ByVal gifWidth As Single, ByVal gifHeight As Single, ByRef Cancel As Boolean) ' notifies download of GIF from URL is complete
Public Event RemoteLoadFailure() ' notifies failure to download GIF
Attribute RemoteLoadFailure.VB_Description = "事件发生在一个远程加载的GIF无法读取"
Private z_CbMem   As Long    'Callback allocated memory address
Private z_Cb()    As Long    'Callback thunk array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
' GDI32 APIs used
Private Declare Function SetDIBColorTable Lib "gdi32.dll" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, ByRef pcRGBQuad As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lppt As POINTAPI) As Long
Private Declare Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

' APIs used to create files
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

' constants used
Private Const HALFTONE As Long = 4&  ' used for SetStretchBltMode API
Private Const PALETTECOUNT = 256&    ' number colors in 8 bit palette
Private Const BLACK_BRUSH As Long = 4&
Private Const WHITE_BRUSH As Long = 0&

' Standard Window UDTs
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SAFEARRAYBOUND
End Type
Private Type BITMAPINFOHEADER   ' used to create a DIB
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
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette(0 To PALETTECOUNT - 1) As Long
End Type


' Custom UDTs
Private Type ColorTableSTRCUT
    Index As Long               ' 0=global; 1-xxx are local tables, if any
    Tables() As Long            ' 2D array (0 to 256, 0 to xxx); color table(s)
End Type
Private Type GIFcoreProperties
    Width As Long               ' overall GIF width
    Height As Long              ' overall height
    Loops As Long               ' Nr loops defined in GIF (0=infinite)
    ScaleCx As Single ' pre calculated ratios when scaling/stretching is used
    ScaleCy As Single ' pre calculated ratios when scaling/stretching is used
End Type
Private Type GIFframeProperties
    Dimensions As RECT          ' bounding rectangle of frame
    Delay As Long               ' length of time (ms) frame stays visible
    TblIndex As Long            ' the ColorTableSTRCUT this frame uses
    imgOffset As Long           ' file byte postion where image begins in gif file
                                ' & after decompressed, position where img begins in DIB strip
    TransIndex As Byte          ' which palette index is to be transparent
    IsTransparent As Byte       ' does frame use transparency (0=no, 1=yes)
    Disposal As Byte            ' the disposal method for this frame (0-7)
End Type
Private Type CoreDCInfo
    DC As Long                  ' off-screen DC
    hDib As Long                ' DIB created for the DC (our DIB strip)
    dibPtr As Long              ' pointer for DIB
    hBmp As Long                ' original monochrome bitmap when DC was created
End Type
Private Type BufferDCInfo
    DC As Long                  ' off-screen DC
    hDib As Long                ' DIB created for the DC
    dibPtr As Long              ' pointer for DIB
    hBmp As Long                ' original monochrome bitmap when DC was created
    hDibBW As Long              ' if monochrome buffer mask is needed, handle to the mask
    dibPtrBW As Long            ' pointer to the monochrome mask
End Type

' Custom Enumerations
Public Enum ScaleGIFConstants    ' See Stretch property. Settings can force control to resize
    gfsClip = 0                  ' will never scale, nor stretch
    gfsScaleAlways = 1           ' will always scale up or down as needed
    gfsStretch = 2               ' will stretch to fit, not scaled, distorted
    gfsShrinkScaleToFit = 3      ' will only scale down if needed else scale is 1:1
    gfsActualSize = 4            ' self-explanatory
End Enum
Public Enum AnimationActions     ' See Action property
    gfaStop = 0                  ' stop on current frame, reset current frame to first
    gfaPlay = 1                  ' start/restart from current frame
    gfaPause = 2                 ' stop on current frame, do not reset frame nr
    gfaForward = 3               ' show next frame only, then pause
    gfaReset = 4                 ' restart from 1st frame
End Enum
Public Enum MirrorConstants     ' See MirrorGIF routine
    gfmNone = 0                 ' no mirroring
    gfmHorizontal = 1           ' mirror horizontally
    gfmVertical = 2             ' mirror vertically
    gfmHorAndVer = 3            ' mirror on X & Y axis
End Enum
Public Enum DelayModeConstants  ' See DelayAnimation property
    gfdNone = 0                 ' entire GIF is built before 1st frame is displayed
    gfdDelayStartup = 1         ' first frame is displayed immediately, then rest are processed
    gfdDoNotAnimate = 2         ' Same as gfdDelayStartup but does not initialize animation
End Enum
    
' All Class-Level variables are prefixed with c_
Private c_OffSetX As Long      ' Optional: modifies the left edge where rendering will occur on control
Private c_OffSetY As Long      ' Optional: modifies the top edge where rendering will occur

Private c_SolidBgkFill As Long  ' optional color override of transparendy to solid fill
Private c_SolidBkgUsed As Boolean
Private c_MinDelay As Integer   ' any delay less than this value will use this value
Private c_AniLoops As Integer   ' during animation: loops remaining
Private c_aniState As AnimationActions ' current animation state
Private c_ScaleMode As ScaleGIFConstants ' image scaling options
Private c_curFrame As Long      ' during animation: which frame is being rendered
Private c_DelayLoad As DelayModeConstants     ' pauses processing during runtime until after first frame is rendered
Private c_Mirror As MirrorConstants ' optional horizontal and/or vertical mirroring
Private c_StepDelay As Long         ' optional overrride of per frame delay, 0 returns to actual delays

Private c_DC As CoreDCInfo          ' holds the decompressed GIF, all frames
Private c_BkBuff As BufferDCInfo    ' holds the backbuffer & Mask if needed

Private c_gifProps As GIFcoreProperties ' overall GIF properties
Private c_Frames() As GIFframeProperties ' collection of individual frame properties
Private c_ColorTables As ColorTableSTRCUT  ' collection of color tables used in the GIF
Private c_maskTable() As Long       ' GIF Mask palette, B&W palette if needed

' following used during decompressing GIF & released immediately
Private c_DataLen() As POINTAPI     ' tracks frames file positions and sizes
Private c_aBuff() As Byte           ' general use byte array
Private c_DIBarray() As Byte        ' another general use array (used in ConvertStripToGIF & BuildDIBstrip)
Private c_aPOT() As Long            ' Power of 2 look up table
Private c_gifData() As Byte         ' source GIF data during design time only
Private c_ClrTblCopy As ColorTableSTRCUT ' back up of color table. See CacheColorTables, RestoreColorTables


' internal timer related variables
Private c_tmrOwner As Long
Private c_Ptr As Long



Public Property Let AnimatedGIF(aGif As StdPicture)
Attribute AnimatedGIF.VB_Description = "返回/设置一个图形显示在一个控制"
    Set AnimatedGIF = aGif
End Property
Public Property Set AnimatedGIF(aGif As StdPicture)
    
    If UserControl.Ambient.UserMode = True Then Exit Property

    Dim tBag As PropertyBag
    Dim sMode As ScaleGIFConstants
    
    If aGif Is Nothing Then
        Erase c_gifData()
    Else
        Set tBag = New PropertyBag
        tBag.WriteProperty "myGIF", aGif    ' 5 char name, don't change it
        ReDim c_gifData(1 To UBound(tBag.Contents) - 53) ' position where GIF data starts
        Call ArrayFromVarRef(tBag.Contents, 54) ' 54 based on property name length
        Set tBag = Nothing
    End If
    If Not LoadGif = 0& Then
        If Ambient.UserMode = False Then
            sMode = c_ScaleMode             ' cache setting
            c_ScaleMode = gfsActualSize     ' force new setting
            Call UserControl_Resize         ' resize if needed
            c_ScaleMode = sMode             ' replace setting
            c_gifProps.ScaleCx = 1          ' set scale ratios
            c_gifProps.ScaleCy = 1
        End If
    End If
    UserControl.Refresh
    PropertyChanged "AnimatedGIF"
End Property
Public Property Get AnimatedGIF() As StdPicture
    If Not c_curFrame = 0& Then Set AnimatedGIF = ConvertStripToGIF(c_curFrame)
End Property

Public Property Get CurrentFrame() As Long
Attribute CurrentFrame.VB_Description = "返回当前帧"
    CurrentFrame = c_curFrame
End Property
Public Property Let Enabled(Enable As Boolean)
    UserControl.Enabled = Enable
    PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，确定对象是否可以响应用户产生的事件"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property
Public Property Let FrameCount(nrFrames As Long)
Attribute FrameCount.VB_Description = "返回数字的GIF内的帧"
End Property
Public Property Get FrameCount() As Long
    If Not c_curFrame = 0& Then FrameCount = UBound(c_Frames)
End Property

Public Property Get ActualWidth() As Long
Attribute ActualWidth.VB_Description = "原始的宽度整体GIF图像"
    ActualWidth = Int(ScaleX(c_gifProps.Width, vbPixels, vbContainerSize))
End Property
Public Property Get ActualHeight() As Long
Attribute ActualHeight.VB_Description = "原高度的整体GIF图像"
    ActualHeight = Int(ScaleY(c_gifProps.Height, vbPixels, vbContainerSize))
End Property

Public Property Let Stretch(newScale As ScaleGIFConstants)
Attribute Stretch.VB_Description = "返回/设置一个值，决定是否调整大小以适应图形图像的大小控制"
    If newScale < gfsClip Or newScale > gfsActualSize Then Exit Property
    If Not c_ScaleMode = newScale Then
        Dim lAction As AnimationActions
        c_ScaleMode = newScale
        lAction = c_aniState    ' cache
        Me.Action = gfaPause    ' pause if needed
        Call UserControl_Resize ' resize depending on c_ScaleMode
        Me.Action = lAction     ' restart animation if needed
        If Not lAction = gfaPlay Then UserControl.Refresh
        PropertyChanged "Stretch"
    End If
End Property
Public Property Get Stretch() As ScaleGIFConstants
    Stretch = c_ScaleMode
End Property

Public Property Let MinFrameDelay(ByVal Delay As Long)
Attribute MinFrameDelay.VB_Description = "最低的毫秒数一帧将保持前被替换。"
    If Delay < 10 Then Delay = 10           ' ensure absolute minimum delay
    If Delay > 32700 Then Delay = 32700     ' ensure absolute maximum delay
    If Not Delay = c_MinDelay Then
        c_MinDelay = Delay
        PropertyChanged "Delay"
    End If
End Property
Public Property Get MinFrameDelay() As Long
    MinFrameDelay = c_MinDelay
End Property

Public Property Let Loops(Count As Long)
Attribute Loops.VB_Description = "The number of loops that will occur before animation stops. Zero will loop infintely"
    If Not c_gifProps.Loops = Count Then
        c_gifProps.Loops = Abs(Count)
        PropertyChanged "Loops"
    End If
End Property
Public Property Get Loops() As Long
    Loops = c_gifProps.Loops
End Property

Public Property Get LoopsRemaining() As Long
Attribute LoopsRemaining.VB_Description = "Returns remaining number of loops before animation terminates"
    If c_curFrame = 0& Then ' no gif loaded
        LoopsRemaining = 0&
    ElseIf c_gifProps.Loops = 0 Then ' infinite
        LoopsRemaining = &H7FFFFFFF ' a really high number
    Else
        LoopsRemaining = c_AniLoops ' loops remaining
    End If
End Sub

Public Property Let Mirrored(MirrorStyle As MirrorConstants)
Attribute Mirrored.VB_Description = "镜图像水平或垂直"
    
    If MirrorStyle < gfmNone Or MirrorStyle > gfmHorAndVer Then Exit Property
    
    If c_curFrame = 0& Then
        c_Mirror = MirrorStyle
        PropertyChanged "Mirrored"
        
    ElseIf Not MirrorStyle = c_Mirror Then
        
        Dim lAction As Long
        lAction = c_aniState                ' cache current state
        Me.Action = gfaPause                ' pause animation
        MirrorGIF MirrorStyle, (Ambient.UserMode = False), Ambient.UserMode
        c_Mirror = MirrorStyle              ' save new value
        If lAction = gfaPlay Then           ' continue animation or refresh
            Me.Action = lAction
        Else
            UserControl.Refresh
        End If
        PropertyChanged "Mirrored"
    
    End If

End Property
Public Property Get Mirrored() As MirrorConstants
    Mirrored = c_Mirror
End Property

Public Property Let Action(Act As AnimationActions)
    If c_curFrame = 0& Then Exit Property
    If UserControl.Ambient.UserMode = False Then Exit Property
    Select Case Act
    Case gfaForward
        ManageTimers 0
        c_curFrame = c_curFrame + 1 ' wrap back to 1st frame if needed
        If c_curFrame > UBound(c_Frames) Then c_curFrame = 1
        c_aniState = gfaPause
        UserControl.Refresh
        RaiseEvent FrameChanged(c_curFrame, False)
    Case gfaPause
        If Not c_aniState = gfaStop Then
            ManageTimers 0
            c_aniState = gfaPause
        End If
    Case gfaPlay
        ManageTimers 1
    Case gfaReset           ' basically a Stop & then Play call
        ManageTimers 0
        c_AniLoops = c_gifProps.Loops ' reset loops remaining
        c_curFrame = 1      ' reset current frame
        UserControl.Refresh
        RaiseEvent FrameChanged(c_curFrame, False)
        ManageTimers 1
    Case gfaStop
        ManageTimers 0
        c_AniLoops = c_gifProps.Loops ' reset loops remaining
        c_curFrame = 1      ' reset current frame
        UserControl.Refresh
        RaiseEvent FrameChanged(c_curFrame, False)
    End Select
End Property
Public Property Get Action() As AnimationActions
Attribute Action.VB_Description = "启动，停止和暂停动画"
Attribute Action.VB_MemberFlags = "400"
    Action = c_aniState
End Property

Public Property Let DelayAnimation(Mode As DelayModeConstants)
Attribute DelayAnimation.VB_Description = "在运行时，动画是暂停，直到第一帧加载和显示"
    If Mode < gfdNone Or Mode > (gfdDelayStartup Or gfdDoNotAnimate) Then Exit Property
    If Not c_DelayLoad = Mode Then
        c_DelayLoad = Mode
        PropertyChanged "DelayAnimation"
    End If
End Property
Public Property Get DelayAnimation() As DelayModeConstants
    DelayAnimation = c_DelayLoad
End Property

Public Property Let MouseIcon(MousePic As StdPicture)
Attribute MouseIcon.VB_Description = "设置一个自定义鼠标图标"
    Set MouseIcon = MousePic
End Property
Public Property Set MouseIcon(MousePic As StdPicture)
    On Error Resume Next
    Set UserControl.MouseIcon = MousePic
    If MousePic Is Nothing Then UserControl.MousePointer = vbDefault
    PropertyChanged "MouseIcon"
End Property
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MousePointer(Pointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "返回/设置鼠标指针的类型时，显示了一个对象的一部分"
    ' same as VB's MousePointer property
    On Error Resume Next
    UserControl.MousePointer = Pointer
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let OffsetX(NewOffset As Long)
Attribute OffsetX.VB_Description = "离开调整个像素渲染帧的边缘"
    c_OffSetX = NewOffset
End Property
Public Property Get OffsetX() As Long
    OffsetX = c_OffSetX
    If Ambient.UserMode = False Or c_aniState <> gfaPlay Then UserControl.Refresh
End Property

Public Property Let OffsetY(NewOffset As Long)
Attribute OffsetY.VB_Description = "顶部边缘的调整提供n个像素帧"
    c_OffSetY = NewOffset
    If Ambient.UserMode = False Or c_aniState <> gfaPlay Then UserControl.Refresh
End Property
Public Property Get OffsetY() As Long
    OffsetY = c_OffSetY
End Property

Public Property Get SolidBkgColor() As OLE_COLOR
Attribute SolidBkgColor.VB_Description = "透明的GIF会因在这一背景颜色"
    SolidBkgColor = c_SolidBgkFill
End Property
Public Property Let SolidBkgColor(ByVal Color As OLE_COLOR)
    c_SolidBgkFill = Color
    If c_SolidBkgUsed = True And Ambient.UserMode = False Then UserControl.Refresh
End Property
Public Property Get SolidBkgColorUsed() As Boolean
Attribute SolidBkgColorUsed.VB_Description = "使用固体Bkg Color属性，否则忽略它"
    SolidBkgColorUsed = c_SolidBkgUsed
End Property
Public Property Let SolidBkgColorUsed(ByVal IgnoreTransparency As Boolean)
    c_SolidBkgUsed = IgnoreTransparency
    If Ambient.UserMode = False Then UserControl.Refresh
End Property

Public Property Let SteppedDelay(ByVal newDelay As Long)
Attribute SteppedDelay.VB_Description = "对所有的帧由提供的价值呈现。零延迟擦除阶梯"
    If newDelay < 1& Then
        c_StepDelay = 0&
    Else
        c_StepDelay = newDelay
    End If
End Property
Public Property Get SteppedDelay() As Long
    SteppedDelay = c_StepDelay
End Property

Public Function LoadAnimatedGIF_File(ByVal filename As String) As Boolean
Attribute LoadAnimatedGIF_File.VB_Description = "Optional method to assign a GIF"
    ' loads a GIF by file name during runtime
    Dim lRtn As Long, bUnicode As Boolean
    
    On Error Resume Next
    
    Const INVALID_HANDLE_VALUE = -1&
    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&

    ' test to see if a file exists
    If isStringANSI(filename) Then
        lRtn = GetFileAttributes(filename)
    Else
        bUnicode = True
        lRtn = GetFileAttributesW(StrPtr(filename))
    End If
    
    If Not lRtn = INVALID_HANDLE_VALUE Then
        
        Dim flags As Long, Access As Long
        Dim Disposition As Long, Share As Long
        Dim hFile As Long
    
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    
        If bUnicode Then hFile = CreateFileW(StrPtr(filename), Access, Share, 0&, Disposition, flags, 0&)
        If hFile = 0 Then
            hFile = CreateFile(filename, Access, Share, 0&, Disposition, flags, 0&)
        End If
        If Not hFile = -1 Then
            lRtn = GetFileSize(hFile, 0&)
            If lRtn Then
                SetFilePointer hFile, 0&, 0&, 0&
                ReDim c_gifData(1 To lRtn)
                ReadFile hFile, c_gifData(1), lRtn, flags, ByVal 0&
                If Not flags = lRtn Then Erase c_gifData()
            End If
            CloseHandle hFile
        End If
    End If
    If Not LoadGif() = 0& Then
        If c_ScaleMode = gfsActualSize Or c_ScaleMode = gfsScaleAlways Or c_ScaleMode = gfsShrinkScaleToFit Then
            Call UserControl_Resize
        End If
        LoadAnimatedGIF_File = True
    End If
End Function

Public Sub LoadAnimatedGIF_Remote(ByVal RemotePath As String)
Attribute LoadAnimatedGIF_Remote.VB_Description = "Optional method to assign a GIF"
    If Len(RemotePath) = 0 Then Exit Sub
    On Error Resume Next
    UserControl.AsyncRead RemotePath, vbAsyncTypeByteArray, "remoteGIF", vbAsyncReadGetFromCacheIfNetFail
    If ERR Then
        ERR.Clear
        RaiseEvent RemoteLoadFailure
    End If
    
End Sub

Public Function LoadAnimatedGIF_Array(Bytes() As Byte) As Boolean
Attribute LoadAnimatedGIF_Array.VB_Description = "可选的方法分配一个GIF"
    Dim lRtn As Long, SA As SAFEARRAY2D
    On Error GoTo EH
    lRtn = IsArrayEmpty(VarPtrArray(Bytes))
    If Not lRtn = 0 Then
        CopyMemory SA, ByVal lRtn, 2&
        If SA.cDims < 3 Then
            CopyMemory SA, ByVal lRtn, 16 + (SA.cDims * 8)
            If SA.cDims = 1 Then
                ReDim c_gifData(1 To (SA.cbElements * SA.rgSABound(0).cElements))
            Else
                ReDim c_gifData(1 To (SA.cbElements * (SA.rgSABound(0).cElements * SA.rgSABound(1).cElements)))
            End If
            CopyMemory c_gifData(1), ByVal SA.pvData, UBound(c_gifData)
        End If
    End If
    If Not LoadGif() = 0& Then
        If c_ScaleMode = gfsActualSize Or c_ScaleMode = gfsScaleAlways Or c_ScaleMode = gfsShrinkScaleToFit Then
            ' resize to user-defined scaled ratios
            Call UserControl_Resize
        End If
        LoadAnimatedGIF_Array = True
    End If
EH:
End Function

Public Function GetPalette(ByVal FrameIndex As Long, paletteEntries() As Long) As Boolean
Attribute GetPalette.VB_Description = "返回调色板的选择帧作品"
    If c_curFrame = 0& Then Exit Function
    
    Dim bRestart As Boolean
    If Not FrameIndex < 0 Then
    
        If c_aniState = gfaPlay Then
            Me.Action = gfaPause
            bRestart = True
        End If
    
        If FrameIndex = 0 Then
            ReDim paletteEntries(0 To c_ColorTables.Tables(PALETTECOUNT, 0) - 1)
            CopyMemory paletteEntries(0), c_ColorTables.Tables(0, 0), (UBound(paletteEntries) + 1) * 4&
            GetPalette = True
        ElseIf FrameIndex <= UBound(c_Frames) Then
            If Not c_Frames(FrameIndex).TblIndex = 0 Then
                ReDim paletteEntries(0 To c_ColorTables.Tables(PALETTECOUNT, c_Frames(FrameIndex).TblIndex) - 1)
                CopyMemory paletteEntries(0), c_ColorTables.Tables(0, c_Frames(FrameIndex).TblIndex), (UBound(paletteEntries) + 1) * 4&
                GetPalette = True
            End If
        End If
        If bRestart Then Me.Action = gfaPlay
    End If
End Function

Public Function SetPalette(ByVal FrameIndex As Long, paletteEntries() As Long) As Boolean
Attribute SetPalette.VB_Description = "为选定的帧作品集调色板"
        
    If c_curFrame = 0& Then Exit Function
    If IsArrayEmpty(VarPtrArray(paletteEntries)) = 0& Then Exit Function
    
    Dim bRestart As Boolean
    If Not FrameIndex < 0 Then
        
        If c_aniState = gfaPlay Then
            Me.Action = gfaPause
            bRestart = True
        End If
        On Error Resume Next
        Dim UB As Long, LB As Long
        
        LB = LBound(paletteEntries)
        UB = UBound(paletteEntries)
        If ERR Then
            ERR.Clear
        ElseIf FrameIndex = 0 Then
            If Abs(UB - LB) + 1 >= c_ColorTables.Tables(PALETTECOUNT, 0) Then
                CopyMemory c_ColorTables.Tables(0, 0), paletteEntries(LB), c_ColorTables.Tables(PALETTECOUNT, 0) * 4&
                SetPalette = True
            End If
        ElseIf FrameIndex <= UBound(c_Frames) Then
            If Not c_Frames(FrameIndex).TblIndex = 0 Then
                If Abs(UB - LB) + 1 >= c_ColorTables.Tables(PALETTECOUNT, c_Frames(FrameIndex).TblIndex) Then
                    CopyMemory c_ColorTables.Tables(0, c_Frames(FrameIndex).TblIndex), paletteEntries(LB), c_ColorTables.Tables(PALETTECOUNT, c_Frames(FrameIndex).TblIndex) * 4&
                    SetPalette = True
                End If
            End If
        End If
        If bRestart Then
            Me.Action = gfaPlay
        Else
            UserControl.Refresh
        End If
    End If
    
End Function

Public Sub Refresh()
Attribute Refresh.VB_Description = "对一个对象的完整重绘"
    UserControl.Refresh
End Sub

Public Function CacheColorTables() As Boolean
Attribute CacheColorTables.VB_Description = "Have control save GIF color tables inpre paration of callig Set Palette and laterresto ring original palette"
    If c_ClrTblCopy.Index = 0 Then
        c_ClrTblCopy = c_ColorTables
        c_ClrTblCopy.Index = 1  ' flag indicating already cached
        CacheColorTables = True
    End If
End Function

Public Function RestoreColorTables(bReleaseCache As Boolean) As Boolean
Attribute RestoreColorTables.VB_Description = "色彩还原的GIF缓存表，从表缓存的颜色"
    If c_ClrTblCopy.Index Then
        c_ColorTables = c_ClrTblCopy
        c_ColorTables.Index = -1 ' forces next frame to validate color table
        If bReleaseCache Then
            Erase c_ClrTblCopy.Tables()
            c_ClrTblCopy.Index = 0  ' flag indicating no cache exists
        End If
        RestoreColorTables = True
    End If
End Function
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    Dim bLoad As Boolean
    If AsyncProp.BytesRead < 15 Then
        RaiseEvent RemoteLoadFailure
    Else
        Dim cx As Long, cy As Long
        ReDim c_gifData(1 To 10)
        ArrayFromVarRef AsyncProp.Value, LBound(AsyncProp.Value)
        Select Case Left$(LCase(StrConv(c_gifData, vbUnicode)), 6)
            Case "gif89a", "gif87a"
                CopyMemory cx, c_gifData(7), 2&
                CopyMemory cy, c_gifData(9), 2&
                RaiseEvent RemoteLoadComplete(Int(ScaleX(cx, vbPixels, vbContainerSize)), Int(ScaleY(cy, vbPixels, vbContainerSize)), bLoad)
                bLoad = Not bLoad
            Case Else
                RaiseEvent RemoteLoadFailure
        End Select
        Erase c_gifData()
    End If
    If bLoad = True Then
        ReDim c_gifData(1 To AsyncProp.BytesRead)
        ArrayFromVarRef AsyncProp.Value, LBound(AsyncProp.Value)
        If Not LoadGif() = 0& Then
            If c_ScaleMode = gfsActualSize Or c_ScaleMode = gfsScaleAlways Or c_ScaleMode = gfsShrinkScaleToFit Then
                Call UserControl_Resize
            End If
        End If
    End If

End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    If AsyncProp.StatusCode = vbAsyncStatusCodeError Then
        CancelAsyncRead AsyncProp.PropertyName
        RaiseEvent RemoteLoadFailure
    End If

End Sub

Private Sub UserControl_Hide()
    Me.Action = gfaPause
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit  ' allows selecting image during design time
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEMOVE(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MOUSEUP(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_Paint()

    RenderFrame c_curFrame, UserControl.hdc
    If Ambient.UserMode = False Then
        Rectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
End Sub

Private Sub UserControl_Initialize()
    c_MinDelay = 50
    c_ScaleMode = gfsScaleAlways
    c_DelayLoad = gfdDelayStartup
    UserControl.DrawStyle = vbDot
    UserControl.FOREColor = vbWindowFrame
End Sub

Private Sub UserControl_Terminate()
    UnloadGIF
    zTerminate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        c_gifData = .ReadProperty("GIF", c_gifData())
        c_MinDelay = .ReadProperty("Delay", c_MinDelay)
        c_ScaleMode = .ReadProperty("Stretch", gfsScaleAlways)
        c_DelayLoad = .ReadProperty("DelayLoad", gfdDelayStartup)
        If c_DelayLoad = -1 Then c_DelayLoad = gfdDelayStartup
        c_Mirror = .ReadProperty("Mirror", gfmNone)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        c_SolidBgkFill = .ReadProperty("BkgFill", 0&)
        c_SolidBkgUsed = .ReadProperty("BkgUsed", False)
        c_StepDelay = .ReadProperty("StepDelay", 0&)
        c_OffSetX = .ReadProperty("OffsetX", 0&)
        c_OffSetY = .ReadProperty("OffsetY", 0&)
    End With
    If Not LoadGif() = 0& Then
    
        c_gifProps.Loops = PropBag.ReadProperty("Loops", c_gifProps.Loops)
        c_AniLoops = c_gifProps.Loops
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "GIF", c_gifData()
        .WriteProperty "Delay", c_MinDelay, 50
        .WriteProperty "Stretch", c_ScaleMode, gfsScaleAlways
        .WriteProperty "Loops", c_gifProps.Loops, 0
        .WriteProperty "DelayLoad", c_DelayLoad, gfdDelayStartup
        .WriteProperty "Mirror", c_Mirror, gfmNone
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "OffsetX", c_OffSetX, 0&
        .WriteProperty "OffsetY", c_OffSetY, 0&
        .WriteProperty "BkgFill", c_SolidBgkFill, 0&
        .WriteProperty "BkgUsed", c_SolidBkgUsed, False
        .WriteProperty "StepDelay", c_StepDelay, 0&
    End With
End Sub

Private Sub UserControl_Resize()

    If Not c_curFrame = 0& Then
        Dim cx As Long, cy As Long
        If c_gifProps.Height < 0 Then Exit Sub
        cx = UserControl.ScaleWidth
        cy = UserControl.ScaleHeight
        ScaleToDestination 0&, 0&, cx, cy, c_gifProps.ScaleCx, c_gifProps.ScaleCy
        c_gifProps.Height = -c_gifProps.Height
        UserControl.Size ScaleX(cx, vbPixels, vbTwips), ScaleY(cy, vbPixels, vbTwips)
        c_gifProps.Height = -c_gifProps.Height
    End If
    
End Sub

Private Sub UserControl_Show()
    If c_aniState = gfaPause Then ManageTimers 1
End Sub

Private Sub ArrayFromVarRef(inArray() As Byte, offset As Long)
    CopyMemory c_gifData(1), inArray(offset), UBound(c_gifData)
End Sub

Private Function IsArrayEmpty(ByVal FarPointer As Long) As Long
  CopyMemory IsArrayEmpty, ByVal FarPointer, 4&
End Function

Private Function ByteAlignOnWord(ByVal BitDepth As Byte, ByVal Width As Long) As Long
    ByteAlignOnWord = (((Width * BitDepth) + &H1F) And Not &H1F&) \ &H8
End Function

Private Function LoadGif() As Long
On Error GoTo EH
    Dim newProps As GIFcoreProperties
    Dim nrItems As Long
    
    Call UnloadGIF
    c_gifProps = newProps
    
    If Not IsArrayEmpty(VarPtrArray(c_gifData)) = 0& Then
    
           ReDim c_aPOT(0 To 8)
        c_aPOT(0) = 1
        For nrItems = 1 To 8
            c_aPOT(nrItems) = c_aPOT(nrItems - 1) * 2
        Next
        
        nrItems = ParseGIF()
        
        On Error GoTo EH
        If Not nrItems = 0 Then
            Erase c_aBuff()
            c_Ptr = 0
            If BuildDIBstrip(True) = False Then
                nrItems = 0
            Else
                If Ambient.UserMode = True Then
                    If nrItems = 1 Then
                        UserControl.Refresh
                        RaiseEvent FrameChanged(nrItems, False)
                    End If
                End If
            End If
        End If
    End If

EH:
    If nrItems = 0 Then
        If ERR Then ERR.Clear
        Erase c_gifData
        'UnloadGIF
    End If
    LoadGif = nrItems
    
End Function

Private Sub UnloadGIF()
    On Error Resume Next
    ManageTimers 0&
    Erase c_Frames()
    c_aniState = gfaStop
    c_curFrame = 0&
    c_Ptr = 0&
    
    With c_DC
        If Not .hDib = 0& Then
            DeleteObject SelectObject(.DC, .hBmp)
            .hDib = 0&
            .hBmp = 0&
        End If
        If Not .DC = 0& Then
            DeleteDC .DC
            .DC = 0&
        End If
    End With
    With c_BkBuff
        If Not .hDib = 0& Then
            DeleteObject SelectObject(.DC, .hBmp)
            .hDib = 0&
            .hBmp = 0&
        End If
        If Not .hDibBW = 0& Then
            DeleteObject .hDibBW
            .hDibBW = 0&
        End If
        If Not .DC = 0& Then
            DeleteDC .DC
            .DC = 0&
        End If
    End With
  
    Erase c_ColorTables.Tables()
    Erase c_ClrTblCopy.Tables()
    Erase c_maskTable()
    Erase c_aPOT()
    Erase c_DataLen()
    c_ColorTables.Index = 0&
    c_ClrTblCopy.Index = 0&
    zTerminate
    If ERR Then ERR.Clear
End Sub

Private Function ParseGIF() As Long

    On Error Resume Next
    ' a modified routine from some of my other GIF postings
    ' This one is scaled back and skips many GIF blocks not needed for displaying.
    ' It also tracks a few file positions for later use in the BuildDIBstrip routine
    
    Dim gByte As Byte                       ' general purpose Byte
    Dim gLong As Long                       ' general purpose Long
    Dim gString As String                   ' general purpose String
    Dim lFrameCount As Integer              ' nr frames in GIF
    Dim bGlobalColorTable As Boolean        ' does a global exist; if not, then possibly corrupt gif
    Dim cFrame As GIFframeProperties        ' frame data
    Dim emptyFrame As GIFframeProperties    ' overall GIF data
    
    On Error GoTo ExitReadRoutine
    
    ReDim c_DataLen(1 To 10)  ' used to track start position of image description & size of compressed data
    c_Ptr = 1
    With c_gifProps
    
        ' read signature
        Call ReadGifFile_Variable(6)
        gString = LCase(StrConv(c_aBuff, vbUnicode))
        If Not (gString = "gif89a" Or gString = "gif87a") Then Exit Function
        
        ' skip to the global color table information
        c_Ptr = 11
        gByte = ReadGifFile_Byte()             ' the packed byte
        ' GIF Logical Screen Descriptor packed byte per specs
            ' bit pos = 0: nr of bits = 1 ' global color table used
            ' bit pos = 1: nr of bits = 3 ' original resolution
            ' bit pos = 4: nr of bits = 1 ' table sorted
            ' bit pos = 5: nr of bits = 3 ' palette bit depth (can be 0 thru 7 inclusive)
        
        c_Ptr = 14  ' move ahead to next byte after header
        If (gByte And 128) = 128 Then ' color table used. If so, read it
            c_ColorTables.Index = 0
            Call ReadGifFile_ColorTable(0, (gByte And &H7) + 1)
            bGlobalColorTable = True
        'Else no global color table; probably uses local color tables
        End If
    
    End With
    
    ' scan thru the entire file to find all the images & other key data
    With c_gifProps
    
        Do
            Select Case ReadGifFile_Byte()    ' read a single byte
            Case 0  ' block terminators
            Case 33 'Extension Introducer
                Select Case ReadGifFile_Byte() ' read the extension type
                
                Case 255    ' application extension, look for a loop count
                    ' the Netscape extension should always be before any images
                    ' Get the length of extension: will always be 11
                    gByte = ReadGifFile_Byte()
                    ' read next 11 bytes & see if it is a netscape extension
                    Call ReadGifFile_Variable(gByte)
                    gString = UCase(StrConv(c_aBuff, vbUnicode))
                    If gString = "NETSCAPE2.0" Then
                        ' within the data, we can extract the number of loops
                        ' an animated gif is suppose to run.
                        gByte = ReadGifFile_Byte()
                        If gByte = 3 Then   ' valid netscape extension
                            ' move ahead one byte & the next two is the loop count
                            c_Ptr = c_Ptr + 1
                            .Loops = (ReadGifFile_Integer And &HFFFF&) ' convert unsigned Integer to Long
                        Else
                            c_Ptr = c_Ptr - 1
                        End If
                    End If
                    SkipGifBlock
                
                Case 249    ' Graphic Control Label
                            ' (description of frame & is an optional block) 8 bytes
                    ' Graphic Control Extension (packed byte)
                        'bit pos = 0: nr of bits = 3 ' reserved
                        'bit pos = 3: nr of bits = 3 ' disposal method
                        'bit pos = 6: nr of bits = 1 ' user input flag
                        'bit pos = 7: nr of bits = 1 ' transparency
                    
                    cFrame = emptyFrame ' begin a new frame structure
                    With cFrame
                        .imgOffset = c_Ptr - 2 ' image starts here in the file
                        gByte = ReadGifFile_Byte() ' skip static byte (block size of 4)
                        ' get next byte which contains several properties
                        gByte = ReadGifFile_Byte()
                        
                        ' how is frame cleared after it is shown
                        .Disposal = ((gByte \ &H4) And &H3)
                        If .Disposal = 3 Then
                            If lFrameCount = 0& Then .Disposal = lFrameCount
                        ElseIf .Disposal > 3 Then
                            .Disposal = 0   ' treat reserved disposal codes as zero
                        End If
                        
                        ' how long does frame stay before being disposed & make 1/1000sec vs 1/100sec
                        .Delay = (ReadGifFile_Integer And &HFFFF&) * 10 ' convert unsigned Integer to Long
                        
                        .IsTransparent = (gByte And &H1) ' has transparency.
                        If .IsTransparent = 1 Then
                            .TransIndex = ReadGifFile_Byte  ' transparent index
                        Else
                            c_Ptr = .imgOffset + 7 ' skip transindex byte
                        End If
                        ' skip rest of block
                        Call SkipGifBlock
                    End With
                    
                Case Else   ' Comment block, plain text extension, or Unknown extension
                    Call SkipGifBlock
                End Select
                
            Case 44 ' Image Descriptor (image dimensions & color table)
                ' location of image within logical window
                ' GIF Image Descriptor (packed byte)
                    'bit pos = 0: nr of bits = 1 ' local color table used
                    'bit pos = 1: nr of bits = 1 ' interlaced
                    'bit pos = 2: nr of bits = 1 ' table sorted
                    'bit pos = 3: nr of bits = 2 ' reserved
                    'bit pos = 5: nr of bits = 3 ' palette bit depth
                With cFrame
                    ' mark position where image description starts
                    c_DataLen(lFrameCount + 1).Y = c_Ptr - 1
                    ' does image start here or did it start with Block 249 above.
                    If .imgOffset = 0 Then .imgOffset = c_DataLen(lFrameCount + 1).Y
                    
                    ' convert unsigned integers to long
                    .Dimensions.Left = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Top = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Right = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Bottom = (ReadGifFile_Integer And &HFFFF&)
                    
                    ' next byte indicates if local color table used
                    gByte = ReadGifFile_Byte()
                    If (gByte And 128) = 128 Then   ' local color table used.
                        .TblIndex = c_ColorTables.Index + 1 ' update ref & read table
                        Call ReadGifFile_ColorTable(.TblIndex, (gByte And &H7) + 1)
                        c_ColorTables.Index = .TblIndex     ' update table count
                    Else
                        If bGlobalColorTable = False Then Exit Function ' corrupted gif
                        ' frame says to use a global table, but no global color table exists
                    End If
                    ' skip the LZW byte & move to 1st byte of image
                    gByte = ReadGifFile_Byte()
                    SkipGifBlock ' done with block
                End With
                
                ' calculate image size in compressed bytes (includes local table if used)
                lFrameCount = lFrameCount + 1
                c_DataLen(lFrameCount).X = c_Ptr - c_DataLen(lFrameCount).Y
                If c_DataLen(lFrameCount).X < 3 Then ' then invalid image data
                    lFrameCount = lFrameCount - 1 ' roll back the frame
                Else
                    ReDim Preserve c_Frames(1 To lFrameCount)
                    If lFrameCount = UBound(c_DataLen) Then
                        ReDim Preserve c_DataLen(1 To lFrameCount + 5)
                    End If
                    c_Frames(lFrameCount) = cFrame
                    ' ensure logical window large enough for every frame
                    With c_Frames(lFrameCount).Dimensions
                        If .Left + .Right > c_gifProps.Width Then c_gifProps.Width = .Left + .Right
                        If .Top + .Bottom > c_gifProps.Height Then c_gifProps.Height = .Top + .Bottom
                    End With
                End If
                cFrame = emptyFrame ' always start with a new frame
                
            Case 59 ' Trailer (end of images)
                ' Although more images may exist, this flag tells us not to use any others
                Exit Do
            Case Else
                ' shouldn't happen; abort with what we have
                Exit Do
            End Select
        Loop
          
    End With
    ' got this far. good to go
          
ExitReadRoutine:
    If ERR Then ERR.Clear
    If Not lFrameCount = 0 Then
        If lFrameCount > UBound(c_Frames) Then ReDim Preserve c_Frames(1 To lFrameCount)
    End If
    ParseGIF = lFrameCount
    
End Function

Private Sub SkipGifBlock()
    For c_Ptr = c_Ptr To UBound(c_gifData)
        If c_gifData(c_Ptr) = 0 Then Exit For
        c_Ptr = c_Ptr + c_gifData(c_Ptr)
    Next
    c_Ptr = c_Ptr + 1
End Sub
Private Function ReadGifFile_Byte() As Byte

    If c_Ptr > UBound(c_gifData) Then
        ERR.Raise 53, "ReadGifFile", "End of File"
        Exit Function
    End If
    ReadGifFile_Byte = c_gifData(c_Ptr)
    c_Ptr = c_Ptr + 1
End Function
Private Function ReadGifFile_Integer() As Integer
    If c_Ptr + 1 > UBound(c_gifData) Then
        ERR.Raise 53, "ReadGifFile", "End of File"
        Exit Function
    End If
    CopyMemory ReadGifFile_Integer, c_gifData(c_Ptr), 2&
    c_Ptr = c_Ptr + 2
End Function
Private Sub ReadGifFile_Variable(ByVal nrBytes As Long)

    ReDim c_aBuff(0 To nrBytes - 1)
    If c_Ptr + nrBytes - 1 <= UBound(c_gifData) Then
        CopyMemory c_aBuff(0), c_gifData(c_Ptr), nrBytes
        c_Ptr = c_Ptr + nrBytes
    End If
EH:
End Sub
Private Sub ReadGifFile_ColorTable(ByVal TableSlot As Long, ByVal BitDepth As Long)
    Dim C As Long
    ReDim Preserve c_ColorTables.Tables(0 To PALETTECOUNT, 0 To TableSlot)
    If c_Ptr + c_aPOT(BitDepth) * 3 > UBound(c_gifData) Then
        ERR.Raise 53, "ReadGifFile", "End of File"
        Exit Sub
    End If
    For C = 0 To c_aPOT(BitDepth) - 1 ' convert RGB to BGR along the way
        c_ColorTables.Tables(C, TableSlot) = (c_gifData(c_Ptr) * &H10000) Or (c_gifData(c_Ptr + 1) * &H100&) Or c_gifData(c_Ptr + 2)
        c_Ptr = c_Ptr + 3
    Next
    ' cache number of palette entries used by this table
    c_ColorTables.Tables(PALETTECOUNT, TableSlot) = c_aPOT(BitDepth)
End Sub


Private Function BuildDIBstrip(InitialLoad As Boolean) As Boolean
    Dim f As Long, aPtr As Long
    Dim frameStart As Long, frameStop As Long
    Dim uniquePal(0 To 767) As Byte ' used to ensure palette index uniqueness
    Dim stripBMP As BITMAPINFO
    Dim tPic As StdPicture
    
    On Error GoTo EH
    If InitialLoad Then   ' first time thru, we need to do some things
        Dim maxDataLen As Long, maxHeight As Long, totalWidth As Long
        Dim bTransparency As Byte, tDC As Long
        Dim bEraseAll As Boolean, bNeedMask As Boolean
        bEraseAll = True
        For f = 1 To UBound(c_Frames)
            With c_Frames(f)
                ' add the widths of each frame - this will be our DIB's overall width
                ' and may be less than the GIF's overall width
                totalWidth = totalWidth + .Dimensions.Right
                ' keep track of largest frame height - this will be our DIB's overall height
                ' and may be less than the GIF's overall height
                If .Dimensions.Bottom > maxHeight Then maxHeight = .Dimensions.Bottom
                 ' transparency determines mask frame and mask color table creation
                bTransparency = bTransparency Or .IsTransparent
                ' guesstimate size of byte stream needed to create stdPic GIFs
                If c_DataLen(f).X > maxDataLen Then maxDataLen = c_DataLen(f).X
                Select Case .Disposal
                Case 0, 1
                    bEraseAll = False   ' may need a back buffer
                Case 2
                    bNeedMask = True    ' needs back buffer & mask unless all frames are code 2
                Case 3
                    bEraseAll = False   ' definitely needs back buffer
                    bNeedMask = True    ' definitely needs a mask
                End Select
            End With
        Next

        If UserControl.Ambient.UserMode = False Or f = 2 Then
            totalWidth = c_Frames(1).Dimensions.Right
            maxHeight = c_Frames(1).Dimensions.Bottom
            bNeedMask = False   ' never need a buffer or mask for single frames
        Else
            If bEraseAll = True Then
                bNeedMask = False   ' all frames are erased after drawn; no buffer or mask needed
            Else
                ' combination of disposal codes require a buffer and maybe a mask
                If bNeedMask = False Then
                    ' double check mask necessity
                    If c_Frames(1).IsTransparent = 1 Then
                        bNeedMask = True ' when 1st frame has transparency, mask is always needed
                    ElseIf bTransparency = 1 Then
                        ' 1st frame is not transparent, but if other frames are, then we
                        ' will need a mask if the 1st frame isn't same size of entire gif
                        If Not c_Frames(1).Dimensions.Bottom = c_gifProps.Height Then
                            bNeedMask = True
                        ElseIf Not c_Frames(1).Dimensions.Right = c_gifProps.Width Then
                            bNeedMask = True
                        End If
                    End If
                End If
                ' only possible error here would be lack of system resources
                If BuildBackBuffer(bNeedMask) = False Then Exit Function
            End If
        End If
        With stripBMP.bmiHeader            '  create the color dib strip
            .biBitCount = 8
            .biClrUsed = PALETTECOUNT
            .biHeight = maxHeight
            .biWidth = totalWidth
            .biPlanes = 1
            .biSize = 40
        End With
        With stripBMP                       ' fix to address problem discussed at top
            For f = 1 To PALETTECOUNT - 1&  ' ensure each palette entry is used only once
                uniquePal(f * 3& + 2&) = f  ' RGB (byte) - used by GIF stream
                .bmiPalette(f) = f          ' BGR (long) - used by DIB
            Next
        End With
        
        tDC = GetDC(0&)
        ' create a DC & select our newly created strip into it
        c_DC.DC = CreateCompatibleDC(tDC)
        If Not c_DC.DC = 0& Then
            c_DC.hDib = CreateDIBSection(tDC, stripBMP, 0&, c_DC.dibPtr, 0&, 0&) ' create the DIB
            If c_DC.hDib = 0& Then
                DeleteDC c_DC.DC
                c_DC.DC = 0&
            End If
        End If
        ReleaseDC 0&, tDC
        If (c_DC.DC = 0& Or c_DC.hDib = 0&) Then Exit Function
        c_DC.hBmp = SelectObject(c_DC.DC, c_DC.hDib)
        
        ' create an all-black mask array
        If bNeedMask = True Or bTransparency = 1 Then ReDim c_maskTable(0 To PALETTECOUNT - 1)
        
        ReDim c_DIBarray(1 To maxDataLen + 790) ' oversize array to handle any frame

        CopyMemory c_DIBarray(1), c_gifData(1), 13&    ' copy logical window into stream
        CopyMemory c_DIBarray(14), uniquePal(0), PALETTECOUNT * 3& ' copy our unique palette into the global
        c_DIBarray(11) = c_DIBarray(11) Or 135        ' 135=our global exists & is 256 colors
        frameStart = 1&: frameStop = 1&
    
    Else    ' 2nd time thru
        
        If Not UBound(c_Frames) = 1& Then
            For f = 1& To PALETTECOUNT - 1& ' ensure each palette entry is used only once
                uniquePal(f * 3& + 2&) = f  ' RGB (byte) - used by GIF stream
                stripBMP.bmiPalette(f) = f  ' BGR (long) - used by DIB
            Next
            SetDIBColorTable c_DC.DC, 0, 256, stripBMP.bmiPalette(0)
        End If
        frameStart = 2&: frameStop = UBound(c_Frames)
        
    End If

    c_ColorTables.Index = -1&
    For f = frameStart To frameStop
        
        aPtr = 782&
        With c_Frames(f)
            
            If Not c_DataLen(f).Y = .imgOffset Then

                CopyMemory c_DIBarray(aPtr), c_gifData(.imgOffset), 8&
                c_DIBarray(aPtr + 3&) = (c_DIBarray(aPtr + 3&) And Not 1)

                aPtr = aPtr + 8&
            End If
            CopyMemory c_DIBarray(aPtr), c_gifData(c_DataLen(f).Y), c_DataLen(f).X
            CopyMemory c_DIBarray(aPtr + 1), 0&, 4&
            CopyMemory c_DIBarray(7), c_DIBarray(aPtr + 5&), 4&
            If Not .TblIndex = 0& Then
               
                CopyMemory c_DIBarray(aPtr + 10&), uniquePal(0), _
                            c_aPOT((c_DIBarray(aPtr + 9&) And &H7) + 1&) * 3&
            End If
            aPtr = aPtr + c_DataLen(f).X
            c_DIBarray(aPtr) = 59
            
      
            If f = 1& Then .imgOffset = 0& Else _
                .imgOffset = c_Frames(f - 1&).imgOffset + c_Frames(f - 1&).Dimensions.Right
            

            Set tPic = PictureFromByteStream(c_DIBarray(), aPtr)
            If Not tPic Is Nothing Then
  
                tPic.Render c_DC.DC + 0&, .imgOffset + 0&, 0&, .Dimensions.Right + 0&, .Dimensions.Bottom + 0&, 0&, tPic.Height, tPic.Width, -tPic.Height, ByVal 0&

            End If
        End With
        
    Next
    
    If frameStart = 1& Then
        c_curFrame = 1&
        ScaleToDestination 0&, 0&, UserControl.ScaleWidth + 0&, UserControl.ScaleHeight + 0&, c_gifProps.ScaleCx, c_gifProps.ScaleCy
        If UserControl.Ambient.UserMode = True Then
            If Not (c_DelayLoad = gfdNone) Then
                If UBound(c_Frames) = 1& Then
                    InitialLoad = False ' single frame; no secondary recursion needed
                Else
                    MirrorGIF gfmNone, True, False
                    c_Mirror = c_Mirror Or &H20&
                    ManageTimers 2&  ' pause a few ms to let other processes continue
                End If
            Else
                BuildDIBstrip False ' continue building additional frames
            End If
        Else    ' done with our GIF, reset
            InitialLoad = False
        End If
    End If
    ' second time thru and/or finished processing remaining frames
    If InitialLoad = False Then ' clean up & animate as needed
        Erase c_DIBarray
        Erase c_aPOT
        Erase c_DataLen
        MirrorGIF gfmNone, (Ambient.UserMode = False), False
        If UserControl.Ambient.UserMode Then
            Erase c_gifData
            If (c_DelayLoad And gfdDoNotAnimate) = gfdNone Then
                c_AniLoops = c_gifProps.Loops
                ManageTimers 1& ' establish animation timer
            Else
                RaiseEvent FrameChanged(1, False)
            End If
        End If
    End If
    
    BuildDIBstrip = True
EH:
    If ERR Then
'        Stop            ' troublehsooting only
'        Resume
        ERR.Clear
    End If
End Function

Private Function RenderFrame(Index As Long, hdc As Long) As Boolean

 ' Function renders a frame as a result of the class timer firing or a container repaint
 
    If Index = 0 Then Exit Function
    ' ^^ visible control without image can get paint events; nothing to draw
 
    Dim drawRect As RECT                   ' frame's bounding rect
    Dim bStretch As Boolean
    Dim d3Mask() As Byte, d3Color() As Byte ' used to cache current buffer contents as needed
    Dim hBrush As Long
    
    Const Ratio1to1 As Single = 1
    
    If Not c_gifProps.ScaleCx = Ratio1to1 Then
        bStretch = True
    ElseIf Not c_gifProps.ScaleCy = Ratio1to1 Then
        bStretch = True
    End If
    
    If c_BkBuff.hDib = 0& Then  ' no offscreen buffer, no mask
       ' this is the easy render method
       With c_Frames(Index).Dimensions
           SetRect drawRect, (.Left + c_OffSetX) * c_gifProps.ScaleCx, (.Top + c_OffSetY) * c_gifProps.ScaleCy, _
               .Right * c_gifProps.ScaleCx, .Bottom * c_gifProps.ScaleCy
       End With
       TransferFrame Index, hdc, bStretch, drawRect
       
    Else
        ' append to the running mask, this frame's mask
        UpdateMask Index, False, d3Mask(), d3Color()
        
        ' copy current frame to offscreen buffer
        drawRect = c_Frames(Index).Dimensions
        TransferFrame Index, c_BkBuff.DC, False, drawRect
        
        ' now transfer offscreen to screen
        SetRect drawRect, c_OffSetX, c_OffSetY, UserControl.ScaleWidth, UserControl.ScaleHeight
        TransferFrame -Index, hdc, bStretch, drawRect
        
        ' erase or update the running mask
        If c_Frames(Index).Disposal > 1 Then UpdateMask Index, True, d3Mask(), d3Color()
        
     End If
    RenderFrame = True

End Function

Private Sub TransferFrame(ByVal frameNr As Long, hdc As Long, _
                bStretch As Boolean, destR As RECT)

    If destR.Bottom < 1 Or destR.Right < 1 Then Exit Sub
    
    Dim mROP As Long, dcRect As RECT, hBrush As Long
    Dim oldBrEx As POINTAPI, oldStretchMode As Long
    Dim xOffset As Long, xWidth As Long, xHeight As Long
    Dim SrcDC As Long, lTransColor As Long
        
    With c_Frames(Abs(frameNr))
    
        If bStretch = True Then ' set target DC stretch mode & cache original setting
            oldStretchMode = SetStretchBltMode(hdc, HALFTONE)
            SetBrushOrgEx hdc, 0&, 0&, oldBrEx
        End If
        
        If c_SolidBkgUsed Then
            SetRect dcRect, 0, 0, c_gifProps.Width * c_gifProps.ScaleCx, c_gifProps.Height * c_gifProps.ScaleCy
            If c_SolidBgkFill < 0& Then
                hBrush = CreateSolidBrush(GetSysColor(c_SolidBgkFill And &HFF))
            Else
                hBrush = CreateSolidBrush(c_SolidBgkFill)
            End If
            FillRect hdc, dcRect, hBrush
            DeleteObject hBrush
        End If
        
        
        If frameNr < 0& Then
            ' render to screen from the mask frame & back buffer
            xWidth = c_gifProps.Width
            xHeight = c_gifProps.Height
            
            SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
            If bStretch = True Then
                StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                    c_BkBuff.DC, 0&, 0&, xWidth, xHeight, vbSrcAnd
            Else    ' clipping only
                BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, c_BkBuff.DC, 0&, 0&, vbSrcAnd
            End If
            SelectObject c_BkBuff.DC, c_BkBuff.hDib
            mROP = vbSrcPaint           ' set ROP for the color portion
            SrcDC = c_BkBuff.DC ' the back buffer will now be used to render color to screen
        
        Else    ' otherwise we are transfering frame to passed DC (destination DC)
            
            SrcDC = c_DC.DC
            xOffset = .imgOffset
            xWidth = .Dimensions.Right
            xHeight = .Dimensions.Bottom
            
            If .IsTransparent = 1 Then ' color table hack to create masks for paletted images
            
                ' The c_maskTable is all black, no other colors
                c_maskTable(.TransIndex) = vbWhite   ' set transparent index to white
                SetDIBColorTable c_DC.DC, 0, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_maskTable(0) ' put the table to the DIB
                c_maskTable(.TransIndex) = vbBlack ' set the color back to black
                mROP = vbSrcPaint       ' set ROP for the color portion
                ' now draw the mask to the destination DC, note the ROP used >>
                If bStretch = True Then
                    StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                        SrcDC, xOffset, 0&, xWidth, xHeight, vbSrcAnd
                Else    ' clipping only
                    BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, SrcDC, xOffset, 0&, vbSrcAnd
                End If
                
            Else
                mROP = vbSrcCopy        ' no transparency used; default ROP
            End If
    
            ' do we need to select a different table for the current frame.
            If .IsTransparent = 1 Or Not .TblIndex = c_ColorTables.Index Then
                If .IsTransparent = 1 Then
                    ' cache original transparency color & change it to black
                    lTransColor = c_ColorTables.Tables(.TransIndex, .TblIndex)
                    c_ColorTables.Tables(.TransIndex, .TblIndex) = vbBlack
                End If
                c_ColorTables.Index = .TblIndex  ' update current table ref
                ' change the dib colors & replace transparency color
                SetDIBColorTable c_DC.DC, 0, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_ColorTables.Tables(0, .TblIndex)
                If .IsTransparent = 1 Then c_ColorTables.Tables(.TransIndex, .TblIndex) = lTransColor
            End If
            
        End If
        ' blt the color image and done
        If bStretch = True Then
            StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                SrcDC, xOffset, 0&, xWidth, xHeight, mROP
                ' reset to user DC's original settings
                SetStretchBltMode hdc, oldStretchMode
                SetBrushOrgEx hdc, oldBrEx.X, oldBrEx.Y, oldBrEx
        Else
            BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, SrcDC, xOffset, 0&, mROP
        End If
        
    End With

End Sub

Private Sub UpdateMask(Index As Long, bDisposing As Boolean, d3Mask() As Byte, d3Color() As Byte)

    ' routine maintains a running mask for the GIF

    Dim Rows As Long, gD3row As Long
    Dim mOffset As Long, mScanWidth As Long
    Dim maskBytes() As Byte, colorBytes() As Byte
    Dim maskSA As SAFEARRAY2D, colorSA As SAFEARRAY2D
    Dim eRect As RECT, hBrush As Long

    If Index = 1 Or c_Frames(Index).Disposal = 3 Then
        If Not c_BkBuff.hDibBW = 0& Then
            With maskSA                 ' DIB strip overlay
                .cbElements = 1
                .cDims = 2
                .pvData = c_BkBuff.dibPtrBW
                .rgSABound(0).cElements = c_gifProps.Height
                .rgSABound(1).cElements = ByteAlignOnWord(1, c_gifProps.Width)
            End With
            CopyMemory ByVal VarPtrArray(maskBytes), VarPtr(maskSA), 4&
        End If
        
        With colorSA                ' Buffer overlay
            .cbElements = 1
            .cDims = 2
            .pvData = c_BkBuff.dibPtr
            .rgSABound(0).cElements = c_gifProps.Height
            .rgSABound(1).cElements = ByteAlignOnWord(24, c_gifProps.Width)
        End With
        CopyMemory ByVal VarPtrArray(colorBytes), VarPtr(colorSA), 4&
    End If
    
    If bDisposing Then
        
        With c_Frames(Index)
            Select Case .Disposal
            
            Case 2 ' erase area just drawn
                SetRect eRect, .Dimensions.Left, .Dimensions.Top, .Dimensions.Left + .Dimensions.Right, .Dimensions.Top + .Dimensions.Bottom
                FillRect c_BkBuff.DC, eRect, GetStockObject(BLACK_BRUSH)
                If Not c_BkBuff.hDibBW = 0& Then
                    SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
                    FillRect c_BkBuff.DC, eRect, GetStockObject(WHITE_BRUSH)
                    SelectObject c_BkBuff.DC, c_BkBuff.hDib
                End If
                
            Case 3 ' replace with previous contents
                mOffset = .Dimensions.Left * 3&
                mScanWidth = .Dimensions.Right * 3&
                For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                    CopyMemory colorBytes(mOffset, Rows), d3Color(0, gD3row), mScanWidth
                    gD3row = gD3row + 1&
                Next
                If Not c_BkBuff.hDibBW = 0& Then
                    gD3row = 0&
                    mOffset = .Dimensions.Left \ 8
                    mScanWidth = ByteAlignOnWord(1, .Dimensions.Right)
                    For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                        CopyMemory maskBytes(mOffset, Rows), d3Mask(0, gD3row), mScanWidth
                        gD3row = gD3row + 1&
                    Next
                End If
            End Select
        End With
            
    Else            ' copy new frame onto existing
        
        With c_Frames(Index)
            If Index = 1 Then   ' first frame, erase
                ' fill buffer with black (anti-mask)
                FillMemory colorBytes(0, 0), colorSA.rgSABound(1).cElements * c_gifProps.Height, 0
                ' fill mask with white
                If Not c_BkBuff.hDibBW = 0& Then FillMemory maskBytes(0, 0), maskSA.rgSABound(1).cElements * c_gifProps.Height, 255
            End If
            
            If .Disposal = 3 Then   ' cache current buffer and mask contents as needed
            
                mOffset = .Dimensions.Left * 3&
                mScanWidth = .Dimensions.Right * 3&
                ReDim d3Color(0 To mScanWidth - 1&, 0 To .Dimensions.Bottom - 1&)
                For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                    CopyMemory d3Color(0, gD3row), colorBytes(mOffset, Rows), mScanWidth
                    gD3row = gD3row + 1&
                Next
                If Not c_BkBuff.hDibBW = 0& Then
                    gD3row = 0&
                    mScanWidth = ByteAlignOnWord(1, .Dimensions.Right)
                    mOffset = .Dimensions.Left \ 8
                    ReDim d3Mask(0 To mScanWidth - 1&, 0 To .Dimensions.Bottom - 1&)
                    For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                        CopyMemory d3Mask(0, gD3row), maskBytes(mOffset, Rows), mScanWidth
                        gD3row = gD3row + 1&
                    Next
                End If
                
            End If
            
        End With
        
        If Not c_BkBuff.hDibBW = 0& Then    ' update mask as needed
        
            SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
            With c_Frames(Index)
            
                ' transfer frame's mask to the buffer's mask
            
                If .IsTransparent = 1 Then
                    c_maskTable(.TransIndex) = vbWhite
                    SetDIBColorTable c_DC.DC, 0&, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_maskTable(0) ' put the table to the DIB
                    c_maskTable(.TransIndex) = vbBlack
                    BitBlt c_BkBuff.DC, .Dimensions.Left, .Dimensions.Top, .Dimensions.Right, .Dimensions.Bottom, c_DC.DC, .imgOffset, 0&, vbSrcAnd
                Else
                    SetDIBColorTable c_DC.DC, 0&, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_maskTable(0) ' put the table to the DIB
                    BitBlt c_BkBuff.DC, .Dimensions.Left, .Dimensions.Top, .Dimensions.Right, .Dimensions.Bottom, c_DC.DC, .imgOffset, 0&, vbSrcCopy
                End If
                
            End With
            
            SelectObject c_BkBuff.DC, c_BkBuff.hDib
            c_ColorTables.Index = -1&    ' forces color frames to reselect its palette before rendering it
        
        End If
    
    End If
    
    If Not maskSA.pvData = 0& Then CopyMemory ByVal VarPtrArray(maskBytes), 0&, 4&
    If Not colorSA.pvData = 0& Then CopyMemory ByVal VarPtrArray(colorBytes), 0&, 4&
    
End Sub

Private Sub MirrorGIF(newState As MirrorConstants, SingleFrame As Boolean, MirrorMask As Boolean)

    Dim Index As Long, mirrorState As MirrorConstants
    Dim frameNr As Long, nrFrames As Long
    Dim mirrorX As Long, mirrorY As Long
    Dim mirrorCx As Long, mirrorCy As Long
    
    If (c_Mirror And &H20) = &H20 Then ' &H20 is flag indicating we been thru here before
        frameNr = 2& ' routine called second time
        c_Mirror = c_Mirror And Not &H20
    Else
        frameNr = 1& ' routine called first time
    End If
    
    ' remove any mirror options that are already applied
    mirrorState = c_Mirror Xor newState
    If mirrorState = gfmNone Then Exit Sub ' nothing to do
    
    If SingleFrame Then
        nrFrames = 1&    ' delay start up or IDE design, always one frame
    Else
        nrFrames = UBound(c_Frames)
    End If
    
    For Index = frameNr To nrFrames
        ' flip the color frames
        With c_Frames(Index)
            If (mirrorState And gfmHorizontal) = gfmHorizontal Then
                mirrorX = .imgOffset + .Dimensions.Right - 1&
                mirrorCx = -.Dimensions.Right
                .Dimensions.Left = c_gifProps.Width - (.Dimensions.Right + .Dimensions.Left)
            Else
                mirrorX = .imgOffset
                mirrorCx = .Dimensions.Right
            End If
            If (mirrorState And gfmVertical) = gfmVertical Then
                mirrorY = .Dimensions.Bottom - 1&
                mirrorCy = -.Dimensions.Bottom
                .Dimensions.Top = c_gifProps.Height - (.Dimensions.Bottom + .Dimensions.Top)
            Else
                mirrorY = 0&
                mirrorCy = .Dimensions.Bottom
            End If
            ' flip each frame
            StretchBlt c_DC.DC, .imgOffset, 0&, .Dimensions.Right, .Dimensions.Bottom, _
                c_DC.DC, mirrorX, mirrorY, mirrorCx, mirrorCy, vbSrcCopy
        End With
    Next
    If MirrorMask Then
        If Not c_BkBuff.hDib Then
            If (mirrorState And gfmHorizontal) = gfmHorizontal Then
                mirrorX = c_gifProps.Width - 1&
                mirrorCx = -c_gifProps.Width
            Else
                mirrorX = 0&
                mirrorCx = c_gifProps.Width
            End If
            If (mirrorState And gfmVertical) = gfmVertical Then
                mirrorY = c_gifProps.Height - 1&
                mirrorCy = -c_gifProps.Height
            Else
                mirrorY = 0&
                mirrorCy = c_gifProps.Height
            End If
            ' flip the buffer
            StretchBlt c_BkBuff.DC, 0&, 0&, c_gifProps.Width, c_gifProps.Height, _
                c_BkBuff.DC, mirrorX, mirrorY, mirrorCx, mirrorCy, vbSrcCopy
        
            If Not c_BkBuff.hDibBW = 0& Then
                ' flip the mask
                SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
                StretchBlt c_BkBuff.DC, 0&, 0&, c_gifProps.Width, c_gifProps.Height, _
                    c_BkBuff.DC, mirrorX, mirrorY, mirrorCx, mirrorCy, vbSrcCopy
                SelectObject c_BkBuff.DC, c_BkBuff.hDib
            End If
            
        End If
    End If
    
End Sub

Private Sub ScaleToDestination(ByRef X As Long, ByRef Y As Long, _
                ByRef cx As Long, ByRef cy As Long, _
                ByRef xRatio As Single, ByRef yRatio As Single)

    ' function scales an image to the target destination based on the
    ' stretch mode setting, image dimensions, and target dimensions
    
    If c_curFrame = 0& Then Exit Sub ' no gif loaded yet
    
    If c_ScaleMode = gfsClip Then
        xRatio = 1
        yRatio = 1
        If c_gifProps.Width < cx Then cx = c_gifProps.Width
        If c_gifProps.Height < cy Then cy = c_gifProps.Height
    
    ElseIf c_ScaleMode = gfsActualSize Then
        cx = c_gifProps.Width
        cy = c_gifProps.Height
        xRatio = 1
        yRatio = 1
    
    Else ' scaling in one way or another....
        xRatio = cx / c_gifProps.Width
        yRatio = cy / c_gifProps.Height
        Select Case c_ScaleMode
        Case gfsStretch
            ' nothing to calculate when stretching
        Case gfsShrinkScaleToFit
            If xRatio > 1 And yRatio > 1 Then
                xRatio = 1  ' image will fit without scaling; use 1:1 scaling
                yRatio = 1
            Else            ' image must be scaled; use same ratio for width/height
                If xRatio > yRatio Then xRatio = yRatio Else yRatio = xRatio
            End If
        Case Else ' always scale to target dimensions; use same ratio for width/height
            If xRatio > yRatio Then xRatio = yRatio Else yRatio = xRatio
        End Select
    
        If Not (c_ScaleMode = gfsStretch) Then
            cx = c_gifProps.Width * xRatio
            cy = c_gifProps.Height * yRatio
        End If
    End If
End Sub

Private Function PictureFromByteStream(inArray() As Byte, Size As Long) As IPicture
  

    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(LBound(inArray)), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), PictureFromByteStream)
            End If
        End If
    End If
End Function

Private Function ConvertStripToGIF(ByVal Index As Long) As StdPicture



  Const MAX_BITS                    As Long = 12    ' Per GIF docs, 12 is the Max
  Const MAX_BITSHIFT                As Long = 4096  ' 2^MAXBITS
  Const MAX_CODE                    As Long = 4096  ' Should NEVER generate this code
  Const TABLE_SIZE                  As Long = 5003  ' 80% occupancy (hash)
  
  Dim lIdx     As Long      ' hash table index
  Dim lFCode   As Long      ' pixel pattern (hashed)
  Dim LC       As Long      ' most recent pixel
  Dim lEnt     As Long      ' previous known pattern
  Dim lDisp    As Long
  Dim m_lShift As Long
  
  Dim m_lCodeCount As Long      ' count of codes used
  Dim m_lMaxCode As Long        ' Maximum code, given m_lBits
  Dim m_BitBucketCount As Long  ' bit bucket bit counter
  Dim m_BitBucketBuff As Long   ' cache of bits/bytes to process
  Dim m_lCurrentBits As Long    ' current lzw compression size (variable)
  Dim m_lSubBlockSize As Long   ' flag to track last byte in a data subBlock

  '-- Block compression parameters.
  Dim m_lInitBits      As Long  ' baseline LZW compression size
  Dim m_lClearCode     As Long  ' clear code
  Dim m_lHashTable() As Long
  Dim m_lCodeTable() As Long

  On Error GoTo EH
  ReDim m_lHashTable(0 To TABLE_SIZE - 1&)
  ReDim m_lCodeTable(0 To TABLE_SIZE - 1&)

  ' Added by LaVolpe to read from the custom DIB strip
  Dim tSA As SAFEARRAY2D, dibBytes() As Byte        ' DMA DIB processing
  Dim bExistingPattern As Boolean                   ' indicates hash match found
  Dim dibRow As Long, dibCol As Long, aPtr As Long  ' DIB looping parameters
  Dim tBMPI As BITMAPINFO
  
    'Added by LaVolpe -- gif frame custom build; byte at a time
    ' the actual width & height of the DIB strip is not cached; get it
    tBMPI.bmiHeader.biSize = 40
    GetDIBits c_DC.DC, c_DC.hDib, 0&, 0&, ByVal 0&, tBMPI, 0&
    With tSA
        .rgSABound(0).cElements = tBMPI.bmiHeader.biHeight
        .rgSABound(1).cElements = ByteAlignOnWord(8, tBMPI.bmiHeader.biWidth)
        .cbElements = 1
        .cDims = 2
        .pvData = c_DC.dibPtr
    End With
    CopyMemory ByVal VarPtrArray(dibBytes()), VarPtr(tSA), 4&
    
    ' oversize/guesstimate compressed GIF data & include GIF block information
    ReDim c_DIBarray(0 To 800& + (c_Frames(Index).Dimensions.Bottom * c_Frames(Index).Dimensions.Right))
    '^ 800= 13 hdr + 10 img descrip + 8 img ctrl + 768 palette + 1 EOF flag
    With c_Frames(Index)
        ' start building the GIF frame by hand, a byte at a time
        CopyMemory c_DIBarray(0), &H38464947, 4&  ' add the 6-byte GIF89a signature
        CopyMemory c_DIBarray(4), &H6139&, 2&
        CopyMemory c_DIBarray(6), .Dimensions.Right, 2&   ' now the width
        CopyMemory c_DIBarray(8), .Dimensions.Bottom, 2&  ' & height
        ' add the color table flag and the table itself
        c_DIBarray(10) = 135 ' global color table @ 8 bits (128 or 7)
        ' can skip next 2 bytes: bkg window palette index & pixel aspect ratio
        aPtr = 13&
        For lIdx = 0& To c_ColorTables.Tables(PALETTECOUNT, .TblIndex) - 1& ' convert BGR palette to RGB
            c_DIBarray(aPtr) = (c_ColorTables.Tables(lIdx, .TblIndex) \ &H10000) And &HFF
            c_DIBarray(aPtr + 1&) = (c_ColorTables.Tables(lIdx, .TblIndex) \ &H100&) And &HFF
            c_DIBarray(aPtr + 2&) = c_ColorTables.Tables(lIdx, .TblIndex) And &HFF
            aPtr = aPtr + 3
        Next
        aPtr = 781& ' 256 * 3 + 13 : note that aPtr can come out of loop less than 781
        ' add the image control block if needed
        If .IsTransparent = 1 Then
            CopyMemory c_DIBarray(aPtr), &H4F921, 3&    ' Introducer(33);BlockID(249);BlockLen(4)
            c_DIBarray(aPtr + 3&) = 1                   ' transparency flag
            ' skip next 2 bytes which is the Delay time; not needed for single frame
            c_DIBarray(aPtr + 6&) = .TransIndex
            ' skip next byte which is the block terminator - always zero
            aPtr = aPtr + 8& ' next position in array
        End If
        ' add the image description block
        c_DIBarray(aPtr) = 44 ' BlockID
        ' Left and Top coords are zeros; write frame width & height
        CopyMemory c_DIBarray(aPtr + 5&), c_DIBarray(6), 4&
        c_DIBarray(aPtr + 10&) = 8 ' LZW compression size; image bit depth
        aPtr = aPtr + 11& '11=img description block size of 10 + LZW compression byte
        ' next comes compressing DIB into LZW sub blocks, then finishing off stream
    End With
    
    'Initialize Masks -- Init LUT for fast 2 ^ x - 1 (was InitMasks routine)
    ReDim c_aPOT(0 To 16)
    c_aPOT(0) = 0
    For lIdx = 1& To 16&
        c_aPOT(lIdx) = 2 * (c_aPOT(lIdx - 1&) + 1&) - 1&
    Next lIdx
    
    '-- Reset output buffer values
    ReDim c_aBuff(0 To 254) ' (was pvCharInit)

    '-- Set up the necessary startup values
    m_lInitBits = 9& '(DIB bit depth + 1)
    m_lCurrentBits = m_lInitBits
    m_lMaxCode = c_aPOT(m_lCurrentBits)
    m_lClearCode = PALETTECOUNT     ' 2^(m_lInitBits - 1)
    m_lCodeCount = PALETTECOUNT + 2& ' 2^(m_lInitBits - 1) + 2
    
    '-- Set hash code range bound for shifting
    lFCode = TABLE_SIZE
    Do While lFCode < 65536
        m_lShift = m_lShift + 1&
        lFCode = lFCode + lFCode
    Loop
    m_lShift = 1& + c_aPOT(8& - m_lShift)
    'Added by LaVolpe -- quick erase, setting all table entries to -1 (was pvClearTable)
    FillMemory m_lHashTable(0), TABLE_SIZE * 4&, 255  ' clear hash table
    
    '-- Start...
    m_lSubBlockSize = 1& ' position for 1st byte in data sub block
    ' all images begin with a clear table flag
    Call pvOutputCode(m_lClearCode, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    
    ' start LZW patterns & also start looping on 2nd pixel
    lEnt = dibBytes(c_Frames(Index).imgOffset, c_Frames(Index).Dimensions.Bottom - 1&)
    dibCol = c_Frames(Index).imgOffset + 1&
    
    'Added by LaVolpe -- looping is my enhancement;
    ' hash algorithm reorganized, comments added, modified very little
    For dibRow = tBMPI.bmiHeader.biHeight - 1& To tBMPI.bmiHeader.biHeight - c_Frames(Index).Dimensions.Bottom Step -1&
    
        ' process each pixel in line of image
        For dibCol = dibCol To c_Frames(Index).imgOffset + c_Frames(Index).Dimensions.Right - 1&
        
            LC = dibBytes(dibCol, dibRow)
            
            lFCode = LC * MAX_BITSHIFT + lEnt   ' add to existing pattern
            lIdx = (LC * m_lShift) Xor lEnt     ' XOR hashing
    
            If lIdx >= TABLE_SIZE Then lIdx = 0& ' added by LaVolpe (sanity check)
            
            If (m_lHashTable(lIdx) = lFCode) Then   ' found existing pattern
                lEnt = m_lCodeTable(lIdx)
            Else
                If (m_lHashTable(lIdx) > -1&) Then ' else Empty slot
                    
                    lDisp = TABLE_SIZE - lIdx     ' Secondary hash (after G. Knott)
                    If (lIdx = 0&) Then lDisp = 1&
                    
                    Do  ' Hash Probing
                        lIdx = lIdx - lDisp
                        If (lIdx < 0&) Then lIdx = lIdx + TABLE_SIZE
        
                        If (m_lHashTable(lIdx) = lFCode) Then
                            lEnt = m_lCodeTable(lIdx)
                            bExistingPattern = True
                            Exit Do
                        End If
        
                    Loop While (m_lHashTable(lIdx) > 0&) ' Continue probing
                End If
                
                If bExistingPattern = True Then
                    bExistingPattern = False    ' reset flag
                Else
                    ' write previous pattern
                    Call pvOutputCode(lEnt, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
                    lEnt = LC  ' set current palette index as previous pattern
                    m_lCodeTable(lIdx) = m_lCodeCount ' store code & hash index
                    m_lHashTable(lIdx) = lFCode
                    
                    ' check for LZW compression increments
                    If m_lCodeCount > m_lMaxCode Then
                        ' ran out of codes for current compression size
                        If (m_lCodeCount = MAX_CODE) Then
                            ' add clear code to output stream (was pvClearBlock)
                            Call pvOutputCode(m_lClearCode, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
                            FillMemory m_lHashTable(0), TABLE_SIZE * 4&, 255  ' clear hash table
                            m_lCurrentBits = m_lInitBits    ' reset baseline LZW compression size
                            m_lCodeCount = PALETTECOUNT + 1& ' reset code counter to 1 less 'cause it is incremented right away
                        Else
                            m_lCurrentBits = m_lCurrentBits + 1& ' increment compression size
                        End If
                        m_lMaxCode = c_aPOT(m_lCurrentBits) ' new max count for current compression size
                    End If
                    m_lCodeCount = m_lCodeCount + 1&  ' increment the number of patterns
                End If
            End If
        Next dibCol
        dibCol = c_Frames(Index).imgOffset  ' reset to 1st pixel of line
    Next dibRow

    '--  Put out the final code & image data termination code
    Call pvOutputCode(lEnt, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    ' finish off the stream & add end of image data flag
    Call pvOutputCode(PALETTECOUNT + 1&, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    ' flush remaining bytes in output bitbucket
    Call pvOutputCode(-1, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    ' the next byte would be a sub block terminator which is zero; ignore & go on
    c_DIBarray(aPtr + 1) = 59   ' add end of file flag
        
    Set ConvertStripToGIF = PictureFromByteStream(c_DIBarray(), aPtr + 2&)

EH:
    If ERR Then
        Stop        ' for testing if error occurs
        ERR.Clear
        Resume      ' for testing if error occurs
    End If
    
    ' clear global arrays, no longer used
    If tSA.cbElements = 1 Then CopyMemory ByVal VarPtrArray(dibBytes()), 0&, 4&
    Erase c_DIBarray()
    Erase c_aPOT()
    Erase c_aBuff()
End Function

Private Function isStringANSI(inText As String) As Boolean

    ' simple test to determine if passed string is ANSI-like or not.
    ' In other words, does it contain unicode characters.
    
    Dim tArray() As Byte
    Dim X As Long
    
    If inText = vbNullString Then
    
        isStringANSI = True
    
    Else
    
        tArray = inText
        For X = LBound(tArray) + 1 To UBound(tArray) Step 2
            If Not tArray(X) = 0 Then Exit For
        Next
        
        isStringANSI = (X > UBound(tArray))
        
    End If

End Function


Private Sub pvOutputCode(ByVal lCode As Long, ByRef arrayPtr As Long, _
                        lBitBucketCount As Long, lBitBucketBuff As Long, _
                        lCurrentBits As Long, lSubBlockSize As Long)



    If lCode < 0 Then
    
        If Not lBitBucketCount = 0 Then ' add last bits of the output buffer to array
            c_aBuff(lSubBlockSize) = lBitBucketBuff
            lSubBlockSize = lSubBlockSize + 1
        End If
        If Not lSubBlockSize = 1 Then ' still have bytes to write
            c_aBuff(0) = lSubBlockSize
            ' ensure array large enough for this block + 5 extra bytes needed to finish off stream
            If UBound(c_DIBarray) < arrayPtr + lSubBlockSize + 6 Then
                ReDim Preserve c_DIBarray(0 To arrayPtr + lSubBlockSize + 6)
            End If
            CopyMemory c_DIBarray(arrayPtr), c_aBuff(0), lSubBlockSize + 1
            arrayPtr = arrayPtr + lSubBlockSize + 1
        End If
    
    Else
        ' add latest code to the bitbucket & track total bit count
        lBitBucketBuff = lBitBucketBuff Or (lCode * (c_aPOT(lBitBucketCount) + 1))
        lBitBucketCount = lBitBucketCount + lCurrentBits
    
        Do Until lBitBucketCount < 8
            ' remove 8 bits at a time and place in output buffer (byte array)
            c_aBuff(lSubBlockSize) = (lBitBucketBuff And &HFF&)
            lBitBucketBuff = lBitBucketBuff \ &H100&
            lBitBucketCount = lBitBucketCount - 8
            If (lSubBlockSize = 254) Then
                ' max allowable subblock data size is 255 byte blocks; write it
                c_aBuff(0) = lSubBlockSize
                If UBound(c_DIBarray) < arrayPtr + 260 Then
                    ' sanity check; unless the image is not compressible we
                    ' shouldn't trigger this Redim. ^^ 260=255+5 end of file/image bytes
                    ReDim Preserve c_DIBarray(0 To arrayPtr + 512)
                End If
                CopyMemory c_DIBarray(arrayPtr), c_aBuff(0), lSubBlockSize + 1
                arrayPtr = arrayPtr + lSubBlockSize + 1
                lSubBlockSize = 0
            End If
            ' keep track of next byte position to write to
            lSubBlockSize = lSubBlockSize + 1
        Loop
    End If

End Sub

Private Function BuildBackBuffer(IncludeMask As Boolean) As Boolean

    Dim dDC As Long, bErrors As Boolean, tBMPI As BITMAPINFO
    ' Simply creates a 24bpp DIB of the overall GIF dimensions
    With tBMPI.bmiHeader
        .biSize = 40
        .biBitCount = 24
        .biHeight = c_gifProps.Height
        .biWidth = c_gifProps.Width
        .biPlanes = 1
    End With
    
    ' if any errors happen, cleanup occurs in the calling routine: LoadGIF
    dDC = GetDC(0&)
    c_BkBuff.DC = CreateCompatibleDC(dDC)
    If c_BkBuff.DC = 0& Then
        bErrors = True
    Else
        c_BkBuff.hDib = CreateDIBSection(dDC, tBMPI, 0, c_BkBuff.dibPtr, 0, 0)
        If c_BkBuff.hDib = 0& Then
            bErrors = True
        Else
            c_BkBuff.hBmp = SelectObject(c_BkBuff.DC, c_BkBuff.hDib)
            If IncludeMask Then
                ' the buffer requires a mask, create that too
                With tBMPI.bmiHeader
                    .biBitCount = 1         ' monochrome mask
                    .biClrUsed = 2
                    .biClrImportant = 2
                End With
                tBMPI.bmiPalette(1) = vbWhite
                c_BkBuff.hDibBW = CreateDIBSection(dDC, tBMPI, 0, c_BkBuff.dibPtrBW, 0, 0)
                bErrors = (c_BkBuff.hDibBW = 0&)
            End If
        End If
    End If
    ReleaseDC 0&, dDC
    BuildBackBuffer = Not bErrors
End Function

Private Function ManageTimers(ByVal Mode As Long) As Boolean

    If Mode = 0 Then
        If UserControl.Ambient.UserMode = True Then KillTimer c_tmrOwner, ObjPtr(Me)
        c_aniState = gfaStop
    Else
        Dim newDelay As Long
        If Not (c_aniState = gfaPlay) Then
            c_aniState = gfaStop
            If Not c_curFrame = 0& Then
                If UserControl.Ambient.UserMode = True Then
                    ' create timer during runtime only unless overridden
                    If UBound(c_Frames) > 1 Or Mode = 2 Then ' and only if we have more than 1 frame
                        On Error Resume Next
                        If Mode = 1 Then                ' runtime timer
                            c_Ptr = zb_AddressOf(1, 4)
                        ElseIf Mode = 2 Then            ' delay startup timer
                            c_Ptr = zb_AddressOf(2, 4)
                        End If
                        If Not c_Ptr = 0 Then  ' so far so good
                            c_tmrOwner = UserControl.Parent.hwnd    ' cache; .Parent is invalid when closing
                            If c_tmrOwner = 0 Then
                                c_Ptr = 0
                            Else
                                ' test the timer function
                                If SetTimer(c_tmrOwner, ObjPtr(Me), 60000, c_Ptr) = 0 Then
                                    c_Ptr = 0 ' timer failure; can't animate
                                Else
                                    KillTimer c_tmrOwner, ObjPtr(Me) ' success, kill the test timer
                                    If Mode = 2 Then    ' delay startup timer
                                        SetTimer c_tmrOwner, ObjPtr(Me), 20, c_Ptr
                                    Else                ' runtime timer
                                        If c_StepDelay Then
                                            newDelay = c_StepDelay
                                        Else
                                            newDelay = c_Frames(c_curFrame).Delay
                                            If newDelay < c_MinDelay Then newDelay = c_MinDelay
                                        End If
                                        SetTimer c_tmrOwner, ObjPtr(Me), newDelay, c_Ptr
                                    End If
                                    c_aniState = gfaPlay
                                    ManageTimers = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function zb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
Const MAX_FUNKS   As Long = 1                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    Call SHOWWRONG("nThunkNo doesn't exist.", 0)
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    Call SHOWWRONG("Callback address not found.", 0)
    Exit Function
  End If
  
  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  End If
  
  If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
    z_Cb(3, nThunkNo) = _
              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, nThunkNo) = &HBB60E089
    z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
  End If
  
  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
  If bIdeSafety Then                                                        'If the user wants IDE protection
    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
  End If
    
  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
  zb_AddressOf = nAddr + 16                                                 'Thunk code start address
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address.
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Sub zTerminate()
    
    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
    If Not z_CbMem = 0 Then                                            'If memory allocated
        If Not VirtualFree(z_CbMem, 0, MEM_RELEASE) = 0 Then
            z_CbMem = 0  'Release; Indicate memory released
            Erase z_Cb()
        End If
    End If
End Sub

Private Function TimerProcBuildUp(ByVal hwnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long

    ManageTimers 0
    BuildDIBstrip False ' continue processing the remaining frames

End Function
Private Function TimerProc(ByVal hwnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    
    KillTimer c_tmrOwner, TimerID    ' stop current timer
    
    Dim bRestart As Boolean
    Dim bLoopComplete As Boolean
    Dim tValue As Long
    
    ' determine next frame in the animation order
    tValue = c_curFrame
    c_curFrame = c_curFrame + 1
    If c_curFrame > UBound(c_Frames) Then
        c_curFrame = 1
        bLoopComplete = True
    End If
    If UBound(c_Frames) = 1 Then  ' single frame gif
        c_AniLoops = 0            ' shouldn't get here/timer should not have been created
        c_aniState = gfaStop
    Else
        ' determine if timer should continue
        bRestart = True
        If bLoopComplete = True Then ' another loop finished
            If Not c_gifProps.Loops = 0 Then  ' has specified number of animation loops
                c_AniLoops = c_AniLoops - 1 ' decrease number of loops remaining
                If c_AniLoops = 0 Then ' end of desired loops, no more animation
                    c_aniState = gfaStop
                    bRestart = False
                End If
            End If
        End If
    End If
    
    UserControl.Refresh
    RaiseEvent FrameChanged(tValue, True)

    If bRestart = True Then ' set timer for current frame
        If c_StepDelay Then
            tValue = c_StepDelay
        Else
            If c_Frames(c_curFrame).Delay < c_MinDelay Then
                tValue = c_MinDelay
            Else
                tValue = c_Frames(c_curFrame).Delay
            End If
        End If
        SetTimer hwnd, TimerID, tValue, c_Ptr ' set timer
    Else
        RaiseEvent LoopsEnded
    End If
EH:
End Function
