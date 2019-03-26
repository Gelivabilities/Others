VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
Option Explicit
'ÒÆ¶¯ÎÞÍÏÓ°
Private Enum eMsgWhen
    [MSG_AFTER] = 1                                  'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
End Enum
    
Private Const ALL_MESSAGES    As Long = -1          'All messages added or deleted
Private Const CODE_LEN        As Long = 197        'Length of the machine code in bytes
Private Const GWL_WNDPROC      As Long = -4          'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04        As Long = 88          'Table B (before) address patch offset
Private Const PATCH_05        As Long = 93          'Table B (before) entry count patch offset
Private Const PATCH_08        As Long = 132        'Table A (after) address patch offset
Private Const PATCH_09        As Long = 137        'Table A (after) entry count patch offset
    
Private Type tSubData                                'Subclass data type
    hwnd                      As Long              'Handle of the window being subclassed
    nAddrSub                  As Long              'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long              'The address of the pre-existing WndProc
    nMsgCntA                  As Long              'Msg after table entry count
    nMsgCntB                  As Long              'Msg before table entry count
    aMsgTblA()                As Long              'Msg after table array
    aMsgTblB()                As Long              'Msg Before table array
End Type
    
Private sc_aSubData()          As tSubData          'Subclass data array
Private sc_aBuf(1 To CODE_LEN) As Byte              'Code buffer byte array
Private sc_pCWP                As Long              'Address of the CallWindowsProc
Private sc_pEbMode            As Long              'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long              'Address of the SetWindowsLong function
    
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
    
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_SETFOCUS = &O7
    
Public Event MsgCome(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

Private Sub Class_Terminate()
    Call Subclass_StopAll
End Sub

'========================================================================================
' Subclass handler: MUST be the first Public routine in this file.
'                  That includes public properties also.
'========================================================================================
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    RaiseEvent MsgCome(bBefore, bHandled, lReturn, lng_hWnd, uMsg, wParam, lParam)
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function AddWindowMsgs(ByVal hwnd As Long) As Boolean
    '-- Start subclassing
    Call Subclass_Start(hwnd)
    Call Subclass_AddMsg(hwnd, ALL_MESSAGES, MSG_BEFORE)
    '-- Success
    AddWindowMsgs = True
End Function

Public Function DeleteWindowMsg(ByVal hwnd As Long) As Boolean
End Function

'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_... routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
    'Parameters:
    '  lng_hWnd - The handle of the window for which the uMsg is to be added to the callback table
    '  uMsg    - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
    '  When    - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Delete a message from the table of those that will invoke a callback.
    'Parameters:
    '  lng_hWnd - The handle of the window for which the uMsg is to be removed from the callback table
    '  uMsg    - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    '  When    - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
    'Return whether we're running in the IDE.
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    'Start subclassing the passed window handle
    'Parameters:
    '  lng_hWnd - The handle of the window to be subclassed
    'Returns;
    '  The sc_aSubData() index
    Dim i                        As Long                      'Loop index
    Dim j                        As Long                      'Loop index
    Dim nSubIdx                  As Long                      'Subclass data index
    Dim sSubCode                As String                    'Subclass code string
    
    Const GMEM_FIXED            As Long = 0                  'Fixed memory GlobalAlloc flag
    Const PAGE_EXECUTE_READWRITE As Long = &H40&              'Allow memory to execute without violating XP SP2 Data Execution Prevention
    Const PATCH_01              As Long = 18                  'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                  'Address of the previous WndProc
    Const PATCH_03              As Long = 78                  'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                'Address of the previous WndProc
    Const PATCH_07              As Long = 121                'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                'Address of the owner object
    Const FUNC_CWP              As String = "CallWindowProcA" 'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"          'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"          'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"            'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"            'Location of the EbMode function if running VB6
    
    'If it's the first time through here..
    If (sc_aBuf(1) = 0) Then
        'Build the hex pair subclass string
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3"
        'Convert the string from hex pairs to bytes and store in the machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                      'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                      'Next pair of hex characters
        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(16) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_aBuf(17) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found.
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If
        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                  'Patch the address of this object instance into the static machine code buffer
        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                  'Get the address of the CallWindowsProc function
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                  'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                    'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then                                                    'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                  'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                  'Create a new sc_aSubData element
        End If
        Subclass_Start = nSubIdx
    End If
    With sc_aSubData(nSubIdx)
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                            'Allocate memory for the machine code WndProc
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i) 'Mark memory as executable
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                'Copy the machine code from the static byte array to the code array in sc_aSubData
        
        .hwnd = lng_hWnd                                                          'Store the hWnd
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                'Set our WndProc in place
        
        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                          'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                          'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                              'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                          'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                              'Patch the relative address of the CallWindowProc api function
    End With
End Function

Private Sub Subclass_StopAll()
    'Stop all subclassing
    Dim i As Long
    
    i = UBound(sc_aSubData())                                                    'Get the upper bound of the subclass data array
    Do While i >= 0                                                              'Iterate through each element
        With sc_aSubData(i)
            If (.hwnd <> 0) Then                                                  'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hwnd)                                        'Subclass_Stop
            End If
        End With
        
        i = i - 1                                                                'Next element
    Loop
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    'Stop subclassing the passed window handle
    'Parameters:
    '  lng_hWnd - The handle of the window to stop being subclassed
    
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                      'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                    'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                    'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                'Release the machine code memory
        .hwnd = 0                                                                'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                            'Clear the before table
        .nMsgCntA = 0                                                            'Clear the after table
        Erase .aMsgTblB                                                          'Erase the before table
        Erase .aMsgTblA                                                          'Erase the after table
    End With
End Sub

'----------------------------------------------------------------------------------------
'These z... routines are exclusively called by the Subclass_... routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    'Worker sub for Subclass_AddMsg
    
    Dim nEntry  As Long                                                            'Message table entry index
    Dim nOff1  As Long                                                            'Machine code buffer offset 1
    Dim nOff2  As Long                                                            'Machine code buffer offset 2
    
    If (uMsg = ALL_MESSAGES) Then                                                'If all messages
        nMsgCnt = ALL_MESSAGES                                                    'Indicates that all messages will callback
    Else                                                                        'Else a specific message number
        Do While nEntry < nMsgCnt                                                'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            
            If (aMsgTbl(nEntry) = 0) Then                                        'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                            'Re-use this entry
                Exit Sub                                                          'Bail
            ElseIf (aMsgTbl(nEntry) = uMsg) Then                                  'The msg is already in the table!
                Exit Sub                                                          'Bail
            End If
        Loop                                                                      'Next entry
        
        nMsgCnt = nMsgCnt + 1                                                    'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                              'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                  'Store the message number in the table
    End If
    
    If (When = eMsgWhen.MSG_BEFORE) Then                                          'If before
        nOff1 = PATCH_04                                                          'Offset to the Before table
        nOff2 = PATCH_05                                                          'Offset to the Before table entry count
    Else                                                                        'Else after
        nOff1 = PATCH_08                                                          'Offset to the After table
        nOff2 = PATCH_09                                                          'Offset to the After table entry count
    End If
    
    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                          'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                        'Patch the appropriate table entry count
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    'Return the memory address of the passed function in the passed dll
    
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                        'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    'Worker sub for Subclass_DelMsg
    Dim nEntry As Long
    
    If (uMsg = ALL_MESSAGES) Then                                                'If deleting all messages
        nMsgCnt = 0                                                              'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                        'If before
            nEntry = PATCH_05                                                    'Patch the before table message count location
        Else                                                                    'Else after
            nEntry = PATCH_09                                                    'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                          'Patch the table message count to zero
    Else                                                                        'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                'For each table entry
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = uMsg) Then                                      'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                              'Mark the table slot as available
                Exit Do                                                          'Bail
            End If
        Loop                                                                      'Next entry
    End If
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    'Get the sc_aSubData() array index of the passed hWnd
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                            'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If (.hwnd = lng_hWnd) Then                                            'If the hWnd of this element is the one we're looking for
                If (Not bAdd) Then                                                'If we're searching not adding
                    Exit Function                                                'Found
                End If
            ElseIf (.hwnd = 0) Then                                              'If this an element marked for reuse.
                If (bAdd) Then                                                    'If we're adding
                    Exit Function                                                'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                          'Decrement the index
    Loop
    
    If (Not bAdd) Then
        Debug.Assert False                                                        'hWnd not found, programmer error
    End If
    
    'If we exit here, we're returning -1, no freed elements were found
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    'Patch the machine code buffer at the indicated offset with the relative address to the target address.
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    'Patch the machine code buffer at the indicated offset with the passed value
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    'Worker function for Subclass_InIDE
    zSetTrue = True
    bValue = True
End Function


