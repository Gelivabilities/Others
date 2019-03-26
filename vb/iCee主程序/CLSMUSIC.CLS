VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private m_objBasicAudio As IBasicAudio
Private m_objBasicVideo As IBasicVideo
Private m_objVideoWindow As IVideoWindow
Private m_objMediaControl As IMediaControl
Private m_objMediaPosition As IMediaPosition
Private m_hWnd As Long
Private m_FileName As String
Private m_HasAudio As Boolean
Private m_HasVideo As Boolean

' Constants
Private Const WS_VISIBLE = &H10000000

' Types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Enums
Public Enum enumState
    stStopped
    stPlaying
    stPaused
End Enum

' Functions


Public Property Get Width()
    ' Return the video width
    On Error Resume Next
    If HasVideo Then
        Width = m_objBasicVideo.SourceWidth
    Else
        Width = 0
    End If
End Property


Public Property Get Height()
    ' Return the video height
    On Error Resume Next
    If HasVideo Then
        Height = m_objBasicVideo.SourceHeight
    Else
        Height = 0
    End If
End Property


Public Property Let Window(ByVal hWnd As Long)
    ' Save the window
    On Error Resume Next
    m_hWnd = hWnd
End Property


Public Property Get Window() As Long
    ' Return the window
    On Error Resume Next
    Window = m_hWnd
End Property


Public Property Get FileName() As String
    ' Return the filename
    On Error Resume Next
    FileName = m_FileName
End Property


Public Property Let FileName(ByVal value As String)
    ' Set the filename
    Dim r As RECT
    Dim temp As Double
    On Error Resume Next
    
    ' Save the filename
    m_FileName = value
    
    ' Unload old objects
    If Not (m_objMediaControl Is Nothing) Then
        m_objMediaControl.Stop
        Set m_objBasicAudio = Nothing
        Set m_objBasicVideo = Nothing
        Set m_objMediaControl = Nothing
        Set m_objVideoWindow = Nothing
        Set m_objMediaPosition = Nothing
    End If
    
    ' Load new objects
    Set m_objMediaControl = New FilgraphManager
    m_objMediaControl.RenderFile FileName
    Set m_objBasicAudio = m_objMediaControl
    Set m_objBasicVideo = m_objMediaControl
    Set m_objVideoWindow = m_objMediaControl
    Set m_objMediaPosition = m_objMediaControl
    
    ' Check to see if there's audio
    temp = m_objBasicAudio.Volume
    If ERR.Number <> 0 Then
        HasAudio = False
    Else
        HasAudio = True
    End If
    ERR.Clear
    
    
    ' Check to see if there's video
    temp = m_objBasicVideo.SourceWidth
    If ERR.Number <> 0 Then
        HasVideo = False
    Else
        HasVideo = True
    End If
    ERR.Clear
    
    ' Hook the video to your window
    If HasVideo Then
        ' Set the video size
        GetWindowRect m_hWnd, r
        With m_objVideoWindow
            .WindowStyle = WS_VISIBLE
            .Top = 0
            .Left = 0
            .Width = r.Right - r.Left
            .Height = r.Bottom - r.Top
            .Owner = m_hWnd
            .Visible = 1
        End With
        ' Show the first frame
        Position = 0
    End If

End Property


Public Sub Play()
    ' Play the file
    On Error Resume Next
    If Len(m_FileName) = 0 Then
        Exit Sub
    End If
    If Position = Duration Then
        Position = 0
    End If
    m_objMediaControl.Run
End Sub


Public Sub Pause()
    'Pause playing
    On Error Resume Next
    m_objMediaControl.Pause
End Sub


Public Sub StopPlaying()
    'Stop playing
    On Error Resume Next
    m_objMediaControl.Stop
End Sub


Public Property Let Position(ByVal NewPos As Double)
    'Set the position
    On Error Resume Next
    With m_objMediaPosition
        If NewPos < 0 Then
            NewPos = 0
        ElseIf NewPos > .Duration Then
            NewPos = .Duration
        End If
        m_objMediaPosition.currentPosition = NewPos
        If HasVideo Then
            m_objMediaControl.Run
            Do Until m_objMediaPosition.currentPosition > NewPos
            Loop
            m_objMediaPosition.currentPosition = NewPos
            m_objMediaControl.Stop
        End If
    End With
End Property


Public Property Get Position() As Double
    ' Get the position
    On Error Resume Next
    Position = m_objMediaPosition.currentPosition
End Property


Public Property Get State() As enumState
    ' Return the state
    On Error Resume Next
    Dim TimeOut As Long
    Dim s As Long
    m_objMediaControl.GetState TimeOut, s
    If s = 0 Then
        State = stStopped
    ElseIf s = 1 Then
        State = stPaused
    Else
        State = stPlaying
    End If
End Property


Public Property Get Speed() As Double
    ' Return the speed (normal speed is 1)
    On Error Resume Next
    Speed = m_objMediaPosition.rate
End Property


Public Property Let Speed(ByVal value As Double)
    ' Set the speed
    On Error Resume Next
    If value < 0.1 Then
        value = 0.1
    End If
    m_objMediaPosition.rate = value
End Property


Public Property Get Duration() As Double
    ' Return how long the song is in milliseconds
    On Error Resume Next
    Duration = m_objMediaPosition.Duration
End Function


Public Property Get Volume() As Long
    ' Return the volume
    On Error Resume Next
    Volume = m_objBasicAudio.Volume
End Property


Public Property Let Volume(ByVal value As Long)
    ' Set the volume
    On Error Resume Next
    If value > 0 Then value = 0
    If value < -10000 Then value = -10000
    m_objBasicAudio.Volume = value
End Property


Public Property Get Balance() As Long
    ' Return the balance
    On Error Resume Next
    Balance = m_objBasicAudio.Balance
End Property


Public Property Let Balance(ByVal value As Long)
    ' Set the balance
    On Error Resume Next
    If value < -10000 Then value = -10000
    If value > 10000 Then value = 10000
    m_objBasicAudio.Balance = value
End Property


Public Property Get HasAudio() As Boolean
    On Error Resume Next
    HasAudio = m_HasAudio
End Property


Private Property Let HasAudio(ByVal value As Boolean)
    On Error Resume Next
    m_HasAudio = value
End Property


Public Property Get HasVideo() As Boolean
    On Error Resume Next
    HasVideo = m_HasVideo
End Property


Private Property Let HasVideo(ByVal value As Boolean)
    On Error Resume Next
    m_HasVideo = value
End Property

