VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ºÙ«–∞ÂªÒ»°
Option Explicit
Private IMG As IPictureDisp
Private mvarIsImage As Boolean 'local copy
Private mvarClipText As String 'local copy
Public Property Let ClipText(ByVal vData As String)
On Error Resume Next
    mvarClipText = vData
End Property
Public Property Get ClipText() As String
    ClipText = mvarClipText
End Property
Public Property Get IsImage() As Boolean
On Error Resume Next
    IsImage = mvarIsImage
End Property
Public Property Set image(ByVal vData As Object)
On Error Resume Next
    Set IMG = vData
    mvarIsImage = True
End Property
Public Property Get image() As Object
On Error Resume Next
    Set image = IMG
End Property
Private Sub Class_Initialize()
On Error Resume Next
    mvarIsImage = False
End Sub
Private Sub Class_Terminate()
On Error Resume Next
    Set IMG = Nothing
End Sub
