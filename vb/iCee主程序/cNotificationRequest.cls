VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNotificationRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Key                   As String
Private m_Icon                  As Object
Private m_Title                 As String
Private m_Description           As String
Private m_EnableClickEvent      As Boolean
Private m_SoundFileLocation     As String
Public Property Get Key() As String
    Key = m_Key
End Property
Public Property Let Key(ByVal vNewValue As String)
    m_Key = vNewValue
End Property

Public Property Get Icon() As Object
    Set Icon = m_Icon
End Property
Public Property Let Icon(ByVal vNewValue As Object)
    Set m_Icon = vNewValue
End Property

Public Property Get Title() As String
    Title = m_Title
End Property
Public Property Let Title(ByVal vNewValue As String)
    m_Title = vNewValue
End Property

Public Property Get Description() As String
    Description = m_Description
End Property
Public Property Let Description(ByVal vNewValue As String)
    m_Description = vNewValue
End Property

Public Property Get EnableClickEvent() As Boolean
    EnableClickEvent = m_EnableClickEvent
End Property
Public Property Let EnableClickEvent(ByVal vNewValue As Boolean)
    m_EnableClickEvent = vNewValue
End Property
