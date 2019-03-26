VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   10800
   ClientLeft      =   19845
   ClientTop       =   420
   ClientWidth     =   19200
   LinkTopic       =   "Form5"
   ScaleHeight     =   10800
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   1680
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   0
      Top             =   0
      Width           =   19200
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public v, stopTime As Double

Private Sub form_load()
v = 300
End Sub

Private Sub Timer1_Timer()
    Form5.Left = Form5.Left - v
    If stopTime = 0 Then
        v = v - 2.36
    End If
    If Form5.Left <= 645 Then
        Form4.Top = 375
        Form5.Left = 645
        v = 0
        stopTime = stopTime + 1
        Else
            stopTime = 0
    End If
    
    If stopTime = 150 Then
        Form5.Left = 660
        v = -500
    End If
    
    If Form5.Left > 19845 Then
    Form5.Left = 19845
    Timer1.Enabled = False
    End If
End Sub
