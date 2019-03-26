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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   1680
   End
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


Public v, stopTime, waitTime As Double

Private Sub Form_load()
v = 300
waitTime = 0
End Sub

Private Sub Timer1_Timer()
    Form5.Left = Form5.Left - v
    If stopTime = 0 Then
        v = v - 2.36
    End If
    If Form5.Left <= 645 Then
        Form4.Top = 375
        Form5.Left = 645
        
        On Error Resume Next
        Dim i As Long

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
        
        Timer2.Enabled = True
        
        Form5.Left = 19845
        If Form2.coins >= Int(Form2.Text7.Text) Then
            Form2.Command3.Enabled = True
        Else
            Form4.Image2.Picture = LoadPicture(App.Path & "\Image\insertcoins.bmp")
            Form4.Image3.Picture = LoadPicture(App.Path & "\Image\" & Int(Form2.Text7.Text) - coins & ".bmp")
        End If
        Form2.Command4.Enabled = True
        Form2.Command5.Enabled = True
        Form2.Command6.Enabled = True
        Form2.Command7.Enabled = True
        Form2.jiroType = 0
        Form2.Label13.Caption = "左右交替敲10次开启下一个次郎"
        Form2.Text1.Enabled = True
        Shell "taskkill.exe /f /im " & Form2.Text4(Form2.jiroType).Text
        Form1.Timer1.Enabled = False
        For i = 0 To 2
            Form2.Text2(i).Enabled = False
            Form2.Text4(i).Enabled = False
        Next
        Form5.Top = -99999
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
waitTime = waitTime + 1
If waitTime = 2 Then
    Form2.youxizhong = False
    waitTime = 0
    Timer2.Enabled = False
End If
End Sub
