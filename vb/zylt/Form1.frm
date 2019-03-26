VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "LYC"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   240
      Top             =   4560
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   0
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Left            =   840
      Top             =   4080
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   99999
      TabIndex        =   0
      Text            =   "0"
      Top             =   99999
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   840
      Top             =   2160
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   360
      Shape           =   3  'Circle
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Text1.Text = Text1.Text + 10
Shape1.Top = Shape1.Top + Text1.Text
If Shape1.Top + Text1.Text > Form1.Height - Shape1.Height - Label1.Height - 500 Then
Timer2.Interval = 20
Timer1.Interval = 0
End If
End Sub

Private Sub Timer2_Timer()
Text1.Text = Text1.Text - 10
Shape1.Top = Shape1.Top - Text1.Text - 10
If Text1.Text < 0 Then
Text1.Text = 0
Shape1.Top = 0
Timer1.Interval = 20
Timer2.Interval = 0
End If
End Sub

Private Sub Timer3_Timer()
If Form1.Width > 1980 Then
Form1.Width = Form1.Width - 200
Form1.Left = Form1.Left + 100
Else
If Form1.Width < 1980 Then
Form1.Width = Form1.Width + 5
Form1.Left = Form1.Left - 5
End If
End If
End Sub



Private Sub Timer4_Timer()
Label1.Caption = "v=" & Text1.Text / 100 & "dm/s"
Label1.Top = Form1.Height - Label1.Height - 500
If Text1.Text / 100 < 1 And Text1.Text / 100 > 0 Then
Label1.Caption = "v=0" & Text1.Text / 100 & "dm/s"
Else
If Text1.Text Mod 100 = 0 Then
Label1.Caption = "v=" & Text1.Text / 100 & ".0dm/s"
End If
End If
End Sub
