VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   10800
   ClientLeft      =   645
   ClientTop       =   360
   ClientWidth     =   19200
   LinkTopic       =   "Form4"
   ScaleHeight     =   10800
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image3 
      Height          =   1215
      Left            =   9480
      Top             =   8880
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   5715
      Top             =   8340
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Top             =   0
      Width           =   20490
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_load()
Form1.Show vbModeless, Form4
End Sub

