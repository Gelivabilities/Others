VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   975
   ClientLeft      =   9225
   ClientTop       =   5550
   ClientWidth     =   1455
   LinkTopic       =   "Form5"
   ScaleHeight     =   975
   ScaleWidth      =   1455
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub


