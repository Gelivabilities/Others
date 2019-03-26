VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   6645
   ClientTop       =   6780
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub

