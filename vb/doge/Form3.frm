VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   8010
   ClientTop       =   5415
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub


