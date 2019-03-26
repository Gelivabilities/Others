VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   975
   ClientLeft      =   8415
   ClientTop       =   6915
   ClientWidth     =   1095
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   975
   ScaleWidth      =   1095
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "Form4.frx":9F33
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub


