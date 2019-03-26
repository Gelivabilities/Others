VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   2460
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text1.Text = strGetIpAdress
End Sub
Private Function strGetIpAdress() As String
Dim wsShell, re, myIp, r, strLine
Set wsShell = CreateObject("WScript.Shell")
Set re = CreateObject("vbScript.RegExp")
re.Pattern = "IP Address"
Set myIp = wsShell.Exec("ipconfig /all")
While Not myIp.StdOut.AtEndOfStream
strLine = myIp.StdOut.ReadLine()
r = re.Test(strLine)
If r Then
strGetIpAdress = Mid(strLine, InStrRev(strLine, ":") + 1)
End If
Wend
End Function

