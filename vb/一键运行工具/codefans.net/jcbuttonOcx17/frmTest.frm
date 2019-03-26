VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test it"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9180
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   612
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CheckBox chkFocusRect 
      Caption         =   "Show Focus Rect"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test This"
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
      Begin prjButton.jcbutton jcbutton1 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         buttonstyle     =   13
         font            =   "frmTest.frx":0000
         backcolor       =   14935011
         caption         =   "Toggle Underline"
         captioneffects  =   0
         tooltiptitle    =   "Toggle Underline"
         tooltiptype     =   1
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Default Button"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin prjButton.jcbutton jcbutton 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         buttonstyle     =   2
         font            =   "frmTest.frx":002C
         backcolor       =   15199212
         caption         =   "&Test Me"
         captioneffects  =   0
         tooltiptype     =   1
      End
      Begin prjButton.jcbutton jcbutton4 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         buttonstyle     =   13
         font            =   "frmTest.frx":0054
         backcolor       =   14935011
         caption         =   "Toggle Size"
         captioneffects  =   0
         tooltiptitle    =   "Toggle Size"
         tooltiptype     =   1
      End
      Begin prjButton.jcbutton jcbutton5 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         buttonstyle     =   13
         font            =   "frmTest.frx":0080
         backcolor       =   14935011
         caption         =   "Toggle Bold"
         captioneffects  =   0
         tooltiptitle    =   "Toggle Bold"
         tooltiptype     =   1
         colorscheme     =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Some Buttons"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
      Begin prjButton.jcbutton jcbutton2 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         buttonstyle     =   3
         font            =   "frmTest.frx":00AC
         backcolor       =   14935011
         caption         =   "Vista"
         captioneffects  =   0
         tooltiptype     =   1
      End
   End
   Begin VB.TextBox txtTest 
      Height          =   5535
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   6915
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.codefans.net
Option Explicit

Private Sub ChkCancel_Click()
    jcbutton.Cancel = Not jcbutton.Cancel
End Sub

Private Sub chkDefault_Click()
    jcbutton.Default = Not jcbutton.Default
End Sub

Private Sub chkEnabled_Click()
    jcbutton.Enabled = Not jcbutton.Enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmButtonDemo.Show
End Sub

Private Sub jcbutton_Click()
    MsgBox "The  [Test Me]  button should be now in IDLE state with no focus" & vbNewLine & _
    "Close this messagebox and the button should gain focus.", vbInformation + vbOKOnly, "Clicked"
End Sub

Private Sub jcbutton1_Click()
    jcbutton.Font.Underline = Not jcbutton.Font.Underline
End Sub

Private Sub jcbutton4_Click()
    jcbutton.Font.Size = IIf(jcbutton.Font.Size = 8.25, 10, 8.25)
End Sub

Private Sub jcbutton5_Click()
    jcbutton.Font.Bold = Not jcbutton.Font.Bold
End Sub

Private Sub Form_Load()

Dim s As String
    
    ' --Test according to ccXPButton (Have you seen it??) code id:-57148
    ' --Wonderful button for XP lovers
    s = " 1. Click the Vista Button as fast as you can." & vbNewLine & _
        "     - The button should follow the up-down state correctly as you click." & vbNewLine & vbNewLine & _
        " 2. Focus on TestMe button. Hold the ALT button, press SPACEBAR without releasing ALT." & vbNewLine & _
        "     - The system menu should be displayed and response to no event." & vbNewLine & vbNewLine & _
        " 3. Click the mouse, hold it down, and move in and out of the button." & vbNewLine & _
        "     - The button should cycle up and down as the mouse enters and exits." & vbNewLine & vbNewLine & _
        " 4. Click the mouse, hold it down, move out of the button and release the mouse." & vbNewLine & _
        "     - This should not induce a click event." & vbNewLine & vbNewLine & _
        " 5. Click the mouse, hold it down, move out then back in the button and release the mouse." & vbNewLine & _
        "     - This should induce a click event." & vbNewLine & vbNewLine & _
        " 6. Mouse click disable, move the mouse over the button and enable it using the spacebar." & vbNewLine & _
        "    - The button should return to the 'Hot' (or mouse over) mode." & vbNewLine & vbNewLine & _
        " 7. Click the button, move the Msgbox's OK button over the XP button, and click OK." & vbNewLine & _
        "     - The button should return to the 'Hot' (or mouse over) mode." & vbNewLine & vbNewLine & _
        " 8. Click the button, hold it down, and hit the Tab key." & vbNewLine & _
        "     - The button should return to the 'Hot' or 'Idle' mode with no click event." & vbNewLine & vbNewLine & _
        " 9. Using the Tab or arrows keys try to change focus between the button and checkbox." & vbNewLine & _
        "     - Focus should alternate between the two controls." & vbNewLine & vbNewLine & _
        "10. Select Alt + T while focus is on any control on the form." & vbNewLine & _
        "     - The click event should be raised and the button should receive the focus." & vbNewLine & vbNewLine & _
        "11. Focus on a button, hold the SPACEBAR down. While button stays on down state," & vbNewLine & _
        "     - click anywhere on the window. If you click on a control," & vbNewLine & _
        "       it may respond on the mouseover event but doesn't transfer focus to control." & vbNewLine & vbNewLine & _
        "12. And finally, select the default checkbox and change focus between controls." & vbNewLine & _
        "     - The [Test Me] button should always draw default unless the focus is on another button." & vbNewLine
        
    txtTest.Text = s
    
End Sub

