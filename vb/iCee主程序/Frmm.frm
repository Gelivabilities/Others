VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Frmm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFBC44&
   Caption         =   "�˵�"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   1080
   ClientWidth     =   15690
   Icon            =   "Frmm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1046
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox SKINDRAW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00207D4F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   8520
      ScaleHeight     =   625
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   380
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   5700
      Begin VB.ListBox LSTLINK 
         Appearance      =   0  'Flat
         Height          =   210
         Left            =   2280
         TabIndex        =   62
         Top             =   5760
         Width           =   2055
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   1830
         Hidden          =   -1  'True
         Left            =   3120
         TabIndex        =   61
         Top             =   6000
         Width           =   2055
      End
      Begin VB.ListBox LISTM 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2280
         TabIndex        =   60
         Top             =   5520
         Width           =   2055
      End
      Begin SHDocVwCtl.WebBrowser WB 
         Height          =   2895
         Left            =   240
         TabIndex        =   59
         Top             =   120
         Width           =   2895
         ExtentX         =   5106
         ExtentY         =   5106
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox SKINLINE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9375
      Left            =   7200
      ScaleHeight     =   625
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   380
      TabIndex        =   70
      Top             =   1200
      Width           =   5700
      Begin VB.PictureBox PTCO 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002A1C05&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1440
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   80
         Top             =   3000
         Width           =   975
      End
      Begin VB.PictureBox PICCOLOR 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002A1C05&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   360
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   79
         Top             =   3000
         Width           =   975
      End
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   990
      Index           =   30
      Left            =   0
      Picture         =   "Frmm.frx":5E62
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   205
      Top             =   5760
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   990
      Index           =   29
      Left            =   0
      Picture         =   "Frmm.frx":8728
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   204
      Top             =   3600
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   990
      Index           =   28
      Left            =   1080
      Picture         =   "Frmm.frx":AE01
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   203
      Top             =   5760
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   990
      Index           =   23
      Left            =   1080
      Picture         =   "Frmm.frx":D11C
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   202
      Top             =   3600
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   990
      Index           =   22
      Left            =   0
      Picture         =   "Frmm.frx":F96C
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   201
      Top             =   4680
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   990
      Index           =   3
      Left            =   1080
      Picture         =   "Frmm.frx":12266
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   200
      Top             =   4680
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   19
      Left            =   6000
      Picture         =   "Frmm.frx":14A45
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   199
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   18
      Left            =   5640
      Picture         =   "Frmm.frx":151AF
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   198
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   15
      Left            =   5280
      Picture         =   "Frmm.frx":15919
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   197
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   8
      Left            =   4920
      Picture         =   "Frmm.frx":16083
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   196
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   6
      Left            =   4560
      Picture         =   "Frmm.frx":167ED
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   195
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   2
      Left            =   4680
      Picture         =   "Frmm.frx":16F57
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   194
      Top             =   5400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox PS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0084536F&
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   1
      Left            =   7560
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   94
      Top             =   5160
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002D3855&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   30
         Left            =   4920
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   193
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H005586C9&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   29
         Left            =   5400
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   192
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00DFE1FD&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   5880
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   191
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00BFABF8&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   190
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H005D73A3&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   189
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00919BC6&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   188
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AAB1D8&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   187
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E5EFF9&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   186
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00A3B0FA&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   3960
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   185
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00CADDFE&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   184
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00DFF5FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   183
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00557296&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   182
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H007DA7DC&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   181
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00BDDDFA&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   180
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H006D8D98&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   179
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00A1D2E2&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   178
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00B5EAFE&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   177
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00DCC49E&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   70
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   176
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00857938&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   69
         Left            =   3960
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   175
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002F1AA4&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   68
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   174
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0099CA3E&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   67
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   173
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00A5A400&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   66
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   172
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0069A400&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   65
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   171
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0027ADED&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   64
         Left            =   4920
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   170
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AAE186&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   63
         Left            =   4920
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   169
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H007AEDDC&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   62
         Left            =   4920
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   168
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0022B380&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   61
         Left            =   5400
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   167
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H009A4FDF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   60
         Left            =   5400
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   166
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H009EB678&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   59
         Left            =   5880
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   165
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00B3A562&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   58
         Left            =   5880
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   164
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H009C7738&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   57
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   163
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H007DEACC&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   56
         Left            =   3960
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   162
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001EF0A7&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   55
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   161
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00DF4F5E&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   54
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   160
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00A65F09&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   53
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   159
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00949C72&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   52
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   158
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00EAB97D&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   51
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   157
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00D29773&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   50
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   156
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00D1F3EF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   49
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   155
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H004A3F7E&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   48
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   154
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00BC76FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   28
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   153
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000D191&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   27
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   152
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000172D&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   26
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   151
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00002F63&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   25
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   150
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFD0A7&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   24
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   149
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0046265F&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   23
         Left            =   3960
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   148
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H009C6813&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   22
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   147
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00239989&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   21
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   146
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00BAD1FE&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   20
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   145
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00937BFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   19
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   144
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002D23F0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   18
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   143
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H000807BD&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   17
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   142
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00080615&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   141
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0054B4C9&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   140
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0074E2DD&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   98
         Left            =   4920
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   139
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H001F1FE2&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   97
         Left            =   5400
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   138
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002EBC7C&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   96
         Left            =   5880
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   137
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0028BEFD&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   47
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   136
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H007C63BD&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   46
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   135
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00483280&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   45
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   134
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00B9DFC3&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   44
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   133
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00CBD4D7&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   43
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   132
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002D6338&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   42
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   131
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H006F7240&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   41
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   130
         Top             =   2520
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H002A1E12&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   40
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   129
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C6C397&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   39
         Left            =   3960
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   128
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00CDE1CA&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   38
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   127
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H006495A3&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   37
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   126
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0043D2F2&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   36
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   125
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H006E5EE1&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   35
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   124
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00A7B984&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   34
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   123
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00224F3C&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   33
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   122
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H004BA080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   32
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   121
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H006FD1C5&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   31
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   120
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00252525&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   95
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   119
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00B864E0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   94
         Left            =   3960
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   118
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00237EFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   93
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   117
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00122DFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   92
         Left            =   4920
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   116
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00771DFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   91
         Left            =   5400
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   115
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H003E11AE&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   90
         Left            =   4920
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   114
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00CBD800&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   89
         Left            =   5400
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   113
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AAAA00&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   88
         Left            =   5880
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   112
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H004F00C1&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   87
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   111
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000004E&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   86
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   110
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00011EAF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   85
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   109
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0020BA83&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   84
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   108
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H003ED191&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   83
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   107
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC456&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   82
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   106
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H005F4D00&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   81
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   105
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00EB7327&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   80
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   104
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00B8581C&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   79
         Left            =   3960
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   103
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00878201&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   78
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   102
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000BA78&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   77
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   101
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFAD1F&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   76
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   100
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H004E1F00&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   75
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   99
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0001B3F4&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   74
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   98
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00B71B68&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   73
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   97
         Top             =   3000
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C06A00&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   72
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   96
         Top             =   3360
         Width           =   495
      End
      Begin VB.PictureBox PCO 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00252525&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   71
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   95
         Top             =   3000
         Width           =   495
      End
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   26
      Left            =   480
      Picture         =   "Frmm.frx":176C1
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   93
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   24
      Left            =   480
      Picture         =   "Frmm.frx":17A4B
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   92
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   42
      Left            =   3480
      Picture         =   "Frmm.frx":17DD5
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   91
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   41
      Left            =   3240
      Picture         =   "Frmm.frx":1815F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   90
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   36
      Left            =   3000
      Picture         =   "Frmm.frx":184E9
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   89
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   34
      Left            =   2760
      Picture         =   "Frmm.frx":18873
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   88
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   32
      Left            =   2520
      Picture         =   "Frmm.frx":18BFD
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   87
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   25
      Left            =   1800
      Picture         =   "Frmm.frx":18F87
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   86
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   9
      Left            =   2040
      Picture         =   "Frmm.frx":19311
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   85
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   1800
      Picture         =   "Frmm.frx":1969B
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   84
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   2040
      Picture         =   "Frmm.frx":19A25
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   83
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   1800
      Picture         =   "Frmm.frx":19DAF
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   82
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   2040
      Picture         =   "Frmm.frx":1A139
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   81
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox SKINEMAIL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   5760
      ScaleHeight     =   625
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   380
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   5700
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   4800
         Index           =   152
         Left            =   480
         ScaleHeight     =   4800
         ScaleWidth      =   4800
         TabIndex        =   50
         Top             =   480
         Width           =   4800
         Begin VB.PictureBox PSINGER 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00DBA349&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6000
            Left            =   0
            ScaleHeight     =   400
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   400
            TabIndex        =   58
            Top             =   4680
            Width           =   6000
            Begin VB.Image X1 
               Height          =   240
               Left            =   0
               Picture         =   "Frmm.frx":1A4C3
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image X2 
               Height          =   240
               Left            =   1080
               Picture         =   "Frmm.frx":1A84D
               Top             =   120
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image X3 
               Height          =   240
               Left            =   1320
               Picture         =   "Frmm.frx":1ABD7
               Top             =   120
               Visible         =   0   'False
               Width           =   240
            End
         End
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H002EBC7C&
            BorderStyle     =   0  'None
            Height          =   150
            Index           =   183
            Left            =   240
            ScaleHeight     =   10
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   10
            TabIndex        =   57
            Top             =   120
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00207D4F&
            BorderStyle     =   0  'None
            Height          =   150
            Index           =   184
            Left            =   120
            ScaleHeight     =   10
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   10
            TabIndex        =   56
            Top             =   120
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H003F3B36&
            BorderStyle     =   0  'None
            Height          =   720
            Index           =   197
            Left            =   2160
            Picture         =   "Frmm.frx":1AF61
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   55
            Top             =   3720
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H003F3B36&
            BorderStyle     =   0  'None
            Height          =   720
            Index           =   198
            Left            =   1440
            Picture         =   "Frmm.frx":1B08C
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   54
            Top             =   3720
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H003F3B36&
            BorderStyle     =   0  'None
            Height          =   720
            Index           =   199
            Left            =   2880
            Picture         =   "Frmm.frx":1B1D7
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   53
            Top             =   3720
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H003F3B36&
            BorderStyle     =   0  'None
            Height          =   720
            Index           =   200
            Left            =   3600
            Picture         =   "Frmm.frx":1B2E3
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   52
            Top             =   3720
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H003F3B36&
            BorderStyle     =   0  'None
            Height          =   720
            Index           =   201
            Left            =   4320
            Picture         =   "Frmm.frx":1B3DF
            ScaleHeight     =   48
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   48
            TabIndex        =   51
            Top             =   3720
            Visible         =   0   'False
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   54
      Left            =   1200
      Picture         =   "Frmm.frx":1B4F8
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   78
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   52
      Left            =   960
      Picture         =   "Frmm.frx":1B882
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   77
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   43
      Left            =   3600
      Picture         =   "Frmm.frx":1BC0C
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   76
      Top             =   6240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   33
      Left            =   2880
      Picture         =   "Frmm.frx":1D836
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   75
      Top             =   6240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   21
      Left            =   2160
      Picture         =   "Frmm.frx":1F460
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   74
      Top             =   6240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   20
      Left            =   2160
      Picture         =   "Frmm.frx":2108A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   73
      Top             =   6960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   11
      Left            =   2880
      Picture         =   "Frmm.frx":2111E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   72
      Top             =   6960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   10
      Left            =   3600
      Picture         =   "Frmm.frx":211B2
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   71
      Top             =   6960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox SKINMUSIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00955C00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   4560
      ScaleHeight     =   625
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   380
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   5700
      Begin MSWinsockLib.Winsock wsListen 
         Left            =   1440
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008BA31F&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   16
      Left            =   0
      Picture         =   "Frmm.frx":21246
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   69
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox IMBK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H007550D4&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   7080
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   68
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   13
      Left            =   960
      Picture         =   "Frmm.frx":215D0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   67
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   12
      Left            =   1200
      Picture         =   "Frmm.frx":2195A
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   66
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   47
      Left            =   3720
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   65
      Top             =   3360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   46
      Left            =   3120
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   64
      Top             =   3360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   720
      Index           =   45
      Left            =   2520
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   63
      Top             =   3360
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   182
      Left            =   1680
      Picture         =   "Frmm.frx":21CE4
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2
      TabIndex        =   48
      Top             =   4320
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   181
      Left            =   1560
      Picture         =   "Frmm.frx":21DF0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2
      TabIndex        =   47
      Top             =   4320
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   180
      Left            =   1200
      Picture         =   "Frmm.frx":21F24
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   46
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   179
      Left            =   1200
      Picture         =   "Frmm.frx":222AE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   178
      Left            =   840
      Picture         =   "Frmm.frx":22638
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   44
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   177
      Left            =   840
      Picture         =   "Frmm.frx":229C2
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   43
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   176
      Left            =   1920
      Picture         =   "Frmm.frx":22D4C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   42
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   175
      Left            =   1920
      Picture         =   "Frmm.frx":230D6
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   41
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   174
      Left            =   1560
      Picture         =   "Frmm.frx":23460
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   40
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00201400&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   173
      Left            =   1560
      Picture         =   "Frmm.frx":237EA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   39
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   147
      Left            =   4680
      Picture         =   "Frmm.frx":23B74
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   38
      Top             =   3360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   146
      Left            =   5520
      Picture         =   "Frmm.frx":23F11
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   37
      Top             =   3360
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   38
      Left            =   1200
      Picture         =   "Frmm.frx":24281
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   37
      Left            =   960
      Picture         =   "Frmm.frx":2460B
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   131
      Left            =   6360
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   130
      Left            =   6360
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   33
      Top             =   1800
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox PICCLIP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   4800
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   32
      Top             =   2160
      Width           =   855
   End
   Begin VB.Timer TMA 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   1080
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   119
      Left            =   1080
      Picture         =   "Frmm.frx":24995
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   40
      Left            =   720
      Picture         =   "Frmm.frx":250FF
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   30
      Top             =   5520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   35
      Left            =   360
      Picture         =   "Frmm.frx":25869
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   29
      Top             =   5520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Timer TMMU 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3240
      Top             =   1080
   End
   Begin VB.Timer TMEA 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   98
      Left            =   4320
      Picture         =   "Frmm.frx":25FD3
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   28
      Top             =   5400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   97
      Left            =   3960
      Picture         =   "Frmm.frx":2673D
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   27
      Top             =   5400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   96
      Left            =   3600
      Picture         =   "Frmm.frx":26EA7
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   26
      Top             =   5400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   95
      Left            =   3000
      Picture         =   "Frmm.frx":27611
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   25
      Top             =   5760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   94
      Left            =   2640
      Picture         =   "Frmm.frx":27CFB
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   24
      Top             =   5760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   93
      Left            =   2280
      Picture         =   "Frmm.frx":283E5
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   23
      Top             =   5760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   92
      Left            =   3000
      Picture         =   "Frmm.frx":28ACF
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   91
      Left            =   2640
      Picture         =   "Frmm.frx":291B9
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   90
      Left            =   2280
      Picture         =   "Frmm.frx":298A3
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   20
      Top             =   5040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   83
      Left            =   360
      Picture         =   "Frmm.frx":29F8D
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   82
      Left            =   720
      Picture         =   "Frmm.frx":2A6F7
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   81
      Left            =   1080
      Picture         =   "Frmm.frx":2AE61
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   68
      Left            =   960
      Picture         =   "Frmm.frx":2B5CB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   17
      Left            =   1200
      Picture         =   "Frmm.frx":2B955
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer TMRCPU 
      Interval        =   100
      Left            =   3720
      Top             =   1080
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   39
      Left            =   2880
      Picture         =   "Frmm.frx":2BCDF
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox da2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   2760
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   10
      ToolTipText     =   "���˵�"
      Top             =   4320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox da3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   3360
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   9
      ToolTipText     =   "���˵�"
      Top             =   4320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox da1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   2160
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   8
      ToolTipText     =   "���˵�"
      Top             =   4320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Timer TimeHon 
      Interval        =   10
      Left            =   4200
      Top             =   360
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008BA31F&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   27
      Left            =   0
      Picture         =   "Frmm.frx":2C069
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox OFFLINE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3120
      Picture         =   "Frmm.frx":2C3F3
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox BusyNow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3600
      Picture         =   "Frmm.frx":2C77D
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Away 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3600
      Picture         =   "Frmm.frx":2CB07
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox HideNow 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3120
      Picture         =   "Frmm.frx":2CE91
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox ONLINE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3360
      Picture         =   "Frmm.frx":2D21B
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008BA31F&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   14
      Left            =   0
      Picture         =   "Frmm.frx":2D5A5
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008BA31F&
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   0
      Picture         =   "Frmm.frx":2D92F
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LBLINK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ټ�,�������ӵ���,�°汾,ʹ�Ҹ�����"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   49
      Top             =   2760
      Width           =   3150
   End
   Begin VB.Image IMGAD 
      Height          =   3135
      Left            =   9480
      Top             =   360
      Width           =   2460
   End
   Begin VB.Menu �ı� 
      Caption         =   "�ı�"
      Begin VB.Menu ȫѡ�ı� 
         Caption         =   "ȫѡ"
      End
      Begin VB.Menu �����ı� 
         Caption         =   "����"
      End
      Begin VB.Menu �����ı� 
         Caption         =   "����"
      End
      Begin VB.Menu ճ���ı� 
         Caption         =   "ճ��"
      End
      Begin VB.Menu ɾ���ı� 
         Caption         =   "ɾ��"
      End
   End
   Begin VB.Menu ͼ���� 
      Caption         =   "ͼƬ����"
      Begin VB.Menu oxox 
         Caption         =   "Ԥ��ͼƬ"
         Index           =   0
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu oxox 
         Caption         =   "������Ϊͷ��"
         Index           =   2
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu oxox 
         Caption         =   "��תͼƬ"
         Index           =   4
      End
      Begin VB.Menu oxox 
         Caption         =   "�����ع�"
         Index           =   5
      End
      Begin VB.Menu oxox 
         Caption         =   "ʹ���˾�"
         Index           =   6
      End
      Begin VB.Menu oxox 
         Caption         =   "�����ַ���"
         Index           =   7
      End
      Begin VB.Menu oxox 
         Caption         =   "Ϳѻ"
         Index           =   8
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu oxox 
         Caption         =   "����"
         Index           =   10
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu oxox 
         Caption         =   "ͼ������"
         Index           =   12
      End
   End
   Begin VB.Menu ���ſ��� 
      Caption         =   "����"
      Begin VB.Menu ��ý�� 
         Caption         =   "���ļ�"
      End
      Begin VB.Menu ��ӵ��� 
         Caption         =   "����ļ�"
      End
      Begin VB.Menu ���ļ��� 
         Caption         =   "���ļ���"
      End
      Begin VB.Menu ����ļ��� 
         Caption         =   "����ļ���"
      End
      Begin VB.Menu swff01122 
         Caption         =   "-"
      End
      Begin VB.Menu �ղؼ� 
         Caption         =   "�ղؼ�"
      End
      Begin VB.Menu ��URL 
         Caption         =   "���ִ�"
      End
      Begin VB.Menu wqE 
         Caption         =   "-"
      End
      Begin VB.Menu �򿪲����б� 
         Caption         =   "���벥���б�"
      End
      Begin VB.Menu ���� 
         Caption         =   "���������б�"
      End
      Begin VB.Menu SHDH 
         Caption         =   "-"
      End
      Begin VB.Menu ��ͬ�� 
         Caption         =   "��ͬ���б�"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "��������"
      Begin VB.Menu ������������ 
         Caption         =   "����"
      End
      Begin VB.Menu ������б� 
         Caption         =   "��ӵ������б�"
      End
      Begin VB.Menu ���� 
         Caption         =   "����ѡ������"
      End
      Begin VB.Menu SFEGH 
         Caption         =   "-"
      End
      Begin VB.Menu ���ŵ�ǰ�б� 
         Caption         =   "���ŵ�ǰ�б�"
      End
   End
   Begin VB.Menu �ļ� 
      Caption         =   "�ļ�"
      Begin VB.Menu ����ѡ�� 
         Caption         =   "����ѡ�еĸ���"
      End
      Begin VB.Menu ɾ��ѡ�� 
         Caption         =   "ɾ��ѡ�еĸ���"
      End
      Begin VB.Menu ����վ 
         Caption         =   "��������վ"
      End
      Begin VB.Menu jdhj 
         Caption         =   "-"
      End
      Begin VB.Menu ȥ�� 
         Caption         =   "ȥ���ظ�"
      End
      Begin VB.Menu ɾ����Ч 
         Caption         =   "ɾ����Ч������"
      End
      Begin VB.Menu ˢ���б� 
         Caption         =   "ˢ���б�"
      End
      Begin VB.Menu ajkhf 
         Caption         =   "-"
      End
      Begin VB.Menu �������� 
         Caption         =   "�鿴ѡ���ļ�����"
      End
      Begin VB.Menu λ�� 
         Caption         =   "���ļ�����λ��"
      End
      Begin VB.Menu SFFGGHH 
         Caption         =   "-"
      End
      Begin VB.Menu ����·�� 
         Caption         =   "�����ļ�·��"
      End
      Begin VB.Menu ���������� 
         Caption         =   "����������"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu Ĭ�ϳ���� 
         Caption         =   "ʹ��ϵͳĬ�ϳ����"
      End
      Begin VB.Menu afhhjj 
         Caption         =   "-"
      End
      Begin VB.Menu �ָ��ļ� 
         Caption         =   "�ļ�����"
      End
   End
   Begin VB.Menu ˳�� 
      Caption         =   "����˳��"
      Begin VB.Menu ����ѭ�� 
         Caption         =   "�����ظ�"
      End
      Begin VB.Menu ˳�򲥷� 
         Caption         =   "˳�򲥷�"
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
      End
      Begin VB.Menu ѭ�� 
         Caption         =   "ѭ������"
      End
   End
   Begin VB.Menu iM 
      Caption         =   "IM"
      Begin VB.Menu mnuStatusOnline 
         Caption         =   "����"
      End
      Begin VB.Menu mnuStatusAway 
         Caption         =   "�뿪"
      End
      Begin VB.Menu mnuStatusDND 
         Caption         =   "��Ҫ����"
      End
      Begin VB.Menu mnuStatusInvisible 
         Caption         =   "����"
      End
      Begin VB.Menu UNLOGIN 
         Caption         =   "ע��"
      End
      Begin VB.Menu SSAD 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "����iCee"
      End
   End
   Begin VB.Menu mnuBuddy 
      Caption         =   "����"
      Begin VB.Menu mnuBuddyMessage 
         Caption         =   "������Ϣ"
      End
      Begin VB.Menu mnuBuddyChat 
         Caption         =   "��ʱ����"
      End
      Begin VB.Menu Զ��Э�� 
         Caption         =   "Զ��Э��"
      End
      Begin VB.Menu ��Ƶ���� 
         Caption         =   "��Ƶ����"
      End
      Begin VB.Menu mnuBuddyFile 
         Caption         =   "�����ļ�"
      End
      Begin VB.Menu mnuBuddyInfo 
         Caption         =   "�鿴����ע����Ϣ"
      End
      Begin VB.Menu ASWF 
         Caption         =   "-"
      End
      Begin VB.Menu ˢ�º����б� 
         Caption         =   "ˢ���б�"
      End
      Begin VB.Menu DFGR 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuddyRemove 
         Caption         =   "ɾ������"
      End
      Begin VB.Menu �ٱ� 
         Caption         =   "�ٱ����û�"
      End
      Begin VB.Menu ���� 
         Caption         =   "���θ��û�"
      End
      Begin VB.Menu mnuBuddyIgnore 
         Caption         =   "����������"
      End
      Begin VB.Menu SDWW 
         Caption         =   "-"
      End
      Begin VB.Menu �����¼ 
         Caption         =   "�鿴�����¼"
      End
      Begin VB.Menu DAFE 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePassword 
         Caption         =   "�޸ĵ�¼����"
      End
      Begin VB.Menu mnuFileChangeInfo 
         Caption         =   "�޸ĸ�����Ϣ"
      End
      Begin VB.Menu xsdf 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "BUG����"
      End
   End
   Begin VB.Menu ��Ч 
      Caption         =   "Ϳѻ��Ч"
      Begin VB.Menu �� 
         Caption         =   "��ͼ��"
      End
      Begin VB.Menu ģ�� 
         Caption         =   "��ͼ��"
      End
      Begin VB.Menu ���� 
         Caption         =   "�������"
      End
      Begin VB.Menu ���� 
         Caption         =   "������Ч"
      End
      Begin VB.Menu �Ҷ� 
         Caption         =   "�һ�ͼ��"
      End
      Begin VB.Menu ��ת 
         Caption         =   "��ת��ɫ"
      End
      Begin VB.Menu ħ�� 
         Caption         =   "ħ��Ч��"
      End
      Begin VB.Menu �ͻ� 
         Caption         =   "�ͻ�Ч��"
      End
      Begin VB.Menu ľ�� 
         Caption         =   "ľ��Ч��"
      End
      Begin VB.Menu ���� 
         Caption         =   "����Ч��"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
      Begin VB.Menu ѣ�� 
         Caption         =   "ѣ��"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu ���ӶԱȶ� 
         Caption         =   "���ӶԱȶ�"
      End
      Begin VB.Menu DEES 
         Caption         =   "-"
      End
      Begin VB.Menu ˮƽ 
         Caption         =   "ˮƽ��ת"
      End
      Begin VB.Menu ��ֱ 
         Caption         =   "��ֱ��ת"
      End
      Begin VB.Menu ˫�� 
         Caption         =   "˫��ת"
      End
      Begin VB.Menu SWRR 
         Caption         =   "-"
      End
      Begin VB.Menu �Ӽ��а�ճ�� 
         Caption         =   "�Ӽ��а�ճ��"
      End
      Begin VB.Menu Ϳѻ���а� 
         Caption         =   "���Ƶ����а�"
      End
      Begin VB.Menu SERSFFG 
         Caption         =   "-"
      End
      Begin VB.Menu ��ɫ�߿� 
         Caption         =   "��Ӻ�ɫ�߿�"
      End
      Begin VB.Menu ����ɫ�ɰ� 
         Caption         =   "������ɫ�ɰ�"
      End
      Begin VB.Menu FGGHH 
         Caption         =   "-"
      End
      Begin VB.Menu ȥ��ͼ�� 
         Caption         =   "���ͼ��"
      End
      Begin VB.Menu SDFGGSAA 
         Caption         =   "-"
      End
      Begin VB.Menu ��ͼ�� 
         Caption         =   "��ͼ��"
      End
      Begin VB.Menu ���� 
         Caption         =   "����ͼ��"
      End
      Begin VB.Menu ����ѷ��� 
         Caption         =   "����"
      End
      Begin VB.Menu DAFFF 
         Caption         =   "-"
      End
      Begin VB.Menu ��Ϊͷ�� 
         Caption         =   "��Ϊͷ��"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ���� 
         Caption         =   "���Խ��˻�����Ϊ������"
      End
      Begin VB.Menu SWFF 
         Caption         =   "-"
      End
      Begin VB.Menu �������� 
         Caption         =   "������������"
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
      End
   End
   Begin VB.Menu �������� 
      Caption         =   "��������"
      Begin VB.Menu ��������� 
         Caption         =   "���������"
      End
      Begin VB.Menu sjkai 
         Caption         =   "-"
      End
      Begin VB.Menu ֹͣ���� 
         Caption         =   "ֹͣ����"
      End
      Begin VB.Menu SWGGSSAA 
         Caption         =   "-"
      End
      Begin VB.Menu ������ 
         Caption         =   "���ļ�"
      End
      Begin VB.Menu ��λ���� 
         Caption         =   "��λ�ļ�"
      End
      Begin VB.Menu WGJK 
         Caption         =   "-"
      End
      Begin VB.Menu ������������ 
         Caption         =   "������������"
      End
      Begin VB.Menu ɾ���������� 
         Caption         =   "ɾ����������"
      End
      Begin VB.Menu �������� 
         Caption         =   "������������"
      End
      Begin VB.Menu sakkii 
         Caption         =   "-"
      End
      Begin VB.Menu ���ɶ�ά�� 
         Caption         =   "�������Ӷ�ά��"
      End
      Begin VB.Menu LLIDI 
         Caption         =   "-"
      End
      Begin VB.Menu ������� 
         Caption         =   "��������б�"
      End
   End
   Begin VB.Menu ������� 
      Caption         =   "���������"
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu ɾ�������ļ� 
         Caption         =   "ɾ�������ļ�"
      End
      Begin VB.Menu SWFFW 
         Caption         =   "-"
      End
      Begin VB.Menu �����ļ� 
         Caption         =   "�����ļ�"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu SJWJ 
         Caption         =   "-"
      End
      Begin VB.Menu ˢ�½��� 
         Caption         =   "ˢ���б�"
      End
   End
   Begin VB.Menu ��� 
      Caption         =   "���"
      Begin VB.Menu �鿴��� 
         Caption         =   "�鿴���"
      End
      Begin VB.Menu ɾ������ 
         Caption         =   "ɾ������"
      End
      Begin VB.Menu �༭��� 
         Caption         =   "�༭���"
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
      End
      Begin VB.Menu SWWAGHGH 
         Caption         =   "-"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
   End
   Begin VB.Menu TCP 
      Caption         =   "TCP"
      Begin VB.Menu �Ͽ�TCP���� 
         Caption         =   "�Ͽ�����"
      End
   End
   Begin VB.Menu �ҵ��ղ� 
      Caption         =   "�ҵ��ղ�"
      Begin VB.Menu ����ղ� 
         Caption         =   "��ӵ������б�"
      End
      Begin VB.Menu ���ȫ���ղ� 
         Caption         =   "���ȫ��"
      End
      Begin VB.Menu ɾ���ղ� 
         Caption         =   "ɾ��"
      End
      Begin VB.Menu ����ղ� 
         Caption         =   "����ղ�"
      End
   End
   Begin VB.Menu �ļ����� 
      Caption         =   "�ļ�����"
      Begin VB.Menu ���ļ� 
         Caption         =   "���ļ�"
      End
      Begin VB.Menu SWFHKK 
         Caption         =   "-"
      End
      Begin VB.Menu ɾ���ļ� 
         Caption         =   "ɾ���ļ�"
      End
      Begin VB.Menu �����ļ� 
         Caption         =   "�����ļ�"
      End
      Begin VB.Menu ճ���ļ� 
         Caption         =   "ճ���ļ�"
      End
      Begin VB.Menu SJJJJAC 
         Caption         =   "-"
      End
      Begin VB.Menu �������ļ� 
         Caption         =   "�������ļ�"
      End
      Begin VB.Menu �������������ļ����µ��ļ� 
         Caption         =   "�������������ļ����µ��ļ�"
      End
      Begin VB.Menu JIIS 
         Caption         =   "-"
      End
      Begin VB.Menu ������ļ� 
         Caption         =   "������ļ�"
      End
      Begin VB.Menu safrr 
         Caption         =   "-"
      End
      Begin VB.Menu �ļ����� 
         Caption         =   "�ļ�����"
      End
      Begin VB.Menu �ļ������� 
         Caption         =   "�ļ�������"
      End
      Begin VB.Menu AAAWWW 
         Caption         =   "-"
      End
      Begin VB.Menu �½��ļ��� 
         Caption         =   "�½��ļ���"
      End
   End
   Begin VB.Menu ϵͳ���� 
      Caption         =   "ϵͳ����"
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
      Begin VB.Menu ��Ļ���� 
         Caption         =   "��Ļ����"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
      Begin VB.Menu QAASDFF 
         Caption         =   "-"
      End
      Begin VB.Menu �˴� 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Frmm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Dim PicBits() As Byte, PICINFO As BITMAP, Cnt As Long, ITIM As Long, LES
Private Sub Form_Load()
On Error Resume Next
Dim I As Integer
iStop = True
Call Init
Call SubClassWindow(Me)
If LONELY_MODE = True Then Exit Sub

wsListen.LocalPort = FT_USE_PORT
wsListen.Listen
Dim PFlowInfo As Flow_INFO, USEBACK As String
PFlowInfo = GetFlowInfo()
LastRecvBytes = PFlowInfo.lngBytesReceived
LastSentBytes = PFlowInfo.lngBytesSent

Call AddToTray(Me)
Call HookForm(Me)  '�����
SetTrayIcon OFFLINE.PICTURE

For I = 0 To 25
frmma.MBK(I).PICTURE = LoadPicture(App.Path & "\SKIN\BK\" & I & ".JPG")
Next
File1.Pattern = "*.mp3"
End Sub
Private Sub Form_Terminate()
Set Frmm = Nothing
End Sub
Private Sub mnuBuddyFile_Click()
On Error Resume Next
If frmma.TreeView1.SelectedItem.Key = "" Then Exit Sub
If frmma.TreeView1.SelectedItem.Text = frmma.Text1.Text Then Exit Sub
Call SendFile(frmma.TreeView1.SelectedItem.Key)
End Sub
Private Function fncGetInfo(lsPicName As String) As PICINFO '��ʹ�ÿؼ����ͼƬ��С
    Dim hBitmap As Long
    Dim res As Long
    Dim Bmp As BITMAP
    res = GetObject(LoadPicture(lsPicName).handle, Len(Bmp), Bmp) 'ȡ��BITMAP�Ľṹ
    fncGetInfo.PicWidth = Bmp.bmWidth
    fncGetInfo.PicHeight = Bmp.bmHeight
End Function

Private Sub OXOX_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Call frmGraphic.View_It(frmGraphic.Select_Pic)
Case 2
FRMHEAD.Show
Call FRMHEAD.OpenFile(frmGraphic.Select_Pic)
Case 4
If UCase(Right(frmGraphic.Select_Pic, 3)) = "PNG" Then Exit Sub
Call frmGraphic.��(frmGraphic.Select_Pic)
Call frmGraphic.pic_turn
Case 5
If UCase(Right(frmGraphic.Select_Pic, 3)) = "PNG" Then Exit Sub
Call frmGraphic.��(frmGraphic.Select_Pic)
Call frmGraphic.Pic_Talking
Case 6
If UCase(Right(frmGraphic.Select_Pic, 3)) = "PNG" Then Exit Sub
Call frmGraphic.��(frmGraphic.Select_Pic)
Call frmGraphic.pic_sun
Case 7
Call frmGraphic.Pic_TXT
Call frmGraphic.MyOpen(frmGraphic.Select_Pic)
Case 8
Call FRMBOARD.OpenFile(frmGraphic.Select_Pic)
FRMBOARD.Show
Case 10
DefCOM = 2
Call frmma.SHAREIT(frmGraphic.Select_Pic)
Case 12
FrmPicInfo.Show
End Select
End Sub



Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
  ConnectReq requestID
End Sub
Private Sub Init()
Dim hMainMenu As Long, HWEB As Long, hcol As Long, hSubMenu As Long, HFIVE As Long, HWQ As Long, HLOOK As Long, HNET As Long, HChat As Long, HTY As Long, hFtp As Long, Hlocal As Long, hTxt As Long, hPic As Long, Hlst As Long, hFr As Long, Hpla As Long, hFile As Long, hZor As Long, HFAV As Long, HABOUT As Long
'ȫ�ֱ����ĳ�ʼ��
g_clrFrame = &H8BA31F    'ѡ����Ŀ������ɫ
g_clrBkgSelect = &H8BA31F    ' RGB(93, 80, 58) 'ѡ����Ŀ��������ɫ
g_clrLeft = &H8BA31F '�˵���ߵ���ɫ

g_clrBkgNormal = vbWhite  ' RGB(253, 251, 250) '������������ɫ
g_clrTxtSelect = RGB(255, 255, 255) 'ѡ���ı�����ɫ
g_clrTxtNormal = RGB(0, 0, 0) '�����ı�����ɫ
g_clrSep = RGB(209, 209, 209) '�ָ��ߵ���ɫ

hMainMenu = GetMenu(Me.hwnd) '�õ����嶥���˵����
hSubMenu = GetSubMenu(hMainMenu, 0) '�õ��ļ��˵��ľ��
Hlocal = GetSubMenu(hMainMenu, 1)
hFtp = GetSubMenu(hMainMenu, 2)
hTxt = GetSubMenu(hMainMenu, 3)
hPic = GetSubMenu(hMainMenu, 4)
Hpla = GetSubMenu(hMainMenu, 5)
Hlst = GetSubMenu(hMainMenu, 6)
hFile = GetSubMenu(hMainMenu, 7)
HTY = GetSubMenu(hMainMenu, 8)
HNET = GetSubMenu(hMainMenu, 9)
HWQ = GetSubMenu(hMainMenu, 10)
HFIVE = GetSubMenu(hMainMenu, 11)
HLOOK = GetSubMenu(hMainMenu, 12)
hZor = GetSubMenu(hMainMenu, 13)
hFr = GetSubMenu(hMainMenu, 14)
hcol = GetSubMenu(hMainMenu, 15)
HABOUT = GetSubMenu(hMainMenu, 16)
'����˵������Ϣ
RegisterMenu hSubMenu, 0, "ȫѡ", 180, 20
RegisterMenu hSubMenu, 1, "����", 180, 20
RegisterMenu hSubMenu, 2, "����", 180, 20
RegisterMenu hSubMenu, 3, "ճ��", 180, 20
RegisterMenu hSubMenu, 4, "ɾ��", 180, 20

RegisterMenu Hlocal, 0, "Ԥ��ͼƬ", 180, 20
RegisterMenu Hlocal, 1, "", 180, 5
RegisterMenu Hlocal, 2, "��Ϊͷ��", 180, 20
RegisterMenu Hlocal, 3, "", 180, 5
RegisterMenu Hlocal, 4, "��תͼƬ", 180, 20
RegisterMenu Hlocal, 5, "�����˾�", 180, 20
RegisterMenu Hlocal, 6, "�����ع��", 180, 20
RegisterMenu Hlocal, 7, "�����ַ���", 180, 20
RegisterMenu Hlocal, 8, "Ϳѻ", 180, 20
RegisterMenu Hlocal, 9, "", 180, 5
RegisterMenu Hlocal, 10, "����ѷ���", 180, 20
RegisterMenu Hlocal, 11, "", 180, 5
RegisterMenu Hlocal, 12, "ͼƬ����", 180, 20

RegisterMenu hFtp, 0, "���ļ�", 180, 20
RegisterMenu hFtp, 1, "����ļ�", 180, 20
RegisterMenu hFtp, 2, "�����ļ���", 180, 20
RegisterMenu hFtp, 3, "����ļ���", 180, 20, PIC(5)
RegisterMenu hFtp, 4, "", 180, 5
RegisterMenu hFtp, 5, "�ҵ��ղ�", 180, 20
RegisterMenu hFtp, 6, "�����ִ�", 180, 20
RegisterMenu hFtp, 7, "", 180, 5
RegisterMenu hFtp, 8, "���벥���б�", 180, 20
RegisterMenu hFtp, 9, "���������б�", 180, 20
RegisterMenu hFtp, 10, "", 180, 5
RegisterMenu hFtp, 11, "��ͬ���б�", 180, 20, PIC(14)

RegisterMenu hTxt, 0, "����", 160, 20
RegisterMenu hTxt, 1, "���", 160, 20
RegisterMenu hTxt, 2, "����", 160, 20
RegisterMenu hTxt, 3, "", 160, 5
RegisterMenu hTxt, 4, "����ȫ��", 160, 20

RegisterMenu hPic, 0, "����", 160, 20
RegisterMenu hPic, 1, "���б�ɾ��", 160, 20
RegisterMenu hPic, 2, "�Ӵ���ɾ��", 160, 20
RegisterMenu hPic, 3, "", 160, 5
RegisterMenu hPic, 4, "ȥ���ظ�", 160, 20
RegisterMenu hPic, 5, "ɾ����Ч", 160, 20
RegisterMenu hPic, 6, "ˢ���б�", 160, 20
RegisterMenu hPic, 7, "", 160, 5
RegisterMenu hPic, 8, "������Ϣ", 160, 20
RegisterMenu hPic, 9, "��λ��", 160, 20
RegisterMenu hPic, 10, "", 160, 5
RegisterMenu hPic, 11, "�����ļ�Դ", 160, 20
RegisterMenu hPic, 12, "����������", 160, 20
RegisterMenu hPic, 13, "��������", 160, 20
RegisterMenu hPic, 14, "�ⲿ��", 160, 20
RegisterMenu hPic, 15, "", 160, 5
RegisterMenu hPic, 16, "�ļ�����", 160, 20

RegisterMenu Hpla, 0, "�����ظ�", 160, 20
RegisterMenu Hpla, 1, "˳�򲥷�", 160, 20
RegisterMenu Hpla, 2, "�������", 160, 20
RegisterMenu Hpla, 3, "ѭ������", 160, 20

RegisterMenu Hlst, 0, "����", 160, 20
RegisterMenu Hlst, 1, "�뿪", 160, 20
RegisterMenu Hlst, 2, "��Ҫ����", 160, 20
RegisterMenu Hlst, 3, "����", 160, 20
RegisterMenu Hlst, 4, "ע����½", 160, 20
RegisterMenu Hlst, 5, "", 160, 5
RegisterMenu Hlst, 6, "��������", 160, 20

RegisterMenu hFile, 0, "������Ϣ", 160, 20
RegisterMenu hFile, 1, "��������", 160, 20
RegisterMenu hFile, 2, "Զ��Э��", 160, 20
RegisterMenu hFile, 3, "��Ƶ����", 160, 20
RegisterMenu hFile, 4, "�����ļ�", 160, 20
RegisterMenu hFile, 5, "Ta����Ϣ", 160, 20
RegisterMenu hFile, 6, "", 160, 5
RegisterMenu hFile, 7, "ˢ���б�", 160, 20
RegisterMenu hFile, 8, "", 160, 5
RegisterMenu hFile, 9, "ɾ������", 160, 20
RegisterMenu hFile, 10, "�ٱ�Ta", 160, 20
RegisterMenu hFile, 11, "����Ta", 160, 20
RegisterMenu hFile, 12, "����������", 160, 20
RegisterMenu hFile, 13, "", 160, 5
RegisterMenu hFile, 14, "�����¼", 160, 20
RegisterMenu hFile, 15, "", 160, 5
RegisterMenu hFile, 16, "�޸�����", 160, 20
RegisterMenu hFile, 17, "�޸���Ϣ", 160, 20
RegisterMenu hFile, 18, "", 160, 5
RegisterMenu hFile, 19, "��������", 160, 20

RegisterMenu HTY, 0, "��ͼ��", 160, 20
RegisterMenu HTY, 1, "��ͼ��", 160, 20
RegisterMenu HTY, 2, "������", 160, 20
RegisterMenu HTY, 3, "�Գ���Ч", 160, 20
RegisterMenu HTY, 4, "�һ�ͼ��", 160, 20
RegisterMenu HTY, 5, "��תɫ��", 160, 20
RegisterMenu HTY, 6, "ħ��ɫ��", 160, 20
RegisterMenu HTY, 7, "Ѥ���ͻ�", 160, 20
RegisterMenu HTY, 8, "�ڰ׷���", 160, 20
RegisterMenu HTY, 9, "�ŵ両��", 160, 20
RegisterMenu HTY, 10, "������Ч��", 160, 20
RegisterMenu HTY, 11, "ѣ��Ч��(��)", 160, 20
RegisterMenu HTY, 12, "��������", 160, 20
RegisterMenu HTY, 13, "��������", 160, 20
RegisterMenu HTY, 14, "���ӶԱȶ�", 160, 20
RegisterMenu HTY, 15, "", 160, 5
RegisterMenu HTY, 16, "ˮƽ��ת", 160, 20
RegisterMenu HTY, 17, "��ֱ��ת", 160, 20
RegisterMenu HTY, 18, "˫��ת", 160, 20
RegisterMenu HTY, 19, "", 160, 5
RegisterMenu HTY, 20, "�Ӽ��а�ճ��", 160, 20
RegisterMenu HTY, 21, "���Ƶ����а�", 160, 20
RegisterMenu HTY, 22, "", 160, 5
RegisterMenu HTY, 23, "����ɫ�߿�", 160, 20
RegisterMenu HTY, 24, "����ɫ�ɰ�", 160, 20
RegisterMenu HTY, 25, "", 160, 5
RegisterMenu HTY, 26, "��ջ���", 160, 20
RegisterMenu HTY, 27, "", 160, 5
RegisterMenu HTY, 28, "��ͼ��", 160, 20
RegisterMenu HTY, 29, "����Ϳѻ", 160, 20
RegisterMenu HTY, 30, "����ѷ���", 160, 20
RegisterMenu HTY, 31, "", 160, 5
RegisterMenu HTY, 32, "��Ϊͷ��", 160, 20

RegisterMenu HNET, 0, "��Ϊ������", 160, 20
RegisterMenu HNET, 1, "", 160, 5
RegisterMenu HNET, 2, "����ɨ��", 160, 20
RegisterMenu HNET, 3, "��ս��", 160, 20

RegisterMenu HWQ, 0, "���������", 160, 20
RegisterMenu HWQ, 1, "", 160, 5
RegisterMenu HWQ, 2, "ֹͣ����", 160, 20
RegisterMenu HWQ, 3, "", 160, 5
RegisterMenu HWQ, 4, "���ļ�", 160, 20
RegisterMenu HWQ, 5, "��λ�ļ�", 160, 20
RegisterMenu HWQ, 6, "", 160, 5
RegisterMenu HWQ, 7, "������������", 160, 20
RegisterMenu HWQ, 8, "ɾ����������", 160, 20
RegisterMenu HWQ, 9, "������������", 160, 20
RegisterMenu HWQ, 10, "", 160, 5
RegisterMenu HWQ, 11, "���ɶ�ά��", 160, 20
RegisterMenu HWQ, 12, "", 160, 5
RegisterMenu HWQ, 13, "����б�", 160, 20

RegisterMenu HFIVE, 0, "��������", 160, 20
RegisterMenu HFIVE, 1, "ɾ�������ļ�", 160, 20
RegisterMenu HFIVE, 2, "", 160, 5
RegisterMenu HFIVE, 3, "��λ����", 160, 20
RegisterMenu HFIVE, 4, "��������", 160, 20
RegisterMenu HFIVE, 5, "", 160, 5
RegisterMenu HFIVE, 6, "ˢ���б�", 160, 20

RegisterMenu HLOOK, 0, "�鿴���", 160, 20
RegisterMenu HLOOK, 1, "ɾ����ʹ���", 160, 20
RegisterMenu HLOOK, 2, "�༭���", 160, 20
RegisterMenu HLOOK, 3, "�������", 160, 20
RegisterMenu HLOOK, 4, "", 160, 5
RegisterMenu HLOOK, 5, "������", 160, 20

RegisterMenu hZor, 0, "�Ͽ�����", 160, 20

RegisterMenu hFr, 0, "���ѡ�е���ǰ�б�", 160, 20
RegisterMenu hFr, 1, "���ȫ������ǰ�б�", 160, 20
RegisterMenu hFr, 2, "ɾ���ղ�", 160, 20
RegisterMenu hFr, 3, "����ղ�", 160, 20

RegisterMenu hcol, 0, "���ļ�", 160, 20
RegisterMenu hcol, 1, "", 160, 5
RegisterMenu hcol, 2, "ɾ���ļ�", 160, 20
RegisterMenu hcol, 3, "�����ļ�", 160, 20
RegisterMenu hcol, 4, "ճ���ļ�", 160, 20
RegisterMenu hcol, 5, "", 160, 5
RegisterMenu hcol, 6, "�������ļ�", 160, 20
RegisterMenu hcol, 7, "���������ļ����µ��ļ�", 160, 20
RegisterMenu hcol, 8, "", 160, 5
RegisterMenu hcol, 9, "������ļ�", 160, 20
RegisterMenu hcol, 10, "", 160, 5  '��ʼ�����ò˵����ı���ɫ,ע�������ɫ��ø�g_clrBkgNormalһ��,Ҫ��Ч������
RegisterMenu hcol, 11, "�ļ�����", 160, 20
RegisterMenu hcol, 12, "�ļ�������", 160, 20
RegisterMenu hcol, 13, "", 160, 5
RegisterMenu hcol, 14, "�½��ļ���", 160, 20

RegisterMenu HABOUT, 0, "������", 160, 20
RegisterMenu HABOUT, 1, "��Ļ����", 160, 20
RegisterMenu HABOUT, 2, "��������", 160, 20
RegisterMenu HABOUT, 3, "����ICEE", 160, 20
RegisterMenu HABOUT, 4, "", 160, 5
RegisterMenu HABOUT, 5, "�˳�", 160, 20
Call SetMenuBar(Me, &H8BA31F) 'RGB(224, 234, 240))
End Sub
Private Sub mnuBuddyRemove_Click()
frmma.�Ƴ�����
End Sub

Private Sub TMA_Timer()
With frmma
If .PICAD.Left >= 0 Then
.PICAD.Visible = True
TMA.Enabled = False
.PICAD.Left = 0
'Set objTimer = New clsWaitableTimer
frmma.TMAD.Enabled = True
'Set objTimer = Nothing
Else
.PICAD.Left = .PICAD.Left + 10
.PF(3).Left = .PICAD.Left + frmma.PICAD.Width
End If
End With
End Sub

Private Sub TMEA_Timer()
With frmma
If .PICAD.Left <= -340 Then
TMEA.Enabled = False
.PF(3).Left = 0
.PDB.Left = 0
.PICAD.Left = -340
.PICAD.Visible = False
Call .LOCKSAFE
If .pl.Left = 0 Then
If AUTOPLAYPIC = True Then .Timers.Enabled = True
.PF(6).Visible = True
End If
.Refresh

Dim PicBox As Control
For Each PicBox In frmma.Controls
If TypeOf PicBox Is PictureBox Then PicBox.Refresh
If TypeOf PicBox Is ICEE_WIN8 Then PicBox.Refresh
Next

Call GetVer
If Left(IEver, 1) <= "7" Then
If IETIP = 0 Then .PF(2).Visible = True: .PF(2).ZOrder 0
End If
If FIRSTRUN = True Then FIRSTRUN = False
Else
.PICAD.Left = .PICAD.Left - 10
.PF(3).Left = .PICAD.Left + .PICAD.Width
.PDB.Left = .PICAD.Left + .PICAD.Width
End If
End With
End Sub

Private Sub TMMU_Timer()
If MOUSEMO = True Then
ITIME = ITIME + 5
If ITIME >= 255 Then
TMMU.Enabled = False
ITIME = 0
MOUSEMO = False
frmma.PICMU.PICTURE = da2.PICTURE
frmma.PICMU.Visible = False
frmma.PNZ.Visible = True
End If
ShowTransparency da2, frmma.PICMU, ITIME, 0, 0
End If
End Sub
Private Sub ShowTransparency(cSrc As PictureBox, cDest As PictureBox, ByVal nLevel As Byte, x As Long, y As Long)
Dim LrProps As rBlendProps
Dim LnBlendPtr As Long
cDest.Cls
LrProps.tBlendAmount = nLevel
CopyMemory LnBlendPtr, LrProps, 4
With cSrc
AlphaBlend cDest.hdc, x, y, .ScaleWidth, .ScaleHeight, _
.hdc, 0, 0, .ScaleWidth, .ScaleHeight, LnBlendPtr
End With
cDest.Refresh
End Sub
Private Sub TMRCPU_Timer()
If LONELY_MODE = True Then Exit Sub

With frmma
'ITIM = ITIM + 1
If .PICAD.Visible = True Then
.PICAD.Cls
.PICAD.PaintPicture IMGAD.PICTURE, 0, 0, .PICAD.ScaleWidth, .PICAD.ScaleHeight
'Call PaintPng(App.Path & "\SKIN\LOADING" & ITIM & ".png", .PICAD.hdc, (.PICAD.ScaleWidth - 48) / 2, (.ScaleHeight - 48) / 2)
Call PaintPng(App.Path & "\SKIN\LOADING.png", .PICAD.hdc, (.PICAD.ScaleWidth - 111) / 2, .PICAD.ScaleHeight - 40)
'Call PaintPng(App.Path & "\SKIN\PT_S.png", .PICAD.hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\LOAD_FM.PNG", .PICAD.hdc, 0, 0)
Else
.CPU
.SHOWMEM
End If
'If ITIM > 6 Then ITIM = 0

End With
End Sub
Private Sub UNLOGIN_Click()
Call ע��
End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
If LONELY_MODE = True Then Exit Sub
If Left(IEver, 1) < 7 Then Exit Sub
WB.Silent = True
Dim I As Integer, s As String, SB As String
LSTLINK.Clear
s = ""
SB = ""
For I = 0 To WB.Document.links.Length - 1
If WB.Document.links.Item(I) <> s Then
SB = WB.Document.links.Item(I).innerText 'SB��ҳ�������г���������
s = WB.Document.links.Item(I) 'S��ҳ�������г�����
If Left(UCase(SB), 7) = "[TODAY]" Then LSTLINK.AddItem SB & "|" & s
End If
Next I
If LSTLINK.ListCount = 0 Then Exit Sub
LBLINK.Caption = Replace(Split(LSTLINK.List(0), "|")(0), "[TODAY]", "")
'С��ע:SPLIT(Ŀ�괮,���Ҵ�)(λ��) ����:split("ABCD|abcd")(1) ����ֵ����"abcd"
If LBLINK.Caption = "" Then frmma.lbthing.Caption = "��ӭʹ��1.24ȫ�°汾,���ྫ�ʵ��㷢��" Else frmma.lbthing.Caption = LBLINK.Caption
TimeHon.Enabled = True
End Sub

Private Sub Wb_DownloadBegin()
WB.Silent = True
End Sub
Private Sub timeHon_Timer()
If LONELY_MODE = True Then Exit Sub
If frmma.lbthing.Top + frmma.lbthing.Height < 0 Then frmma.lbthing.Top = frmma.PC.ScaleHeight
If frmma.lbthing.Top = 8 Then

TimeHon.Interval = 5000
frmma.lbthing.Top = 7
Else
TimeHon.Interval = 1
frmma.lbthing.Top = frmma.lbthing.Top - 1
End If
End Sub
Sub CHECKNET()
Dim TRasCon(255) As RASCONN95
Dim LG As Long
Dim LP As Long
Dim RetVal As Long
TRasCon(0).dwSize = 412
LG = 256 * TRasCon(0).dwSize
RetVal = RasEnumConnections(TRasCon(0), LG, LP)
status.dwSize = 160
RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, status)

End Sub

Private Sub wsListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call SHOWWRONG("��������:" & Number & vbCrLf & Description, 2)
End Sub


Private Sub �༭���_Click()
FRMLRC.L_EDIT
End Sub

Private Sub ���ŵ�ǰ�б�_Click()
With FrmNetMusic
Dim MName As String, aname As String
Dim I As Integer
If .B_LIST.ListCount = 0 Then Exit Sub
For I = 0 To .B_LIST.ListCount - 1
MusicName = .B_LIST.List(I)
aname = Trim(Left$(MusicName, InStr(1, MusicName, "-") - 1))
MName = Trim(Right$(MusicName, Len(MusicName) - InStr(1, MusicName, "-")))
frmma.PLIST.AddItem MName, aname, FindMp3URL(MName, aname), 0
Next
Call frmma.SAVELIST
End With
End Sub

Private Sub ������������_Click()
On Error Resume Next
frmma.PLIST.AddItem FrmNetMusic.M_N, FrmNetMusic.A_N, FrmNetMusic.Will_DL, 0
frmma.Wm.URL = FrmNetMusic.Will_DL
frmma.PLIST.ListIndex = PLIST.ListCount - 1
End Sub

Private Sub �鿴���_Click()
FRMLRC.L_VIEW
End Sub

Private Sub �Ӽ��а�ճ��_Click()
On Error Resume Next
PICCLIP.PICTURE = Clipboard.GetData(2)
Call SavePicture(PICCLIP.PICTURE, App.Path & "\thumbs\TH_CLIP.BMP")
Call FRMBOARD.OpenFile(App.Path & "\thumbs\TH_CLIP.BMP")
End Sub

Private Sub ��URL_Click()
Call frmma.MUSICBOX
End Sub

Private Sub ������_Click()
On Error Resume Next
If Right(FRMDOWN.LVIEW.SelectedItem.SubItems(2), 1) = "%" Then Exit Sub
If FRMDOWN.LVIEW.SelectedItem.SubItems(4) = "" Then Exit Sub
Call SYSTEMOPEN(Dpath & FRMDOWN.LVIEW.SelectedItem.Text)
End Sub

Private Sub ���ļ�_Click()
FRMEX.OPEN_CLICK
End Sub

Private Sub ��λ����_Click()
On Error Resume Next
If FRMDOWN.LVIEW.SelectedItem.SubItems(4) = "" Then Exit Sub
Shell "explorer.exe /select," & Dpath & FRMDOWN.LVIEW.SelectedItem.Text, vbNormalFocus
End Sub

Private Sub �Ͽ�TCP����_Click()
FRMEND.KIILTCP
End Sub

Private Sub ��������_Click()
ShellExecute 0&, vbNullString, "http://tieba.baidu.com/f?ie=utf-8&kw=icee", vbNullString, vbNullString, 0 '����ie
End Sub

Private Sub ������ļ�_Click()
frmma.SHAREIT (FRMEX.Txt_Address.Text & "\" & FRMEX.ListView1.SelectedItem.Text)
End Sub

Private Sub ��������_Click()
DefCOM = 1
Call frmma.SHAREIT(frmma.PLIST.URL(frmma.PLIST.ListIndex))
End Sub

Private Sub ����_Click()
On Error Resume Next
    Dim BD As BmpFile, BS As BmpFile, filename As String
    filename = App.Path & "\THUMBS\THUMBS.Bmp"
    Call SavePicture(FRMBOARD.PICTY.image, filename)
    Call GetBmpFile(filename, BS)
    FuDiao BS, BD
    PutBmpFile App.Path & "\THUMBS\THUMBS.Bmp", BD
    Call FRMBOARD.OpenFile(App.Path & "\THUMBS\THUMBS.Bmp")
    fso.DeleteFile App.Path & "\THUMBS\THUMBS.Bmp"
    fso.DeleteFile App.Path & "\THUMBS\THUMB.Bmp"
End Sub

Private Sub �����ļ�_Click()
FRMEX.Copy_Click
End Sub

Private Sub ������������_Click()
On Error Resume Next
Clipboard.SetText (FRMDOWN.LVIEW.SelectedItem.SubItems(7))
End Sub

Private Sub ����_Click()
FrmWhatNew.Show
End Sub

Private Sub ��������_Click()
On Error GoTo ERR
    GetObject FRMBOARD.PICTY.image, Len(PICINFO), PICINFO
    ReDim PicBits(1 To PICINFO.bmWidth * PICINFO.bmHeight * 3) As Byte
    GetBitmapBits FRMBOARD.PICTY.PICTURE, UBound(PicBits), PicBits(1)
    For Cnt = 1 To UBound(PicBits)
        PicBits(Cnt) = PicBits(Cnt) * 0.618
    Next Cnt
    SetBitmapBits FRMBOARD.PICTY.PICTURE, UBound(PicBits), PicBits(1)
    FRMBOARD.PICTY.Refresh
ERR:
Exit Sub
End Sub

Private Sub ������_Click()
FRMUP.Show
End Sub

Private Sub ��������_Click()
    Call FRMEND.ENDIT
End Sub

Private Sub �����ļ�_Click()
FRMEND.FOLDERPRO
End Sub

Private Sub ��������_Click()
FRMEND.PAPERPRO
End Sub

Private Sub ħ��_Click()
    Dim BD As BmpFile, BS As BmpFile, filename As String
    On Error Resume Next
    filename = App.Path & "\THUMBS\THUMBS.Bmp"
    Call SavePicture(FRMBOARD.PICTY.image, filename)
    Call GetBmpFile(filename, BS)
    Call MoShu(BS, BD)
    Call PutBmpFile(App.Path & "\THUMBS\THUMBS.Bmp", BD)
    Call FRMBOARD.OpenFile(App.Path & "\THUMBS\THUMBS.Bmp")
    fso.DeleteFile App.Path & "\THUMBS\THUMBS.Bmp"
    fso.DeleteFile App.Path & "\THUMBS\THUMB.Bmp"
End Sub

Private Sub ľ��_Click()
On Error Resume Next
    Dim BD As BmpFile, BS As BmpFile, filename As String
    filename = App.Path & "\THUMBS\THUMBS.Bmp"
    Call SavePicture(FRMBOARD.PICTY.image, filename)
    Call GetBmpFile(filename, BS)
    MuKe BS, BD
    PutBmpFile App.Path & "\THUMBS\THUMBS.Bmp", BD
    Call FRMBOARD.OpenFile(App.Path & "\THUMBS\THUMBS.Bmp")
    fso.DeleteFile App.Path & "\THUMBS\THUMBS.Bmp"
    fso.DeleteFile App.Path & "\THUMBS\THUMB.Bmp"
End Sub
Private Sub ����_Click()
Call frmma.����һ��(FRMBOARD.PICTY)
End Sub
Private Sub ����ѡ��_Click()
frmma.���Ÿ���
End Sub

Private Sub ����_Click()
If frmma.lstRes.List(frmma.lstRes.ListIndex) <> "" Then
frmma.Text3.Text = frmma.lstRes.Text
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">���Խ�" & frmma.lstRes.List(frmma.lstRes.ListIndex) & "��Ϊ������"
frmma.PICNET.Visible = False
Call frmma.LOCKSAFE
Call frmma.SUBDRAWIM
frmma.LBITEM(2).Caption = "���½"
frmma.IMJ.Visible = False
frmma.IMG_NT.Visible = True
End If
End Sub
Private Sub ��ֱ_Click()
Call FlipImage(FRMBOARD.PICTY, 1)
End Sub

Private Sub �򿪲����б�_Click()
Dim sFile As String
sFile = ShowOpen(Me.hwnd, "�����б��ļ� M3u" & Chr(0) & "*.m3u", "�򿪲����б�")
If Dir$(sFile) <> vbNullString And sFile <> "" Then Call frmma.Playlist(sFile)
End Sub
Sub �򿪹���()
'�����Ǵ�CD -ROM�Ĺ��̴���:
retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub

Private Sub ��ý��_Click()
Lmenu (0)
End Sub

Private Sub ��ͼ��_Click()
Dim sFile As String
sFile = ShowOpen(Me.hwnd, "BMP�ļ�" & Chr(0) & "*.Bmp" _
& Chr(0) & "JEPG�ļ�" & Chr(0) & "*.jpg;*.jepg" _
& Chr(0) & "Gif" & Chr(0) & "*.gif" _
& Chr(0) & "Png" & Chr(0) & "*.png", "��ͼƬ")
Call FRMBOARD.OpenFile(sFile)
End Sub

Private Sub ���ļ���_Click()
Call OpenDir
End Sub
Private Sub ����ѭ��_Click()
LOLIPOP = 1
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "����ѭ��"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DQ_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">�������ֲ���ģʽΪ����ѭ��"
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
End Sub

Private Sub ����_Click()
frmma.�����б�
End Sub
Private Sub ����_Click()
Call Report
End Sub
Private Sub �ָ��ļ�_Click()
On Error Resume Next
Dim r As Long
With frmma
Dim filename As String
filename = .PLIST.URL(.PLIST.ListIndex)
r = ShowProperties(filename, frmma.hwnd)
End With
If r <= 32 Then Call SHOWWRONG("�Բ���,�鿴�ļ�����ʧ��(���ܵ�ԭ����Ȩ�޲���,���߸�����������)", 2)

End Sub

Private Sub ����·��_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText (frmma.PLIST.URL(frmma.PLIST.ListIndex))
End Sub
Private Sub �����ı�_Click()
Call ����
End Sub

Sub �رչ���()
'�ر�CD -ROM�����´���:
retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub
Private Sub ����վ_Click()
Call frmma.����ɾ������
End Sub
Private Sub �����ı�_Click()
Call ����
End Sub
Private Sub �����¼_Click()
If frmma.Left > FRMHIS.Width Then
FRMHIS.Move frmma.Left - FRMHIS.Width, frmma.Top
Else
FRMHIS.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMHIS.Show
End Sub
Private Sub Ĭ�ϳ����_Click()
On Error Resume Next
Call SYSTEMOPEN(frmma.PLIST.URL(frmma.PLIST.ListIndex))
End Sub

Private Sub ����������_Click()
If frmma.Left > FORMNAME.Width Then
FORMNAME.Move frmma.Left - FORMNAME.Width, frmma.Top
Else
FORMNAME.Move frmma.Left + frmma.Width, frmma.Top
End If
FORMNAME.Show
FORMNAME.txtPath.Text = MMAIN.GetPathFromFileName(frmma.PLIST.URL(frmma.PLIST.ListIndex), "\")
End Sub

Private Sub �������������ļ����µ��ļ�_Click()
If frmma.Left > FORMNAME.Width Then
FORMNAME.Move frmma.Left - FORMNAME.Width, frmma.Top
Else
FORMNAME.Move frmma.Left + frmma.Width, frmma.Top
End If

FORMNAME.Show
FORMNAME.txtPath = FRMEX.Txt_Address.Text
End Sub

Private Sub ����_Click()
frmma.���δ��û�
End Sub

Private Sub ��Ļ����_Click()
FRMKEYBOARD.Show
End Sub

Private Sub ����ղ�_Click()
Call FRMFAV.CLEAR_FAV
End Sub

Private Sub �������_Click()
frmma.lstRes.Clear
End Sub

Private Sub �������_Click()
FRMDOWN.LVIEW.ListItems.Clear
If FRMDOWN.LVIEW.ListItems.Count > 0 Then FRMDOWN.LVIEW.Visible = True Else FRMDOWN.LVIEW.Visible = False
Call FRMDOWN.SAVELIST
Call FRMDOWN.LoadList
End Sub

Private Sub ȥ��ͼ��_Click()
Set FRMBOARD.PT.PICTURE = Nothing
Set FRMBOARD.PICTY.PICTURE = Nothing
FRMBOARD.PT.BackColor = FRMBOARD.PB.BackColor  'Ϳѻ������ɫ
FRMBOARD.PICTY.BackColor = FRMBOARD.PB.BackColor
End Sub

Private Sub ȥ��_Click()
Call frmma.ȥ���ظ�
End Sub

Private Sub ȫѡ�ı�_Click()
Call ȫѡ
End Sub

Private Sub ɾ������_Click()
FRMLRC.L_DELETE
Call FrmNetMusic.L_LRC.ClearLrc
FrmNetMusic.L_LRC.Visible = False
If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText " ICEE����,������������"
End Sub

Private Sub ɾ�������ļ�_Click()
FRMEND.DELPRO
End Sub

Private Sub ɾ���ղ�_Click()
On Error Resume Next
Call FRMFAV.REMOVE_ITEM(FRMFAV.LFAV.SelectedItem.Text)
End Sub

Private Sub ɾ���ı�_Click()
Call ɾ������
End Sub

Private Sub ɾ���ļ�_Click()
FRMEX.Del_Click
End Sub

Private Sub ɾ����Ч_Click()
Call DEL_NONE
End Sub

Private Sub ɾ����������_Click()
If FRMDOWN.LVIEW.ListItems.Count = 0 Then Exit Sub
FRMDOWN.LVIEW.ListItems.REMOVE (FRMDOWN.LVIEW.SelectedItem.Index)
If FRMDOWN.LVIEW.ListItems.Count > 0 Then FRMDOWN.LVIEW.Visible = True Else FRMDOWN.LVIEW.Visible = False
Call FRMDOWN.SAVELIST
Call FRMDOWN.LoadList
End Sub

Private Sub ɾ��ѡ��_Click()
Call Lmenu(2)
End Sub

Private Sub ��Ϊͷ��_Click()
Dim filea As String
filea = App.Path & "\THUMBS\H_Thumbs.Bmp"
Call SavePicture(FRMBOARD.PICTY.image, filea)
FRMHEAD.Show
Call FRMHEAD.OpenFile(filea)

Kill (filea)
End Sub

Private Sub ���ɶ�ά��_Click()
FRMDOWN.��ά��
End Sub

Private Sub �ղؼ�_Click()
If frmma.Left > FRMFAV.Width Then
FRMFAV.Move frmma.Left - FRMFAV.Width, frmma.Top
Else
FRMFAV.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMFAV.Show
End Sub

Private Sub ˢ�º����б�_Click()
Dim I As Integer
With frmma
For I = 1 To .TreeView1.Nodes.Count
.Winsock1.SendData ".getstatus " & .TreeView1.Nodes(I).Key
Next
End With
End Sub

Private Sub ˢ�½���_Click()
Call FRMEND.ListProcess
End Sub

Private Sub ˢ���б�_Click()
frmma.PLIST.Refresh
End Sub
Private Sub ˫��_Click()
Call FlipImage(FRMBOARD.PICTY, 2)
End Sub

Private Sub ˮƽ_Click()
Call FlipImage(FRMBOARD.PICTY, 0)
End Sub

Private Sub ˳�򲥷�_Click()
LOLIPOP = 3
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "˳�򲥷�"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SX_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">�������ֲ���ģʽΪ˳Ѱ����"
End Sub

Private Sub �������_Click()
FRMLRC.MOVEME
FRMLRC.Show
End Sub

Private Sub �������_Click()
LOLIPOP = 0
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "�������"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SJ_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">�������ֲ���ģʽΪ�������"
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
End Sub

Private Sub ����_Click()
Call LOCKME
End Sub
Sub LOCKME()
On Error Resume Next
With frmma
.RUNSAFE
.TXTPOUP.SetFocus
If frmma.Winsock1.State = 7 Then MYSTATUS = 2
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">ICEE�������ɹ�"
.DRAWLOCK
IS_LOCK = True
Frmm.Hide
End With
End Sub
Private Sub ��ӵ���_Click()
Lmenu (4)
End Sub

Private Sub ���ȫ���ղ�_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To FRMFAV.LFAV.ListItems.Count
If FRMFAV.LFAV.ListItems(I).Text <> "" Then frmma.PLIST.AddItem (FRMFAV.LFAV.ListItems(I).Text), "", FRMFAV.LFAV.ListItems(I).SubItems(I)
Next
End Sub

Private Sub ����ղ�_Click()
On Error Resume Next
frmma.PLIST.AddItem FRMFAV.LFAV.SelectedItem.Text, "", FRMFAV.LFAV.SelectedItem.SubItems(1)
End Sub

Private Sub ����ļ���_Click()
Call AddDir
End Sub

Private Sub ���������_Click()
Frmadd.Show
End Sub
Private Sub ������б�_Click()
frmma.PLIST.AddItem FrmNetMusic.M_N, FrmNetMusic.A_N, FrmNetMusic.Will_DL, 0
End Sub

Private Sub ֹͣ����_Click()
Call FRMDOWN.ֹͣ����
End Sub

Private Sub Ϳѻ���а�_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetData FRMBOARD.PICTY.image
End Sub

Private Sub �˴�_Click()
Unload frmma
End Sub

Private Sub λ��_Click()
On Error Resume Next
If Dir(frmma.PLIST.URL(frmma.PLIST.ListIndex)) = "" Then Exit Sub
Shell "explorer.exe /select," & frmma.PLIST.URL(frmma.PLIST.ListIndex), vbNormalFocus
End Sub

Private Sub �ļ�������_Click()
FRMEX.Attribute_Click
End Sub

Private Sub �ļ�����_Click()
FRMEX.PropsShow (FRMEX.Txt_Address.Text & "\" & FRMEX.ListView1.SelectedItem.Text)
End Sub

Private Sub ����_Click()
Call DoFileDownload(StrConv(FrmNetMusic.Will_DL, vbUnicode))
End Sub

Private Sub �½��ļ���_Click()
FRMEX.NewFolder_Click
End Sub

Private Sub ѣ��_Click()
Call ѣ��ͼ��(FRMBOARD.PICTY)
Call SavePicture(FRMBOARD.PICTY.image, App.Path & "\THUMBS\THUMBS.BMP")
FRMBOARD.PT.PICTURE = LoadPicture(App.Path & "\THUMBS\THUMBS.BMP")
Call ѣ��ͼ��(FRMBOARD.PT)
Call SavePicture(FRMBOARD.PT.image, App.Path & "\THUMBS\THUMBS.BMP")
FRMBOARD.PT.PICTURE = LoadPicture(App.Path & "\THUMBS\THUMBS.BMP")
FRMBOARD.PICTY.PaintPicture FRMBOARD.PT.image, 0, 0, FRMBOARD.PT.ScaleWidth, FRMBOARD.PT.ScaleHeight
Set FRMBOARD.PT.PICTURE = Nothing
Call PictureBoxSaveJPG(FRMBOARD.PICTY.image, App.Path & "\MEDIA\Paint\AutoSave.JPG", 100)

End Sub

Private Sub ѭ��_Click()
LOLIPOP = 2
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "�б�ѭ��"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\XH_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">�������ֲ���ģʽΪ�б�ѭ��"
End Sub

Private Sub ��������_Click()
On Error Resume Next
Call FRMMIN.SeeIt(frmma.PLIST.URL(frmma.PLIST.ListIndex))
If frmma.Left > FRMMIN.Width Then
FRMMIN.Move frmma.Left - FRMMIN.Width, frmma.Top
Else
FRMMIN.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMMIN.Show
End Sub

Private Sub �ͻ�_Click()
On Error Resume Next
    Dim BD As BmpFile, BS As BmpFile, filename As String
    filename = App.Path & "\THUMBS\THUMBS.Bmp"
    Call SavePicture(FRMBOARD.PICTY.image, filename)
    Call GetBmpFile(filename, BS)
    Call YouHua(BS, BD, 5)
    Call PutBmpFile(App.Path & "\THUMBS\THUMBS.Bmp", BD)
    FRMBOARD.OpenFile (App.Path & "\THUMBS\THUMBS.Bmp")
    fso.DeleteFile App.Path & "\THUMBS\THUMBS.Bmp"
    fso.DeleteFile App.Path & "\THUMBS\THUMB.Bmp"
End Sub

Private Sub ����ѷ���_Click()
Dim Tfile As String
Tfile = App.Path & "\THUMBS\THUMBS.Bmp"
DefCOM = 0
Call SavePicture(FRMBOARD.PICTY.image, Tfile)
Call frmma.SHAREIT(Tfile)
End Sub
Private Sub ��ͬ��_Click()
Dim JSI As Integer
JSI = GetSetting("ICEE", "Winsock", "Connect", 0)
If JSI = 0 Then
Call SHOWWRONG("���ȵ�¼������!", 2)
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">���������б���ͬ��ʧ��(δ��¼������)"
Else
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">���������б���ͬ���ɹ�"
End If
End Sub

Private Sub ���ӶԱȶ�_Click()
Call UPDB(FRMBOARD.PICTY)
End Sub

Private Sub ��������_Click()
Call UPLD(FRMBOARD.PICTY)
End Sub

Private Sub ճ���ı�_Click()
Call ճ��
End Sub
Sub OpenDir()
Dim sDir As String, a As Integer
With frmma
sDir = BrowseFolder("�������ļ���", frmma)
If sDir = "" Then Exit Sub
File1.Path = sDir
If File1.ListCount > 0 Then
.PLIST.Clear
For a = 0 To File1.ListCount - 1
.PLIST.AddItem LastFileName(File1.Path & "\" & File1.List(a)), "", File1.Path & "\" & File1.List(a), 0
Next
Call .SAVELIST
.Wm.URL = .PLIST.URL(0)
.Wm.Controls.Play
End If
End With
End Sub
Sub AddDir()
Dim sDir As String, a As Integer
sDir = BrowseFolder("����ļ���", frmma)
If sDir = "" Then Exit Sub
With frmma
File1.Path = sDir
If File1.ListCount > 0 Then
For a = 0 To File1.ListCount - 1
.PLIST.AddItem LastFileName(File1.Path & "\" & File1.List(a)), "", File1.Path & "\" & File1.List(a), 0
Next
Call .SAVELIST
End If
End With
End Sub
Private Sub mnuBuddyChat_Click()
frmma.��������
End Sub

Private Sub mnuBuddyIgnore_Click()
Call I_IGNORE
End Sub
Sub I_IGNORE()
On Error Resume Next
With frmma
If .SETME.Enabled = False Then Exit Sub
If .PICIM.Left <> 0 Then Call .ShowIM
.PP.Visible = False
.Winsock1.SendData ".GetIgnoreList"
For I = 1 To .TreeView1.Nodes.Count
.TXTBOX.AddItem .TreeView1.Nodes(I)
Next
.PICIM.BackColor = PTCO.POINT(0, 0)
.DRAWUN
.TXTSER.Text = "������Է�ID"
.PICIG.Visible = True
.IMJ.Visible = True
.RUNSAFE
.LA(1).Caption = "����������"
.LSTBOX.Selected(0) = True
End With
End Sub
Private Sub mnuBuddyInfo_Click()
frmma.PICIM.BackColor = PTCO.POINT(0, 0)
frmma.������Ϣ
End Sub

Private Sub mnuBuddyMessage_Click()
frmma.��ʱ����
End Sub
Private Sub mnuFilePassword_Click()
Call CHANGEPASS
End Sub
Sub CHANGEPASS()
On Error Resume Next
With frmma
If .SETME.Enabled = False Then Exit Sub
.DRAWPASS
If .PICIM.Left <> 0 Then Call .ShowIM
.PP.Visible = False
.PICIM.BackColor = PTCO.POINT(0, 0)
.PICPASS.Visible = True
.RUNSAFE
.LA(1).Caption = "�޸�����"
.TXTPASS.Text = GetInitEntry("IM", "LastPassWord", "")
.IMJ.Visible = True
.TXTPASS.SetFocus
.TXTPASS.SelStart = Len(.TXTPASS.Text)
End With
End Sub
Sub Report()
On Error Resume Next
With frmma
If .SETME.Enabled = False Then Exit Sub
If .PICIM.Left <> 0 Then Call .ShowIM
.PP.Visible = False
.DRAWBUG
.PICBUG.Visible = True
.IMJ.Visible = True
.LA(1).Caption = "BUG�ύ"
.TXTSER.Text = "������Է�ID"
.RUNSAFE
.PICIM.BackColor = PTCO.POINT(0, 0)
.TXTBUG.SetFocus
End With
End Sub

Sub ע��()
Call frmma.��ʼ��
SetTrayIcon Frmm.OFFLINE.PICTURE

End Sub

Private Sub mnuFileChangeInfo_Click()
Call FRMSETINFO.Show
End Sub

Private Sub mnuStatusAway_Click()

mnuStatusOnline.Checked = False
mnuStatusAway.Checked = True
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = False
frmma.Winsock1.SendData ".status Away"
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "Ŀǰ�����뿪״̬"
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">����״̬Ϊ�뿪"
SetTrayIcon Away.PICTURE
MYSTATUS = 1  '1Ϊ�뿪
frmma.DRAWFACE
End Sub

Private Sub mnuStatusDND_Click()
mnuStatusOnline.Checked = False
mnuStatusAway.Checked = False
mnuStatusDND.Checked = True
mnuStatusInvisible.Checked = False
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "Ŀǰ����æµ״̬"
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">����״̬Ϊæµ"
frmma.Winsock1.SendData ".status DND"
SetTrayIcon BusyNow.PICTURE
MYSTATUS = 2  '2Ϊæµ
frmma.DRAWFACE
End Sub


Private Sub mnuStatusInvisible_Click()
mnuStatusOnline.Checked = False
mnuStatusAway.Checked = False
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = True
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "Ŀǰ��������״̬"
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">����״̬Ϊ����"
frmma.Winsock1.SendData ".status Invisible"
SetTrayIcon HideNow.PICTURE
MYSTATUS = 3  '3Ϊ����
frmma.DRAWFACE
End Sub


Private Sub mnuStatusOnline_Click()
mnuStatusOnline.Checked = True
mnuStatusAway.Checked = False
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = False
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "Ŀǰ��������״̬"
frmma.Winsock1.SendData ".status ONLINE"
SetTrayIcon ONLINE.PICTURE
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">����״̬Ϊ����"
MYSTATUS = 0 '0Ϊ����
frmma.DRAWFACE
End Sub

Private Sub ճ���ļ�_Click()
FRMEX.Plaster_Click
End Sub

Private Sub �������ļ�_Click()
FRMEX.ReName_Click
End Sub

Private Sub ��������_Click()
Call frmma.SERCHNET
End Sub
Private Sub ��_Click()
Call Sharpen(FRMBOARD.PICTY, 1)
End Sub
Private Sub ģ��_Click()
Call BlurImage(FRMBOARD.PICTY)
End Sub
Private Sub ����_Click()
Call Noise(FRMBOARD.PICTY, 20)
End Sub
Private Sub ����_Click()
Call Mirror(FRMBOARD.PICTY)
End Sub
Private Sub �Ҷ�_Click()
Call GrayImage(FRMBOARD.PICTY)
End Sub
Private Sub ��ת_Click()
Call InvertImage(FRMBOARD.PICTY)
End Sub
Private Sub ������_Click()
Call MASAK(FRMBOARD.PICTY)
End Sub
Private Sub ����ɫ�ɰ�_Click()
On Error Resume Next
FRMBOARD.PO(0).ScaleMode = 1
Set FRMBOARD.PICTY.PICTURE = FRMBOARD.PICTY.image
Call ShadePicture(FRMBOARD.PICTY, FRMBOARD.PICTY, FRMBOARD.PB.BackColor, 5)
FRMBOARD.PO(0).ScaleMode = 3
End Sub

Private Sub ��ɫ�߿�_Click()
Call StrokeImage(FRMBOARD.PICTY, 15, FRMBOARD.PB.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 13, FRMBOARD.PF.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 10, FRMBOARD.PB.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 8, FRMBOARD.PF.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 5, FRMBOARD.PB.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 1, FRMBOARD.PF.BackColor)
End Sub

Private Sub ������_Click()
On Error Resume Next
If D_L_SHOW = True Then Exit Sub
If IS_NET = True Then Call FrmNetMusic.SETLRC
End Sub
Sub DEL_NONE()
Dim I As Integer
For I = 0 To frmma.PLIST.ListCount - 1
If UCase(Left(frmma.PLIST.URL(I), 4)) = "HTTP" Then
IS_CHK_LIST = True
Call CHECKNET
If status.RasConnState <> &H2000 Then Exit Sub
If MMAIN.FindMp3URL(frmma.PLIST.Title(I), frmma.PLIST.AUTHOR(I)) = "" Then frmma.PLIST.RemoveItem (I)
Else
If PathFileExists(frmma.PLIST.URL(I)) = 0 Then frmma.PLIST.RemoveItem (I)
End If
Next
IS_CHK_LIST = False
End Sub
Sub DRAW_LOGO()
PIC(45).Cls
PIC(46).Cls
PIC(47).Cls
PIC(45).PICTURE = Nothing
PIC(46).PICTURE = Nothing
PIC(47).PICTURE = Nothing
PIC(131).Cls
PIC(130).Cls
R_P_THU = GetInitEntry("SYSTEM", "REPLACE", 0)
PIC(45).BackColor = frmma.iFrame.BackColor
PIC(46).BackColor = frmma.iFrame.BackColor
PIC(47).BackColor = frmma.iFrame.BackColor
Dim LES
LES = BitBlt(PIC(45).hdc, 0, 0, frmma.PICMU.Width, frmma.PICMU.Height, frmma.iFrame.hdc, frmma.PICMU.Left, frmma.PICMU.Top, &HCC0020)
LES = BitBlt(PIC(46).hdc, 0, 0, frmma.PICMU.Width, frmma.PICMU.Height, frmma.iFrame.hdc, frmma.PICMU.Left, frmma.PICMU.Top, &HCC0020)
LES = BitBlt(PIC(47).hdc, 0, 0, frmma.PICMU.Width, frmma.PICMU.Height, frmma.iFrame.hdc, frmma.PICMU.Left, frmma.PICMU.Top, &HCC0020)
frmma.PC.Move 0, 576, 360
frmma.lbthing.Move 50
frmma.PC.ZOrder 1
Call PaintPng(App.Path & "\SKIN\LG_N.PNG", PIC(45).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\LG_H.PNG", PIC(46).hdc, 0, 0)
Call PaintPng(App.Path & "\SKIN\LG_N.PNG", PIC(47).hdc, 0, 0)
PIC(45).Refresh
PIC(46).Refresh
PIC(47).Refresh
Call LoadStyle
End Sub
Sub LoadStyle()
On Error Resume Next
USEBACK = GetInitEntry("SYSTEM", "BACKPICTURE", App.Path + "\SKIN\BK\0.JPG")
IMBK.PICTURE = LoadPicture(USEBACK)
SKINDRAW.PaintPicture IMBK.PICTURE, 0, 0, SKINDRAW.ScaleWidth, SKINDRAW.ScaleHeight
IMGAD.PICTURE = SKINDRAW.image
da1.PICTURE = PIC(45).image
da2.PICTURE = PIC(46).image
da3.PICTURE = PIC(47).image
End Sub
