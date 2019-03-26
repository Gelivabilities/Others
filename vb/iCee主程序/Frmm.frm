VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Frmm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFBC44&
   Caption         =   "菜单"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   1080
   ClientWidth     =   15690
   Icon            =   "Frmm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1046
   StartUpPosition =   2  '屏幕中心
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
      ToolTipText     =   "主菜单"
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
      ToolTipText     =   "主菜单"
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
      ToolTipText     =   "主菜单"
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
      Caption         =   "再见,曾经复杂的我,新版本,使我更精练"
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
   Begin VB.Menu 文本 
      Caption         =   "文本"
      Begin VB.Menu 全选文本 
         Caption         =   "全选"
      End
      Begin VB.Menu 复制文本 
         Caption         =   "复制"
      End
      Begin VB.Menu 剪切文本 
         Caption         =   "剪切"
      End
      Begin VB.Menu 粘贴文本 
         Caption         =   "粘贴"
      End
      Begin VB.Menu 删除文本 
         Caption         =   "删除"
      End
   End
   Begin VB.Menu 图像处理 
      Caption         =   "图片处理"
      Begin VB.Menu oxox 
         Caption         =   "预览图片"
         Index           =   0
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu oxox 
         Caption         =   "裁切作为头像"
         Index           =   2
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu oxox 
         Caption         =   "旋转图片"
         Index           =   4
      End
      Begin VB.Menu oxox 
         Caption         =   "调整曝光"
         Index           =   5
      End
      Begin VB.Menu oxox 
         Caption         =   "使用滤镜"
         Index           =   6
      End
      Begin VB.Menu oxox 
         Caption         =   "生成字符画"
         Index           =   7
      End
      Begin VB.Menu oxox 
         Caption         =   "涂鸦"
         Index           =   8
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu oxox 
         Caption         =   "分享"
         Index           =   10
      End
      Begin VB.Menu oxox 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu oxox 
         Caption         =   "图像详情"
         Index           =   12
      End
   End
   Begin VB.Menu 播放控制 
      Caption         =   "播放"
      Begin VB.Menu 打开媒体 
         Caption         =   "打开文件"
      End
      Begin VB.Menu 添加单个 
         Caption         =   "添加文件"
      End
      Begin VB.Menu 打开文件夹 
         Caption         =   "打开文件夹"
      End
      Begin VB.Menu 添加文件夹 
         Caption         =   "添加文件夹"
      End
      Begin VB.Menu swff01122 
         Caption         =   "-"
      End
      Begin VB.Menu 收藏夹 
         Caption         =   "收藏夹"
      End
      Begin VB.Menu 打开URL 
         Caption         =   "音乐窗"
      End
      Begin VB.Menu wqE 
         Caption         =   "-"
      End
      Begin VB.Menu 打开播放列表 
         Caption         =   "导入播放列表"
      End
      Begin VB.Menu 导出 
         Caption         =   "导出播放列表"
      End
      Begin VB.Menu SHDH 
         Caption         =   "-"
      End
      Begin VB.Menu 云同步 
         Caption         =   "云同步列表"
      End
   End
   Begin VB.Menu 招商 
      Caption         =   "网络音乐"
      Begin VB.Menu 播放网络音乐 
         Caption         =   "播放"
      End
      Begin VB.Menu 添加至列表 
         Caption         =   "添加到播放列表"
      End
      Begin VB.Menu 下载 
         Caption         =   "下载选定歌曲"
      End
      Begin VB.Menu SFEGH 
         Caption         =   "-"
      End
      Begin VB.Menu 播放当前列表 
         Caption         =   "播放当前列表"
      End
   End
   Begin VB.Menu 文件 
      Caption         =   "文件"
      Begin VB.Menu 播放选中 
         Caption         =   "播放选中的歌曲"
      End
      Begin VB.Menu 删除选中 
         Caption         =   "删除选中的歌曲"
      End
      Begin VB.Menu 回收站 
         Caption         =   "移至回收站"
      End
      Begin VB.Menu jdhj 
         Caption         =   "-"
      End
      Begin VB.Menu 去重 
         Caption         =   "去除重复"
      End
      Begin VB.Menu 删除无效 
         Caption         =   "删除无效的任务"
      End
      Begin VB.Menu 刷新列表 
         Caption         =   "刷新列表"
      End
      Begin VB.Menu ajkhf 
         Caption         =   "-"
      End
      Begin VB.Menu 音乐属性 
         Caption         =   "查看选中文件属性"
      End
      Begin VB.Menu 位置 
         Caption         =   "打开文件所在位置"
      End
      Begin VB.Menu SFFGGHH 
         Caption         =   "-"
      End
      Begin VB.Menu 复制路径 
         Caption         =   "复制文件路径"
      End
      Begin VB.Menu 批量重命名 
         Caption         =   "批量重命名"
      End
      Begin VB.Menu 分享音乐 
         Caption         =   "分享音乐"
      End
      Begin VB.Menu 默认程序打开 
         Caption         =   "使用系统默认程序打开"
      End
      Begin VB.Menu afhhjj 
         Caption         =   "-"
      End
      Begin VB.Menu 分割文件 
         Caption         =   "文件属性"
      End
   End
   Begin VB.Menu 顺序 
      Caption         =   "播放顺序"
      Begin VB.Menu 单曲循环 
         Caption         =   "单曲重复"
      End
      Begin VB.Menu 顺序播放 
         Caption         =   "顺序播放"
      End
      Begin VB.Menu 随机播放 
         Caption         =   "随机播放"
      End
      Begin VB.Menu 循环 
         Caption         =   "循环播放"
      End
   End
   Begin VB.Menu iM 
      Caption         =   "IM"
      Begin VB.Menu mnuStatusOnline 
         Caption         =   "在线"
      End
      Begin VB.Menu mnuStatusAway 
         Caption         =   "离开"
      End
      Begin VB.Menu mnuStatusDND 
         Caption         =   "不要打扰"
      End
      Begin VB.Menu mnuStatusInvisible 
         Caption         =   "隐身"
      End
      Begin VB.Menu UNLOGIN 
         Caption         =   "注销"
      End
      Begin VB.Menu SSAD 
         Caption         =   "-"
      End
      Begin VB.Menu 锁定 
         Caption         =   "锁定iCee"
      End
   End
   Begin VB.Menu mnuBuddy 
      Caption         =   "好友"
      Begin VB.Menu mnuBuddyMessage 
         Caption         =   "发送消息"
      End
      Begin VB.Menu mnuBuddyChat 
         Caption         =   "即时聊天"
      End
      Begin VB.Menu 远程协助 
         Caption         =   "远程协助"
      End
      Begin VB.Menu 视频聊天 
         Caption         =   "视频聊天"
      End
      Begin VB.Menu mnuBuddyFile 
         Caption         =   "传送文件"
      End
      Begin VB.Menu mnuBuddyInfo 
         Caption         =   "查看他的注册信息"
      End
      Begin VB.Menu ASWF 
         Caption         =   "-"
      End
      Begin VB.Menu 刷新好友列表 
         Caption         =   "刷新列表"
      End
      Begin VB.Menu DFGR 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuddyRemove 
         Caption         =   "删除好友"
      End
      Begin VB.Menu 举报 
         Caption         =   "举报该用户"
      End
      Begin VB.Menu 屏蔽 
         Caption         =   "屏蔽该用户"
      End
      Begin VB.Menu mnuBuddyIgnore 
         Caption         =   "黑名单管理"
      End
      Begin VB.Menu SDWW 
         Caption         =   "-"
      End
      Begin VB.Menu 聊天记录 
         Caption         =   "查看聊天记录"
      End
      Begin VB.Menu DAFE 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePassword 
         Caption         =   "修改登录密码"
      End
      Begin VB.Menu mnuFileChangeInfo 
         Caption         =   "修改个人信息"
      End
      Begin VB.Menu xsdf 
         Caption         =   "-"
      End
      Begin VB.Menu 反馈 
         Caption         =   "BUG反馈"
      End
   End
   Begin VB.Menu 特效 
      Caption         =   "涂鸦特效"
      Begin VB.Menu 锐化 
         Caption         =   "锐化图像"
      End
      Begin VB.Menu 模糊 
         Caption         =   "羽化图像"
      End
      Begin VB.Menu 噪音 
         Caption         =   "添加噪音"
      End
      Begin VB.Menu 镜像 
         Caption         =   "镜像特效"
      End
      Begin VB.Menu 灰度 
         Caption         =   "灰化图像"
      End
      Begin VB.Menu 反转 
         Caption         =   "反转颜色"
      End
      Begin VB.Menu 魔术 
         Caption         =   "魔术效果"
      End
      Begin VB.Menu 油画 
         Caption         =   "油画效果"
      End
      Begin VB.Menu 木刻 
         Caption         =   "木刻效果"
      End
      Begin VB.Menu 浮雕 
         Caption         =   "浮雕效果"
      End
      Begin VB.Menu 马赛克 
         Caption         =   "马赛克"
      End
      Begin VB.Menu 眩晕 
         Caption         =   "眩晕"
      End
      Begin VB.Menu 增加亮度 
         Caption         =   "增加亮度"
      End
      Begin VB.Menu 减少亮度 
         Caption         =   "减少亮度"
      End
      Begin VB.Menu 增加对比度 
         Caption         =   "增加对比度"
      End
      Begin VB.Menu DEES 
         Caption         =   "-"
      End
      Begin VB.Menu 水平 
         Caption         =   "水平翻转"
      End
      Begin VB.Menu 垂直 
         Caption         =   "垂直翻转"
      End
      Begin VB.Menu 双向 
         Caption         =   "双向翻转"
      End
      Begin VB.Menu SWRR 
         Caption         =   "-"
      End
      Begin VB.Menu 从剪切板粘贴 
         Caption         =   "从剪切板粘贴"
      End
      Begin VB.Menu 涂鸦剪切板 
         Caption         =   "复制到剪切板"
      End
      Begin VB.Menu SERSFFG 
         Caption         =   "-"
      End
      Begin VB.Menu 黑色边框 
         Caption         =   "添加黑色边框"
      End
      Begin VB.Menu 背景色蒙版 
         Caption         =   "背景颜色蒙版"
      End
      Begin VB.Menu FGGHH 
         Caption         =   "-"
      End
      Begin VB.Menu 去除图像 
         Caption         =   "清除图像"
      End
      Begin VB.Menu SDFGGSAA 
         Caption         =   "-"
      End
      Begin VB.Menu 打开图像 
         Caption         =   "打开图像"
      End
      Begin VB.Menu 保存 
         Caption         =   "保存图像"
      End
      Begin VB.Menu 与好友分享 
         Caption         =   "分享"
      End
      Begin VB.Menu DAFFF 
         Caption         =   "-"
      End
      Begin VB.Menu 设为头像 
         Caption         =   "设为头像"
      End
   End
   Begin VB.Menu 主机 
      Caption         =   "主机"
      Begin VB.Menu 尝试 
         Caption         =   "尝试将此机器作为服务器"
      End
      Begin VB.Menu SWFF 
         Caption         =   "-"
      End
      Begin VB.Menu 重新搜索 
         Caption         =   "重新搜索主机"
      End
      Begin VB.Menu 清空数据 
         Caption         =   "清空数据"
      End
   End
   Begin VB.Menu 下载任务 
      Caption         =   "下载任务"
      Begin VB.Menu 添加新任务 
         Caption         =   "添加新任务"
      End
      Begin VB.Menu sjkai 
         Caption         =   "-"
      End
      Begin VB.Menu 停止下载 
         Caption         =   "停止下载"
      End
      Begin VB.Menu SWGGSSAA 
         Caption         =   "-"
      End
      Begin VB.Menu 打开任务 
         Caption         =   "打开文件"
      End
      Begin VB.Menu 定位下载 
         Caption         =   "定位文件"
      End
      Begin VB.Menu WGJK 
         Caption         =   "-"
      End
      Begin VB.Menu 复制下载链接 
         Caption         =   "复制下载链接"
      End
      Begin VB.Menu 删除下载任务 
         Caption         =   "删除下载任务"
      End
      Begin VB.Menu 重新下载 
         Caption         =   "重新下载任务"
      End
      Begin VB.Menu sakkii 
         Caption         =   "-"
      End
      Begin VB.Menu 生成二维码 
         Caption         =   "生成链接二维码"
      End
      Begin VB.Menu LLIDI 
         Caption         =   "-"
      End
      Begin VB.Menu 清空下载 
         Caption         =   "清空下载列表"
      End
   End
   Begin VB.Menu 任务管理 
      Caption         =   "任务管理器"
      Begin VB.Menu 结束进程 
         Caption         =   "结束进程"
      End
      Begin VB.Menu 删除进程文件 
         Caption         =   "删除进程文件"
      End
      Begin VB.Menu SWFFW 
         Caption         =   "-"
      End
      Begin VB.Menu 进程文件 
         Caption         =   "进程文件"
      End
      Begin VB.Menu 进程属性 
         Caption         =   "进程属性"
      End
      Begin VB.Menu SJWJ 
         Caption         =   "-"
      End
      Begin VB.Menu 刷新进程 
         Caption         =   "刷新列表"
      End
   End
   Begin VB.Menu 歌词 
      Caption         =   "歌词"
      Begin VB.Menu 查看歌词 
         Caption         =   "查看歌词"
      End
      Begin VB.Menu 删除关联 
         Caption         =   "删除关联"
      End
      Begin VB.Menu 编辑歌词 
         Caption         =   "编辑歌词"
      End
      Begin VB.Menu 搜索歌词 
         Caption         =   "搜索歌词"
      End
      Begin VB.Menu SWWAGHGH 
         Caption         =   "-"
      End
      Begin VB.Menu 桌面歌词 
         Caption         =   "桌面歌词"
      End
   End
   Begin VB.Menu TCP 
      Caption         =   "TCP"
      Begin VB.Menu 断开TCP连接 
         Caption         =   "断开连接"
      End
   End
   Begin VB.Menu 我的收藏 
      Caption         =   "我的收藏"
      Begin VB.Menu 添加收藏 
         Caption         =   "添加到播放列表"
      End
      Begin VB.Menu 添加全部收藏 
         Caption         =   "添加全部"
      End
      Begin VB.Menu 删除收藏 
         Caption         =   "删除"
      End
      Begin VB.Menu 清空收藏 
         Caption         =   "清空收藏"
      End
   End
   Begin VB.Menu 文件管理 
      Caption         =   "文件管理"
      Begin VB.Menu 打开文件 
         Caption         =   "打开文件"
      End
      Begin VB.Menu SWFHKK 
         Caption         =   "-"
      End
      Begin VB.Menu 删除文件 
         Caption         =   "删除文件"
      End
      Begin VB.Menu 复制文件 
         Caption         =   "复制文件"
      End
      Begin VB.Menu 粘贴文件 
         Caption         =   "粘贴文件"
      End
      Begin VB.Menu SJJJJAC 
         Caption         =   "-"
      End
      Begin VB.Menu 重命名文件 
         Caption         =   "重命名文件"
      End
      Begin VB.Menu 批量重命名该文件夹下的文件 
         Caption         =   "批量重命名该文件夹下的文件"
      End
      Begin VB.Menu JIIS 
         Caption         =   "-"
      End
      Begin VB.Menu 分享该文件 
         Caption         =   "分享该文件"
      End
      Begin VB.Menu safrr 
         Caption         =   "-"
      End
      Begin VB.Menu 文件属性 
         Caption         =   "文件属性"
      End
      Begin VB.Menu 文件夹属性 
         Caption         =   "文件夹属性"
      End
      Begin VB.Menu AAAWWW 
         Caption         =   "-"
      End
      Begin VB.Menu 新建文件夹 
         Caption         =   "新建文件夹"
      End
   End
   Begin VB.Menu 系统托盘 
      Caption         =   "系统托盘"
      Begin VB.Menu 检查更新 
         Caption         =   "检查更新"
      End
      Begin VB.Menu 屏幕键盘 
         Caption         =   "屏幕键盘"
      End
      Begin VB.Menu 反馈问题 
         Caption         =   "反馈问题"
      End
      Begin VB.Menu 关于 
         Caption         =   "关于"
      End
      Begin VB.Menu QAASDFF 
         Caption         =   "-"
      End
      Begin VB.Menu 退粗 
         Caption         =   "退出"
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
Call HookForm(Me)  '活动窗体
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
Private Function fncGetInfo(lsPicName As String) As PICINFO '不使用控件获得图片大小
    Dim hBitmap As Long
    Dim res As Long
    Dim Bmp As BITMAP
    res = GetObject(LoadPicture(lsPicName).handle, Len(Bmp), Bmp) '取得BITMAP的结构
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
Call frmGraphic.打开(frmGraphic.Select_Pic)
Call frmGraphic.pic_turn
Case 5
If UCase(Right(frmGraphic.Select_Pic, 3)) = "PNG" Then Exit Sub
Call frmGraphic.打开(frmGraphic.Select_Pic)
Call frmGraphic.Pic_Talking
Case 6
If UCase(Right(frmGraphic.Select_Pic, 3)) = "PNG" Then Exit Sub
Call frmGraphic.打开(frmGraphic.Select_Pic)
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
'全局变量的初始化
g_clrFrame = &H8BA31F    '选择项目外框的颜色
g_clrBkgSelect = &H8BA31F    ' RGB(93, 80, 58) '选中项目背景的颜色
g_clrLeft = &H8BA31F '菜单左边的颜色

g_clrBkgNormal = vbWhite  ' RGB(253, 251, 250) '正常背景的颜色
g_clrTxtSelect = RGB(255, 255, 255) '选中文本的颜色
g_clrTxtNormal = RGB(0, 0, 0) '正常文本的颜色
g_clrSep = RGB(209, 209, 209) '分割线的颜色

hMainMenu = GetMenu(Me.hwnd) '得到窗体顶级菜单句柄
hSubMenu = GetSubMenu(hMainMenu, 0) '得到文件菜单的句柄
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
'保存菜单项的信息
RegisterMenu hSubMenu, 0, "全选", 180, 20
RegisterMenu hSubMenu, 1, "复制", 180, 20
RegisterMenu hSubMenu, 2, "剪切", 180, 20
RegisterMenu hSubMenu, 3, "粘贴", 180, 20
RegisterMenu hSubMenu, 4, "删除", 180, 20

RegisterMenu Hlocal, 0, "预览图片", 180, 20
RegisterMenu Hlocal, 1, "", 180, 5
RegisterMenu Hlocal, 2, "作为头像", 180, 20
RegisterMenu Hlocal, 3, "", 180, 5
RegisterMenu Hlocal, 4, "旋转图片", 180, 20
RegisterMenu Hlocal, 5, "增加滤镜", 180, 20
RegisterMenu Hlocal, 6, "调整曝光度", 180, 20
RegisterMenu Hlocal, 7, "生成字符画", 180, 20
RegisterMenu Hlocal, 8, "涂鸦", 180, 20
RegisterMenu Hlocal, 9, "", 180, 5
RegisterMenu Hlocal, 10, "与好友分享", 180, 20
RegisterMenu Hlocal, 11, "", 180, 5
RegisterMenu Hlocal, 12, "图片详情", 180, 20

RegisterMenu hFtp, 0, "打开文件", 180, 20
RegisterMenu hFtp, 1, "添加文件", 180, 20
RegisterMenu hFtp, 2, "播放文件夹", 180, 20
RegisterMenu hFtp, 3, "添加文件夹", 180, 20, PIC(5)
RegisterMenu hFtp, 4, "", 180, 5
RegisterMenu hFtp, 5, "我的收藏", 180, 20
RegisterMenu hFtp, 6, "打开音乐窗", 180, 20
RegisterMenu hFtp, 7, "", 180, 5
RegisterMenu hFtp, 8, "导入播放列表", 180, 20
RegisterMenu hFtp, 9, "导出播放列表", 180, 20
RegisterMenu hFtp, 10, "", 180, 5
RegisterMenu hFtp, 11, "云同步列表", 180, 20, PIC(14)

RegisterMenu hTxt, 0, "播放", 160, 20
RegisterMenu hTxt, 1, "添加", 160, 20
RegisterMenu hTxt, 2, "下载", 160, 20
RegisterMenu hTxt, 3, "", 160, 5
RegisterMenu hTxt, 4, "播放全部", 160, 20

RegisterMenu hPic, 0, "播放", 160, 20
RegisterMenu hPic, 1, "从列表删除", 160, 20
RegisterMenu hPic, 2, "从磁盘删除", 160, 20
RegisterMenu hPic, 3, "", 160, 5
RegisterMenu hPic, 4, "去除重复", 160, 20
RegisterMenu hPic, 5, "删除无效", 160, 20
RegisterMenu hPic, 6, "刷新列表", 160, 20
RegisterMenu hPic, 7, "", 160, 5
RegisterMenu hPic, 8, "更多信息", 160, 20
RegisterMenu hPic, 9, "打开位置", 160, 20
RegisterMenu hPic, 10, "", 160, 5
RegisterMenu hPic, 11, "复制文件源", 160, 20
RegisterMenu hPic, 12, "批量重命名", 160, 20
RegisterMenu hPic, 13, "分享音乐", 160, 20
RegisterMenu hPic, 14, "外部打开", 160, 20
RegisterMenu hPic, 15, "", 160, 5
RegisterMenu hPic, 16, "文件属性", 160, 20

RegisterMenu Hpla, 0, "单曲重复", 160, 20
RegisterMenu Hpla, 1, "顺序播放", 160, 20
RegisterMenu Hpla, 2, "随机播放", 160, 20
RegisterMenu Hpla, 3, "循环播放", 160, 20

RegisterMenu Hlst, 0, "在线", 160, 20
RegisterMenu Hlst, 1, "离开", 160, 20
RegisterMenu Hlst, 2, "不要打扰", 160, 20
RegisterMenu Hlst, 3, "隐身", 160, 20
RegisterMenu Hlst, 4, "注销登陆", 160, 20
RegisterMenu Hlst, 5, "", 160, 5
RegisterMenu Hlst, 6, "锁定程序", 160, 20

RegisterMenu hFile, 0, "发送消息", 160, 20
RegisterMenu hFile, 1, "窗口聊天", 160, 20
RegisterMenu hFile, 2, "远程协助", 160, 20
RegisterMenu hFile, 3, "视频聊天", 160, 20
RegisterMenu hFile, 4, "发送文件", 160, 20
RegisterMenu hFile, 5, "Ta的信息", 160, 20
RegisterMenu hFile, 6, "", 160, 5
RegisterMenu hFile, 7, "刷新列表", 160, 20
RegisterMenu hFile, 8, "", 160, 5
RegisterMenu hFile, 9, "删除好友", 160, 20
RegisterMenu hFile, 10, "举报Ta", 160, 20
RegisterMenu hFile, 11, "屏蔽Ta", 160, 20
RegisterMenu hFile, 12, "黑名单管理", 160, 20
RegisterMenu hFile, 13, "", 160, 5
RegisterMenu hFile, 14, "聊天记录", 160, 20
RegisterMenu hFile, 15, "", 160, 5
RegisterMenu hFile, 16, "修改密码", 160, 20
RegisterMenu hFile, 17, "修改信息", 160, 20
RegisterMenu hFile, 18, "", 160, 5
RegisterMenu hFile, 19, "反馈问题", 160, 20

RegisterMenu HTY, 0, "锐化图像", 160, 20
RegisterMenu HTY, 1, "羽化图像", 160, 20
RegisterMenu HTY, 2, "添加噪点", 160, 20
RegisterMenu HTY, 3, "对称特效", 160, 20
RegisterMenu HTY, 4, "灰化图像", 160, 20
RegisterMenu HTY, 5, "反转色彩", 160, 20
RegisterMenu HTY, 6, "魔术色彩", 160, 20
RegisterMenu HTY, 7, "绚丽油画", 160, 20
RegisterMenu HTY, 8, "黑白分明", 160, 20
RegisterMenu HTY, 9, "古典浮雕", 160, 20
RegisterMenu HTY, 10, "马赛克效果", 160, 20
RegisterMenu HTY, 11, "眩晕效果(新)", 160, 20
RegisterMenu HTY, 12, "增加亮度", 160, 20
RegisterMenu HTY, 13, "减少亮度", 160, 20
RegisterMenu HTY, 14, "增加对比度", 160, 20
RegisterMenu HTY, 15, "", 160, 5
RegisterMenu HTY, 16, "水平翻转", 160, 20
RegisterMenu HTY, 17, "垂直翻转", 160, 20
RegisterMenu HTY, 18, "双向翻转", 160, 20
RegisterMenu HTY, 19, "", 160, 5
RegisterMenu HTY, 20, "从剪切板粘贴", 160, 20
RegisterMenu HTY, 21, "复制到剪切板", 160, 20
RegisterMenu HTY, 22, "", 160, 5
RegisterMenu HTY, 23, "背景色边框", 160, 20
RegisterMenu HTY, 24, "背景色蒙版", 160, 20
RegisterMenu HTY, 25, "", 160, 5
RegisterMenu HTY, 26, "清空画板", 160, 20
RegisterMenu HTY, 27, "", 160, 5
RegisterMenu HTY, 28, "打开图像", 160, 20
RegisterMenu HTY, 29, "导出涂鸦", 160, 20
RegisterMenu HTY, 30, "与好友分享", 160, 20
RegisterMenu HTY, 31, "", 160, 5
RegisterMenu HTY, 32, "设为头像", 160, 20

RegisterMenu HNET, 0, "作为服务器", 160, 20
RegisterMenu HNET, 1, "", 160, 5
RegisterMenu HNET, 2, "重新扫描", 160, 20
RegisterMenu HNET, 3, "清空结果", 160, 20

RegisterMenu HWQ, 0, "添加新任务", 160, 20
RegisterMenu HWQ, 1, "", 160, 5
RegisterMenu HWQ, 2, "停止下载", 160, 20
RegisterMenu HWQ, 3, "", 160, 5
RegisterMenu HWQ, 4, "打开文件", 160, 20
RegisterMenu HWQ, 5, "定位文件", 160, 20
RegisterMenu HWQ, 6, "", 160, 5
RegisterMenu HWQ, 7, "复制下载链接", 160, 20
RegisterMenu HWQ, 8, "删除下载任务", 160, 20
RegisterMenu HWQ, 9, "重新下载任务", 160, 20
RegisterMenu HWQ, 10, "", 160, 5
RegisterMenu HWQ, 11, "生成二维码", 160, 20
RegisterMenu HWQ, 12, "", 160, 5
RegisterMenu HWQ, 13, "清空列表", 160, 20

RegisterMenu HFIVE, 0, "结束进程", 160, 20
RegisterMenu HFIVE, 1, "删除进程文件", 160, 20
RegisterMenu HFIVE, 2, "", 160, 5
RegisterMenu HFIVE, 3, "定位进程", 160, 20
RegisterMenu HFIVE, 4, "进程属性", 160, 20
RegisterMenu HFIVE, 5, "", 160, 5
RegisterMenu HFIVE, 6, "刷新列表", 160, 20

RegisterMenu HLOOK, 0, "查看歌词", 160, 20
RegisterMenu HLOOK, 1, "删除歌词关联", 160, 20
RegisterMenu HLOOK, 2, "编辑歌词", 160, 20
RegisterMenu HLOOK, 3, "搜索歌词", 160, 20
RegisterMenu HLOOK, 4, "", 160, 5
RegisterMenu HLOOK, 5, "桌面歌词", 160, 20

RegisterMenu hZor, 0, "断开连接", 160, 20

RegisterMenu hFr, 0, "添加选中到当前列表", 160, 20
RegisterMenu hFr, 1, "添加全部到当前列表", 160, 20
RegisterMenu hFr, 2, "删除收藏", 160, 20
RegisterMenu hFr, 3, "清空收藏", 160, 20

RegisterMenu hcol, 0, "打开文件", 160, 20
RegisterMenu hcol, 1, "", 160, 5
RegisterMenu hcol, 2, "删除文件", 160, 20
RegisterMenu hcol, 3, "复制文件", 160, 20
RegisterMenu hcol, 4, "粘贴文件", 160, 20
RegisterMenu hcol, 5, "", 160, 5
RegisterMenu hcol, 6, "重命名文件", 160, 20
RegisterMenu hcol, 7, "重命名该文件夹下的文件", 160, 20
RegisterMenu hcol, 8, "", 160, 5
RegisterMenu hcol, 9, "分享该文件", 160, 20
RegisterMenu hcol, 10, "", 160, 5  '初始化设置菜单栏的背景色,注意这个颜色最好跟g_clrBkgNormal一样,要不效果不好
RegisterMenu hcol, 11, "文件属性", 160, 20
RegisterMenu hcol, 12, "文件夹属性", 160, 20
RegisterMenu hcol, 13, "", 160, 5
RegisterMenu hcol, 14, "新建文件夹", 160, 20

RegisterMenu HABOUT, 0, "检查更新", 160, 20
RegisterMenu HABOUT, 1, "屏幕键盘", 160, 20
RegisterMenu HABOUT, 2, "反馈问题", 160, 20
RegisterMenu HABOUT, 3, "关于ICEE", 160, 20
RegisterMenu HABOUT, 4, "", 160, 5
RegisterMenu HABOUT, 5, "退出", 160, 20
Call SetMenuBar(Me, &H8BA31F) 'RGB(224, 234, 240))
End Sub
Private Sub mnuBuddyRemove_Click()
frmma.移除好友
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
Call 注销
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
SB = WB.Document.links.Item(I).innerText 'SB是页面中所有超链接文字
s = WB.Document.links.Item(I) 'S是页面中所有超链接
If Left(UCase(SB), 7) = "[TODAY]" Then LSTLINK.AddItem SB & "|" & s
End If
Next I
If LSTLINK.ListCount = 0 Then Exit Sub
LBLINK.Caption = Replace(Split(LSTLINK.List(0), "|")(0), "[TODAY]", "")
'小备注:SPLIT(目标串,查找串)(位置) 例如:split("ABCD|abcd")(1) 返回值就是"abcd"
If LBLINK.Caption = "" Then frmma.lbthing.Caption = "欢迎使用1.24全新版本,更多精彩等你发现" Else frmma.lbthing.Caption = LBLINK.Caption
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
Call SHOWWRONG("发生错误:" & Number & vbCrLf & Description, 2)
End Sub


Private Sub 编辑歌词_Click()
FRMLRC.L_EDIT
End Sub

Private Sub 播放当前列表_Click()
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

Private Sub 播放网络音乐_Click()
On Error Resume Next
frmma.PLIST.AddItem FrmNetMusic.M_N, FrmNetMusic.A_N, FrmNetMusic.Will_DL, 0
frmma.Wm.URL = FrmNetMusic.Will_DL
frmma.PLIST.ListIndex = PLIST.ListCount - 1
End Sub

Private Sub 查看歌词_Click()
FRMLRC.L_VIEW
End Sub

Private Sub 从剪切板粘贴_Click()
On Error Resume Next
PICCLIP.PICTURE = Clipboard.GetData(2)
Call SavePicture(PICCLIP.PICTURE, App.Path & "\thumbs\TH_CLIP.BMP")
Call FRMBOARD.OpenFile(App.Path & "\thumbs\TH_CLIP.BMP")
End Sub

Private Sub 打开URL_Click()
Call frmma.MUSICBOX
End Sub

Private Sub 打开任务_Click()
On Error Resume Next
If Right(FRMDOWN.LVIEW.SelectedItem.SubItems(2), 1) = "%" Then Exit Sub
If FRMDOWN.LVIEW.SelectedItem.SubItems(4) = "" Then Exit Sub
Call SYSTEMOPEN(Dpath & FRMDOWN.LVIEW.SelectedItem.Text)
End Sub

Private Sub 打开文件_Click()
FRMEX.OPEN_CLICK
End Sub

Private Sub 定位下载_Click()
On Error Resume Next
If FRMDOWN.LVIEW.SelectedItem.SubItems(4) = "" Then Exit Sub
Shell "explorer.exe /select," & Dpath & FRMDOWN.LVIEW.SelectedItem.Text, vbNormalFocus
End Sub

Private Sub 断开TCP连接_Click()
FRMEND.KIILTCP
End Sub

Private Sub 反馈问题_Click()
ShellExecute 0&, vbNullString, "http://tieba.baidu.com/f?ie=utf-8&kw=icee", vbNullString, vbNullString, 0 '调用ie
End Sub

Private Sub 分享该文件_Click()
frmma.SHAREIT (FRMEX.Txt_Address.Text & "\" & FRMEX.ListView1.SelectedItem.Text)
End Sub

Private Sub 分享音乐_Click()
DefCOM = 1
Call frmma.SHAREIT(frmma.PLIST.URL(frmma.PLIST.ListIndex))
End Sub

Private Sub 浮雕_Click()
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

Private Sub 复制文件_Click()
FRMEX.Copy_Click
End Sub

Private Sub 复制下载链接_Click()
On Error Resume Next
Clipboard.SetText (FRMDOWN.LVIEW.SelectedItem.SubItems(7))
End Sub

Private Sub 关于_Click()
FrmWhatNew.Show
End Sub

Private Sub 减少亮度_Click()
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

Private Sub 检查更新_Click()
FRMUP.Show
End Sub

Private Sub 结束进程_Click()
    Call FRMEND.ENDIT
End Sub

Private Sub 进程文件_Click()
FRMEND.FOLDERPRO
End Sub

Private Sub 进程属性_Click()
FRMEND.PAPERPRO
End Sub

Private Sub 魔术_Click()
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

Private Sub 木刻_Click()
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
Private Sub 保存_Click()
Call frmma.保存一下(FRMBOARD.PICTY)
End Sub
Private Sub 播放选中_Click()
frmma.播放歌曲
End Sub

Private Sub 尝试_Click()
If frmma.lstRes.List(frmma.lstRes.ListIndex) <> "" Then
frmma.Text3.Text = frmma.lstRes.Text
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">尝试将" & frmma.lstRes.List(frmma.lstRes.ListIndex) & "作为服务器"
frmma.PICNET.Visible = False
Call frmma.LOCKSAFE
Call frmma.SUBDRAWIM
frmma.LBITEM(2).Caption = "请登陆"
frmma.IMJ.Visible = False
frmma.IMG_NT.Visible = True
End If
End Sub
Private Sub 垂直_Click()
Call FlipImage(FRMBOARD.PICTY, 1)
End Sub

Private Sub 打开播放列表_Click()
Dim sFile As String
sFile = ShowOpen(Me.hwnd, "播放列表文件 M3u" & Chr(0) & "*.m3u", "打开播放列表")
If Dir$(sFile) <> vbNullString And sFile <> "" Then Call frmma.Playlist(sFile)
End Sub
Sub 打开光驱()
'以下是打开CD -ROM的过程代码:
retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub

Private Sub 打开媒体_Click()
Lmenu (0)
End Sub

Private Sub 打开图像_Click()
Dim sFile As String
sFile = ShowOpen(Me.hwnd, "BMP文件" & Chr(0) & "*.Bmp" _
& Chr(0) & "JEPG文件" & Chr(0) & "*.jpg;*.jepg" _
& Chr(0) & "Gif" & Chr(0) & "*.gif" _
& Chr(0) & "Png" & Chr(0) & "*.png", "打开图片")
Call FRMBOARD.OpenFile(sFile)
End Sub

Private Sub 打开文件夹_Click()
Call OpenDir
End Sub
Private Sub 单曲循环_Click()
LOLIPOP = 1
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "单曲循环"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\DQ_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换音乐播放模式为单曲循环"
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
End Sub

Private Sub 导出_Click()
frmma.导出列表
End Sub
Private Sub 反馈_Click()
Call Report
End Sub
Private Sub 分割文件_Click()
On Error Resume Next
Dim r As Long
With frmma
Dim filename As String
filename = .PLIST.URL(.PLIST.ListIndex)
r = ShowProperties(filename, frmma.hwnd)
End With
If r <= 32 Then Call SHOWWRONG("对不起,查看文件属性失败(可能的原因是权限不足,或者歌曲来自网络)", 2)

End Sub

Private Sub 复制路径_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText (frmma.PLIST.URL(frmma.PLIST.ListIndex))
End Sub
Private Sub 复制文本_Click()
Call 复制
End Sub

Sub 关闭光驱()
'关闭CD -ROM用以下代码:
retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub
Private Sub 回收站_Click()
Call frmma.物理删除歌曲
End Sub
Private Sub 剪切文本_Click()
Call 剪切
End Sub
Private Sub 聊天记录_Click()
If frmma.Left > FRMHIS.Width Then
FRMHIS.Move frmma.Left - FRMHIS.Width, frmma.Top
Else
FRMHIS.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMHIS.Show
End Sub
Private Sub 默认程序打开_Click()
On Error Resume Next
Call SYSTEMOPEN(frmma.PLIST.URL(frmma.PLIST.ListIndex))
End Sub

Private Sub 批量重命名_Click()
If frmma.Left > FORMNAME.Width Then
FORMNAME.Move frmma.Left - FORMNAME.Width, frmma.Top
Else
FORMNAME.Move frmma.Left + frmma.Width, frmma.Top
End If
FORMNAME.Show
FORMNAME.txtPath.Text = MMAIN.GetPathFromFileName(frmma.PLIST.URL(frmma.PLIST.ListIndex), "\")
End Sub

Private Sub 批量重命名该文件夹下的文件_Click()
If frmma.Left > FORMNAME.Width Then
FORMNAME.Move frmma.Left - FORMNAME.Width, frmma.Top
Else
FORMNAME.Move frmma.Left + frmma.Width, frmma.Top
End If

FORMNAME.Show
FORMNAME.txtPath = FRMEX.Txt_Address.Text
End Sub

Private Sub 屏蔽_Click()
frmma.屏蔽此用户
End Sub

Private Sub 屏幕键盘_Click()
FRMKEYBOARD.Show
End Sub

Private Sub 清空收藏_Click()
Call FRMFAV.CLEAR_FAV
End Sub

Private Sub 清空数据_Click()
frmma.lstRes.Clear
End Sub

Private Sub 清空下载_Click()
FRMDOWN.LVIEW.ListItems.Clear
If FRMDOWN.LVIEW.ListItems.Count > 0 Then FRMDOWN.LVIEW.Visible = True Else FRMDOWN.LVIEW.Visible = False
Call FRMDOWN.SAVELIST
Call FRMDOWN.LoadList
End Sub

Private Sub 去除图像_Click()
Set FRMBOARD.PT.PICTURE = Nothing
Set FRMBOARD.PICTY.PICTURE = Nothing
FRMBOARD.PT.BackColor = FRMBOARD.PB.BackColor  '涂鸦画板颜色
FRMBOARD.PICTY.BackColor = FRMBOARD.PB.BackColor
End Sub

Private Sub 去重_Click()
Call frmma.去除重复
End Sub

Private Sub 全选文本_Click()
Call 全选
End Sub

Private Sub 删除关联_Click()
FRMLRC.L_DELETE
Call FrmNetMusic.L_LRC.ClearLrc
FrmNetMusic.L_LRC.Visible = False
If D_L_SHOW = True Then FrmNetMusic.cDeskLrc.ShowText " ICEE音乐,音乐您的生活"
End Sub

Private Sub 删除进程文件_Click()
FRMEND.DELPRO
End Sub

Private Sub 删除收藏_Click()
On Error Resume Next
Call FRMFAV.REMOVE_ITEM(FRMFAV.LFAV.SelectedItem.Text)
End Sub

Private Sub 删除文本_Click()
Call 删除文字
End Sub

Private Sub 删除文件_Click()
FRMEX.Del_Click
End Sub

Private Sub 删除无效_Click()
Call DEL_NONE
End Sub

Private Sub 删除下载任务_Click()
If FRMDOWN.LVIEW.ListItems.Count = 0 Then Exit Sub
FRMDOWN.LVIEW.ListItems.REMOVE (FRMDOWN.LVIEW.SelectedItem.Index)
If FRMDOWN.LVIEW.ListItems.Count > 0 Then FRMDOWN.LVIEW.Visible = True Else FRMDOWN.LVIEW.Visible = False
Call FRMDOWN.SAVELIST
Call FRMDOWN.LoadList
End Sub

Private Sub 删除选中_Click()
Call Lmenu(2)
End Sub

Private Sub 设为头像_Click()
Dim filea As String
filea = App.Path & "\THUMBS\H_Thumbs.Bmp"
Call SavePicture(FRMBOARD.PICTY.image, filea)
FRMHEAD.Show
Call FRMHEAD.OpenFile(filea)

Kill (filea)
End Sub

Private Sub 生成二维码_Click()
FRMDOWN.二维码
End Sub

Private Sub 收藏夹_Click()
If frmma.Left > FRMFAV.Width Then
FRMFAV.Move frmma.Left - FRMFAV.Width, frmma.Top
Else
FRMFAV.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMFAV.Show
End Sub

Private Sub 刷新好友列表_Click()
Dim I As Integer
With frmma
For I = 1 To .TreeView1.Nodes.Count
.Winsock1.SendData ".getstatus " & .TreeView1.Nodes(I).Key
Next
End With
End Sub

Private Sub 刷新进程_Click()
Call FRMEND.ListProcess
End Sub

Private Sub 刷新列表_Click()
frmma.PLIST.Refresh
End Sub
Private Sub 双向_Click()
Call FlipImage(FRMBOARD.PICTY, 2)
End Sub

Private Sub 水平_Click()
Call FlipImage(FRMBOARD.PICTY, 0)
End Sub

Private Sub 顺序播放_Click()
LOLIPOP = 3
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "顺序播放"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SX_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换音乐播放模式为顺寻播放"
End Sub

Private Sub 搜索歌词_Click()
FRMLRC.MOVEME
FRMLRC.Show
End Sub

Private Sub 随机播放_Click()
LOLIPOP = 0
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "随机播放"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\SJ_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换音乐播放模式为随机播放"
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
End Sub

Private Sub 锁定_Click()
Call LOCKME
End Sub
Sub LOCKME()
On Error Resume Next
With frmma
.RUNSAFE
.TXTPOUP.SetFocus
If frmma.Winsock1.State = 7 Then MYSTATUS = 2
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">ICEE被锁定成功"
.DRAWLOCK
IS_LOCK = True
Frmm.Hide
End With
End Sub
Private Sub 添加单个_Click()
Lmenu (4)
End Sub

Private Sub 添加全部收藏_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To FRMFAV.LFAV.ListItems.Count
If FRMFAV.LFAV.ListItems(I).Text <> "" Then frmma.PLIST.AddItem (FRMFAV.LFAV.ListItems(I).Text), "", FRMFAV.LFAV.ListItems(I).SubItems(I)
Next
End Sub

Private Sub 添加收藏_Click()
On Error Resume Next
frmma.PLIST.AddItem FRMFAV.LFAV.SelectedItem.Text, "", FRMFAV.LFAV.SelectedItem.SubItems(1)
End Sub

Private Sub 添加文件夹_Click()
Call AddDir
End Sub

Private Sub 添加新任务_Click()
Frmadd.Show
End Sub
Private Sub 添加至列表_Click()
frmma.PLIST.AddItem FrmNetMusic.M_N, FrmNetMusic.A_N, FrmNetMusic.Will_DL, 0
End Sub

Private Sub 停止下载_Click()
Call FRMDOWN.停止下载
End Sub

Private Sub 涂鸦剪切板_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetData FRMBOARD.PICTY.image
End Sub

Private Sub 退粗_Click()
Unload frmma
End Sub

Private Sub 位置_Click()
On Error Resume Next
If Dir(frmma.PLIST.URL(frmma.PLIST.ListIndex)) = "" Then Exit Sub
Shell "explorer.exe /select," & frmma.PLIST.URL(frmma.PLIST.ListIndex), vbNormalFocus
End Sub

Private Sub 文件夹属性_Click()
FRMEX.Attribute_Click
End Sub

Private Sub 文件属性_Click()
FRMEX.PropsShow (FRMEX.Txt_Address.Text & "\" & FRMEX.ListView1.SelectedItem.Text)
End Sub

Private Sub 下载_Click()
Call DoFileDownload(StrConv(FrmNetMusic.Will_DL, vbUnicode))
End Sub

Private Sub 新建文件夹_Click()
FRMEX.NewFolder_Click
End Sub

Private Sub 眩晕_Click()
Call 眩晕图像(FRMBOARD.PICTY)
Call SavePicture(FRMBOARD.PICTY.image, App.Path & "\THUMBS\THUMBS.BMP")
FRMBOARD.PT.PICTURE = LoadPicture(App.Path & "\THUMBS\THUMBS.BMP")
Call 眩晕图像(FRMBOARD.PT)
Call SavePicture(FRMBOARD.PT.image, App.Path & "\THUMBS\THUMBS.BMP")
FRMBOARD.PT.PICTURE = LoadPicture(App.Path & "\THUMBS\THUMBS.BMP")
FRMBOARD.PICTY.PaintPicture FRMBOARD.PT.image, 0, 0, FRMBOARD.PT.ScaleWidth, FRMBOARD.PT.ScaleHeight
Set FRMBOARD.PT.PICTURE = Nothing
Call PictureBoxSaveJPG(FRMBOARD.PICTY.image, App.Path & "\MEDIA\Paint\AutoSave.JPG", 100)

End Sub

Private Sub 循环_Click()
LOLIPOP = 2
frmma.PZOR.Cls
frmma.PZOR.ToolTipText = "列表循环"
LES = BitBlt(frmma.PZOR.hdc, 0, 0, frmma.PZOR.Width, frmma.PZOR.Height, frmma.PP.hdc, frmma.PZOR.Left, frmma.PZOR.Top, &HCC0020)
Call PaintPng(App.Path & "\SKIN\XH_N.PNG", frmma.PZOR.hdc, 0, 0)
frmma.PZOR.Refresh
If IS_MINI_LIST = True Then Call FRMLIST.REZOR
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换音乐播放模式为列表循环"
End Sub

Private Sub 音乐属性_Click()
On Error Resume Next
Call FRMMIN.SeeIt(frmma.PLIST.URL(frmma.PLIST.ListIndex))
If frmma.Left > FRMMIN.Width Then
FRMMIN.Move frmma.Left - FRMMIN.Width, frmma.Top
Else
FRMMIN.Move frmma.Left + frmma.Width, frmma.Top
End If
FRMMIN.Show
End Sub

Private Sub 油画_Click()
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

Private Sub 与好友分享_Click()
Dim Tfile As String
Tfile = App.Path & "\THUMBS\THUMBS.Bmp"
DefCOM = 0
Call SavePicture(FRMBOARD.PICTY.image, Tfile)
Call frmma.SHAREIT(Tfile)
End Sub
Private Sub 云同步_Click()
Dim JSI As Integer
JSI = GetSetting("ICEE", "Winsock", "Connect", 0)
If JSI = 0 Then
Call SHOWWRONG("请先登录服务器!", 2)
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">尝试音乐列表云同步失败(未登录服务器)"
Else
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">尝试音乐列表云同步成功"
End If
End Sub

Private Sub 增加对比度_Click()
Call UPDB(FRMBOARD.PICTY)
End Sub

Private Sub 增加亮度_Click()
Call UPLD(FRMBOARD.PICTY)
End Sub

Private Sub 粘贴文本_Click()
Call 粘贴
End Sub
Sub OpenDir()
Dim sDir As String, a As Integer
With frmma
sDir = BrowseFolder("打开音乐文件夹", frmma)
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
sDir = BrowseFolder("添加文件夹", frmma)
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
frmma.好友聊天
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
.TXTSER.Text = "请输入对方ID"
.PICIG.Visible = True
.IMJ.Visible = True
.RUNSAFE
.LA(1).Caption = "黑名单管理"
.LSTBOX.Selected(0) = True
End With
End Sub
Private Sub mnuBuddyInfo_Click()
frmma.PICIM.BackColor = PTCO.POINT(0, 0)
frmma.好友信息
End Sub

Private Sub mnuBuddyMessage_Click()
frmma.即时聊天
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
.LA(1).Caption = "修改密码"
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
.LA(1).Caption = "BUG提交"
.TXTSER.Text = "请输入对方ID"
.RUNSAFE
.PICIM.BackColor = PTCO.POINT(0, 0)
.TXTBUG.SetFocus
End With
End Sub

Sub 注销()
Call frmma.初始化
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
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "目前处于离开状态"
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换状态为离开"
SetTrayIcon Away.PICTURE
MYSTATUS = 1  '1为离开
frmma.DRAWFACE
End Sub

Private Sub mnuStatusDND_Click()
mnuStatusOnline.Checked = False
mnuStatusAway.Checked = False
mnuStatusDND.Checked = True
mnuStatusInvisible.Checked = False
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "目前处于忙碌状态"
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换状态为忙碌"
frmma.Winsock1.SendData ".status DND"
SetTrayIcon BusyNow.PICTURE
MYSTATUS = 2  '2为忙碌
frmma.DRAWFACE
End Sub


Private Sub mnuStatusInvisible_Click()
mnuStatusOnline.Checked = False
mnuStatusAway.Checked = False
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = True
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "目前处于隐身状态"
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换状态为隐身"
frmma.Winsock1.SendData ".status Invisible"
SetTrayIcon HideNow.PICTURE
MYSTATUS = 3  '3为隐身
frmma.DRAWFACE
End Sub


Private Sub mnuStatusOnline_Click()
mnuStatusOnline.Checked = True
mnuStatusAway.Checked = False
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = False
SetTrayTip "ICEE-" & frmma.Text1.Text & vbCrLf & "目前处于在线状态"
frmma.Winsock1.SendData ".status ONLINE"
SetTrayIcon ONLINE.PICTURE
FRMHIS.TXTSYS.Text = FRMHIS.TXTSYS.Text & vbCrLf & Now & ">更换状态为在线"
MYSTATUS = 0 '0为在线
frmma.DRAWFACE
End Sub

Private Sub 粘贴文件_Click()
FRMEX.Plaster_Click
End Sub

Private Sub 重命名文件_Click()
FRMEX.ReName_Click
End Sub

Private Sub 重新搜索_Click()
Call frmma.SERCHNET
End Sub
Private Sub 锐化_Click()
Call Sharpen(FRMBOARD.PICTY, 1)
End Sub
Private Sub 模糊_Click()
Call BlurImage(FRMBOARD.PICTY)
End Sub
Private Sub 噪音_Click()
Call Noise(FRMBOARD.PICTY, 20)
End Sub
Private Sub 镜像_Click()
Call Mirror(FRMBOARD.PICTY)
End Sub
Private Sub 灰度_Click()
Call GrayImage(FRMBOARD.PICTY)
End Sub
Private Sub 反转_Click()
Call InvertImage(FRMBOARD.PICTY)
End Sub
Private Sub 马赛克_Click()
Call MASAK(FRMBOARD.PICTY)
End Sub
Private Sub 背景色蒙版_Click()
On Error Resume Next
FRMBOARD.PO(0).ScaleMode = 1
Set FRMBOARD.PICTY.PICTURE = FRMBOARD.PICTY.image
Call ShadePicture(FRMBOARD.PICTY, FRMBOARD.PICTY, FRMBOARD.PB.BackColor, 5)
FRMBOARD.PO(0).ScaleMode = 3
End Sub

Private Sub 黑色边框_Click()
Call StrokeImage(FRMBOARD.PICTY, 15, FRMBOARD.PB.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 13, FRMBOARD.PF.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 10, FRMBOARD.PB.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 8, FRMBOARD.PF.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 5, FRMBOARD.PB.BackColor)
Call StrokeImage(FRMBOARD.PICTY, 1, FRMBOARD.PF.BackColor)
End Sub

Private Sub 桌面歌词_Click()
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
